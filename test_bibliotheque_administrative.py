import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import ModuleType, SimpleNamespace
import sys
from unittest.mock import Mock, patch
from zipfile import ZipFile

from bibliotheque_administrative import (DOSSIERS_CATEGORIES, classer_bibliotheque_cinq_dossiers,
                                         creer_bibliotheque_administrative,
                                         copier_documents_administratifs, importer_dans_bibliotheque,
                                         ranger_copies_secours, DOSSIER_COPIES_SECOURS)
from main import HorizonChantierApp, mettre_fichier_corbeille


class BibliothequeAdministrativeTests(unittest.TestCase):
    def test_copies_secours_capture_conservees_hors_documents_courants(self):
        with TemporaryDirectory() as temp:
            for chantier in ('004_Mariemont', '006_Nouveau chantier déplacé'):
                dossier = Path(temp) / chantier / 'Administratif'
                noms = ('Liste de références JT BATI Mariemont',
                        'Coordonnée concernant la fourniture des copeaux JT BATI Mariemont')
                for nom in noms:
                    (dossier / nom).mkdir(parents=True)
                    (dossier / f'{nom}.docx').write_bytes(b'original courant')
                    (dossier / nom / f'Copie de secours de {nom}.docx').write_bytes(b'secours')
                (dossier / 'Liste copie.docx').write_bytes(b'copie volontaire')
                self.assertEqual(ranger_copies_secours(dossier), 2)
                self.assertEqual(ranger_copies_secours(dossier), 0)
                for nom in noms:
                    relatif = Path(nom) / f'Copie de secours de {nom}.docx'
                    self.assertFalse((dossier / relatif).exists())
                    self.assertEqual((dossier / DOSSIER_COPIES_SECOURS / relatif).read_bytes(), b'secours')
                    self.assertEqual((dossier / f'{nom}.docx').read_bytes(), b'original courant')
                self.assertEqual((dossier / 'Liste copie.docx').read_bytes(), b'copie volontaire')

    def test_nouvelles_versions_secours_et_conflits_sans_ecrasement(self):
        with TemporaryDirectory() as temp:
            dossier = Path(temp)
            nom = 'Copie de secours de Liste.docx'
            source = dossier / nom
            for contenu in (b'version 1', b'version 2', b'version 2'):
                source.write_bytes(contenu)
                self.assertEqual(ranger_copies_secours(dossier), 1)
            self.assertEqual(sorted(p.read_bytes() for p in (dossier / DOSSIER_COPIES_SECOURS).iterdir()),
                             [b'version 1', b'version 2'])
            (dossier / 'Liste.WBK').write_bytes(b'Word Windows')
            self.assertEqual(ranger_copies_secours(dossier), 1)
            self.assertEqual((dossier / DOSSIER_COPIES_SECOURS / 'Liste.WBK').read_bytes(), b'Word Windows')

    def test_secours_non_supprime_si_copie_echoue(self):
        with TemporaryDirectory() as temp:
            dossier = Path(temp)
            source = dossier / 'Copie de secours de Liste.docx'
            source.write_bytes(b'original')
            with patch('bibliotheque_administrative.shutil.copyfileobj', side_effect=OSError('Partage indisponible')):
                with self.assertRaises(OSError):
                    ranger_copies_secours(dossier)
            self.assertEqual(source.read_bytes(), b'original')
            self.assertEqual(list((dossier / DOSSIER_COPIES_SECOURS).iterdir()), [])

    def test_corbeille_limitee_aux_fichiers_du_dossier(self):
        with TemporaryDirectory() as temp:
            chantier = Path(temp) / "Chantier/Administratif"
            chantier.mkdir(parents=True)
            pdf = chantier / "ancien.pdf"
            pdf.write_bytes(b"pdf")
            autre = Path(temp) / "autre.pdf"
            autre.write_bytes(b"pdf")
            word = chantier / "modifiable.docx"
            word.write_bytes(b"word")
            fondation = ModuleType("Foundation")
            fondation.NSFileManager = SimpleNamespace(defaultManager=lambda: gestionnaire)
            fondation.NSURL = SimpleNamespace(fileURLWithPath_=lambda chemin: chemin)
            gestionnaire = Mock()
            gestionnaire.trashItemAtURL_resultingItemURL_error_.return_value = (True, None, None)
            with patch("main.sys.platform", "darwin"), patch.dict(sys.modules, {"Foundation": fondation}):
                mettre_fichier_corbeille(pdf, chantier)
                mettre_fichier_corbeille(word, chantier)
                with self.assertRaises(ValueError):
                    mettre_fichier_corbeille(autre, chantier)
            self.assertEqual([appel.args[0] for appel in gestionnaire.trashItemAtURL_resultingItemURL_error_.call_args_list],
                             [str(pdf), str(word)])
            gestionnaire.trashItemAtURL_resultingItemURL_error_.return_value = (False, None, "accès refusé")
            with patch("main.sys.platform", "darwin"), patch.dict(sys.modules, {"Foundation": fondation}):
                with self.assertRaises(OSError):
                    mettre_fichier_corbeille(pdf, chantier)

    def test_corbeille_windows_utilise_le_chemin_sans_interpolation(self):
        with TemporaryDirectory() as temp:
            document = Path(temp) / "titre d'entreprise.docx"
            document.write_bytes(b"word")
            with patch("main.sys.platform", "win32"), \
                 patch("main.subprocess.run") as executer:
                mettre_fichier_corbeille(document, Path(temp))
            self.assertEqual(executer.call_args.kwargs["env"]["HORIZON_TRASH_FILE"], str(document))

    def test_bibliotheque_en_cinq_dossiers_et_bundle_commun(self):
        with TemporaryDirectory() as temp:
            base = Path(temp)
            source = base / "Chantiers/Modele/Administratif"
            commun = source / "Document prêt/Dossier à envoyer/JT BATI"
            commun.mkdir(parents=True)
            (commun / "Certificat d'agréation.pdf").write_bytes(b"agrement")
            (commun / "ATTESTATION 19:01:2005.docx").write_bytes(b"modifiable")
            (commun.parent / "marche-particulier.pdf").write_bytes(b"particulier")
            bibliotheque = creer_bibliotheque_administrative(base, source, alimenter_depuis_source=True)
            self.assertTrue((bibliotheque / DOSSIERS_CATEGORIES[3] / "COMMUN__Certificat d'agréation.pdf").is_file())
            self.assertTrue((bibliotheque / DOSSIERS_CATEGORIES[4] / "COMMUN__ATTESTATION 19_01_2005.docx").is_file())
            self.assertFalse(any(p.name == "marche-particulier.pdf" for p in bibliotheque.rglob("*")))
            self.assertEqual({p.name for p in bibliotheque.iterdir()}, set(DOSSIERS_CATEGORIES))

    def test_ouverture_ne_copie_pas_les_documents_du_modele(self):
        with TemporaryDirectory() as temp:
            base = Path(temp)
            source = base / 'Modele/Administratif'
            source.mkdir(parents=True)
            (source / 'document.pdf').write_bytes(b'original')
            bibliotheque = creer_bibliotheque_administrative(base, source)
            self.assertFalse(any(p.is_file() for p in bibliotheque.rglob('*')))
            self.assertEqual((source / 'document.pdf').read_bytes(), b'original')

    def test_import_et_renommage_sans_ecraser(self):
        with TemporaryDirectory() as temp:
            base = Path(temp)
            bibliotheque = base / "Bibliotheque_administrative"
            bibliotheque.mkdir()
            for nom in DOSSIERS_CATEGORIES:
                (bibliotheque / nom).mkdir()
            source = base / "sources"
            source.mkdir()
            original = source / "méthodologie.odt"
            original.write_bytes(b"version 1")
            self.assertEqual(importer_dans_bibliotheque([original], bibliotheque), (1, 0))
            self.assertEqual(importer_dans_bibliotheque([original], bibliotheque), (0, 1))
            original.write_bytes(b"version 2")
            self.assertEqual(importer_dans_bibliotheque([original], bibliotheque), (1, 0))
            categorie = bibliotheque / DOSSIERS_CATEGORIES[2]
            self.assertEqual((categorie / "méthodologie.odt").read_bytes(), b"version 1")
            self.assertEqual((categorie / "méthodologie_2.odt").read_bytes(), b"version 2")
            chantier = base / "Chantiers/Exemple"
            choix = categorie / "méthodologie.odt"
            self.assertEqual(copier_documents_administratifs([choix], bibliotheque, chantier,
                                                             nom_copie="Méthodologie chantier A"), (1, 0))
            self.assertTrue((chantier / "Administratif/Méthodologie chantier A.odt").is_file())
            self.assertEqual(copier_documents_administratifs([choix], bibliotheque, chantier,
                                                             nom_copie="Méthodologie chantier A"), (0, 1))
            with self.assertRaises(ValueError):
                copier_documents_administratifs([choix], bibliotheque, chantier, nom_copie="../autre")

    def test_classement_archive_et_conserve_les_documents(self):
        with TemporaryDirectory() as temp:
            base = Path(temp)
            bibliotheque = base / "Bibliotheque_administrative"
            bibliotheque.mkdir()
            (bibliotheque / "COMMUN__Certificat d'agréation.pdf").write_bytes(b"certificat")
            (bibliotheque / "diplome.jpg").write_bytes(b"image")
            comptes, archive = classer_bibliotheque_cinq_dossiers(bibliotheque)
            self.assertEqual(sum(comptes.values()), 2)
            self.assertTrue((bibliotheque / DOSSIERS_CATEGORIES[3] / "Certificat d'agréation.pdf").is_file())
            self.assertTrue((bibliotheque / DOSSIERS_CATEGORIES[0] / "diplome.jpg").is_file())
            self.assertEqual({p.name for p in bibliotheque.iterdir()}, set(DOSSIERS_CATEGORIES))
            with ZipFile(archive) as sauvegarde:
                self.assertIn("COMMUN__Certificat d'agréation.pdf", sauvegarde.namelist())



if __name__ == "__main__":
    unittest.main()
