import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import ModuleType, SimpleNamespace
import sys
from unittest.mock import Mock, patch
from zipfile import ZipFile

from bibliotheque_administrative import (DOSSIERS_CATEGORIES, classer_bibliotheque_cinq_dossiers,
                                         creer_bibliotheque_administrative,
                                         copier_documents_administratifs, importer_dans_bibliotheque)
from main import HorizonChantierApp, mettre_fichier_corbeille


class BibliothequeAdministrativeTests(unittest.TestCase):
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
