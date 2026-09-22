import tempfile
import unittest
from contextlib import ExitStack
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from main import (HorizonChantierApp, _poste_pour_fichier_bibliotheque,
                  creer_dossier_bibliotheque, parent_nouveau_dossier_bibliotheque,
                  renommer_dossier_bibliotheque)


class ImportBibliotheques(unittest.TestCase):
    def test_documents_cout_ouvrables_meme_avec_conflit_excel(self):
        with tempfile.TemporaryDirectory() as tmp, ExitStack() as stack:
            chantier = Path(tmp) / 'Chantier'
            chantier.mkdir()
            ouverture = Mock()
            app = SimpleNamespace(_nom_chantier_selectionne=lambda: 'Chantier',
                                  _dossier_chantier_selectionne=lambda: chantier,
                                  _ouvrir_documents_chantier=ouverture)
            stack.enter_context(patch('main._chemin_fichier_chantier', return_value=chantier / 'Cout_securite.xlsx'))
            stack.enter_context(patch('main.lire_postes', return_value=[{'libelle': '193 · Ancien libellé'}]))
            stack.enter_context(patch('main.dossier_poste', return_value=chantier / 'Documents_installation/193'))
            synchroniser = stack.enter_context(patch('main.actualiser_liens', side_effect=ValueError('Conflit Excel 193')))
            erreur = stack.enter_context(patch('main.messagebox.showerror'))
            HorizonChantierApp.ouvrir_documents_installation(app)
            ouverture.assert_called_once_with('Documents_installation')
            synchroniser.assert_not_called()
            erreur.assert_not_called()

    def test_ouverture_cout_ne_modifie_pas_le_classeur(self):
        with tempfile.TemporaryDirectory() as tmp, ExitStack() as stack:
            chantier = Path(tmp)
            classeur = chantier / 'Cout_securite.xlsx'
            classeur.write_bytes(b'classeur existant')
            app = SimpleNamespace(_nom_chantier_selectionne=lambda: 'Chantier',
                                  _dossier_chantier_selectionne=lambda: chantier,
                                  _bibliotheque_cout=lambda: chantier / 'Bibliotheque',
                                  _preparer_ouverture_document=lambda _chemin: True,
                                  _doit_rappeler_pdf_historique=lambda _chemin: False)
            stack.enter_context(patch('main._chemin_fichier_chantier', return_value=classeur))
            synchroniser = stack.enter_context(patch('main.actualiser_liens', side_effect=ValueError('Conflit Excel 193')))
            avertissement = stack.enter_context(patch('main.messagebox.showwarning'))
            ouvrir = stack.enter_context(patch('main.ouvrir_chemin'))
            HorizonChantierApp._ouvrir_fichier_chantier(app, 'Cout_securite.xlsx')
            synchroniser.assert_not_called()
            avertissement.assert_not_called()
            ouvrir.assert_called_once_with(classeur)

    def test_rubriques_securite_ajoutables_avant_et_apres_le_modele(self):
        with tempfile.TemporaryDirectory() as tmp:
            racine = Path(tmp)
            (racine / '021_Existant').mkdir()
            (racine / '192_Existant').mkdir()
            debut = creer_dossier_bibliotheque(racine, '1', 'Préparation sécurité')
            fin = creer_dossier_bibliotheque(racine, '193', 'Contrôle final')
            self.assertEqual(debut.name, '001_Préparation_sécurité')
            self.assertEqual(fin.name, '193_Contrôle_final')
            with self.assertRaises(FileExistsError):
                creer_dossier_bibliotheque(racine, '021', 'Doublon')
            with self.assertRaises(ValueError):
                creer_dossier_bibliotheque(racine, '../2', 'Invalide')

    def test_dossier_numerote_dans_categorie_administrative_ou_technique(self):
        with tempfile.TemporaryDirectory() as tmp:
            racine = Path(tmp)
            categorie = racine / '01_Personnel_et_diplomes'
            categorie.mkdir()
            dossier = creer_dossier_bibliotheque(categorie, '6', 'Nouveaux documents')
            self.assertEqual(dossier.parent, categorie)
            self.assertEqual(dossier.name, '006_Nouveaux_documents')
            self.assertEqual(parent_nouveau_dossier_bibliotheque(racine, dossier, False), categorie)
            self.assertEqual(parent_nouveau_dossier_bibliotheque(racine, dossier, True), racine)

    def test_renommer_rubrique_conserve_numero_et_documents(self):
        with tempfile.TemporaryDirectory() as tmp:
            racine = Path(tmp)
            ancien = creer_dossier_bibliotheque(racine, '193', 'Ancien nom')
            (ancien / 'offre.pdf').write_bytes(b'document')
            nouveau = renommer_dossier_bibliotheque(ancien, 'Protection des accès pour engin', racine)
            self.assertEqual(nouveau.name, '193_Protection_des_accès_pour_engin')
            self.assertEqual((nouveau / 'offre.pdf').read_bytes(), b'document')
            self.assertFalse(ancien.exists())

    def test_bibliotheque_securite_ajoute_193_sans_toucher_aux_existants(self):
        with tempfile.TemporaryDirectory() as tmp, ExitStack() as stack:
            base = Path(tmp)
            chantiers = base / 'Chantiers'
            chantiers.mkdir()
            stack.enter_context(patch('main.dossier_chantiers', return_value=chantiers))
            stack.enter_context(patch('main._modeles_creation_chantier', return_value=(base, {}, {})))
            stack.enter_context(patch('main._chemin_fichier_chantier', return_value=base / 'modele.xlsx'))
            stack.enter_context(patch('main.lire_postes', return_value=[{'libelle': '192 · Poste'}]))
            stack.enter_context(patch('main.dossier_poste', return_value=base / '192_Poste'))
            bibliotheque = HorizonChantierApp._bibliotheque_cout(SimpleNamespace())
            attendu = bibliotheque / '193_Protection_des_accès_pour_engin'
            self.assertTrue(attendu.is_dir())
            (attendu / 'document.pdf').write_bytes(b'conserve')
            HorizonChantierApp._bibliotheque_cout(SimpleNamespace())
            self.assertEqual((attendu / 'document.pdf').read_bytes(), b'conserve')

    def test_import_direct_isole_sans_ecrasement_et_annulation(self):
        for categorie in ('Technique', 'Administratif', 'Documents_installation'):
            with self.subTest(categorie=categorie), tempfile.TemporaryDirectory() as tmp, ExitStack() as stack:
                base = Path(tmp)
                bibliotheque = base / 'Bibliotheque'
                bibliotheque.mkdir()
                source = bibliotheque / 'document.pdf'
                source.write_bytes(b'original bibliotheque')
                chantier = base / 'Chantiers' / 'Selectionne'
                dossier = chantier / categorie
                if categorie == 'Documents_installation':
                    dossier /= '007_Poste_propre_au_chantier'
                dossier.mkdir(parents=True)
                (dossier / source.name).write_bytes(b'document existant a conserver')
                autre = base / 'Chantiers' / 'Autre'
                autre.mkdir()
                app = SimpleNamespace(_nom_chantier_selectionne=lambda: 'Selectionne',
                                      _dossier_chantier_selectionne=lambda: chantier,
                                      _bibliotheque_cout=lambda: bibliotheque,
                                      _ouvrir_documents_chantier=Mock())
                def choisir(*args, **kwargs):
                    if 'poste_selection' in kwargs:
                        kwargs['poste_selection']['choix'] = dossier.name
                    return str(source)
                stack.enter_context(patch('main.dossier_chantiers', return_value=chantier.parent))
                stack.enter_context(patch('main._modeles_creation_chantier', return_value=(base, {}, {})))
                stack.enter_context(patch('main.creer_bibliotheque_administrative', return_value=bibliotheque))
                stack.enter_context(patch('main.creer_bibliotheque_classee', return_value=(bibliotheque, 1, 1)))
                stack.enter_context(patch('main.lire_postes', return_value=[{'libelle': '007 - Poste'}]))
                stack.enter_context(patch('main.dossier_poste', return_value=dossier))
                synchroniser = stack.enter_context(patch('main.actualiser_liens', side_effect=ValueError('Conflit Excel 193')))
                selection = stack.enter_context(patch('main.choisir_fichier_bibliotheque', side_effect=choisir))
                ouvrir = stack.enter_context(patch('main.ouvrir_chemin'))
                confirmer = stack.enter_context(patch('main.demander_confirmation'))
                stack.enter_context(patch('main.messagebox.showinfo'))
                erreur = stack.enter_context(patch('main.messagebox.showerror'))
                for _ in range(2):
                    HorizonChantierApp._importer_depuis_bibliotheque(app, categorie)
                erreur.assert_not_called()
                synchroniser.assert_not_called()
                ouvrir.assert_not_called()
                confirmer.assert_not_called()
                self.assertEqual(source.read_bytes(), b'original bibliotheque')
                self.assertEqual((dossier / source.name).read_bytes(), b'document existant a conserver')
                self.assertEqual(sorted(p.read_bytes() for p in dossier.iterdir()),
                                 [b'document existant a conserver', b'original bibliotheque'])
                self.assertEqual(list(autre.iterdir()), [])
                app._ouvrir_documents_chantier.assert_not_called()
                selection.side_effect = None
                selection.return_value = ''
                HorizonChantierApp._importer_depuis_bibliotheque(app, categorie)
                app._ouvrir_documents_chantier.assert_not_called()

    def test_poste_du_cout_deduit_du_dossier_de_bibliotheque(self):
        racine = Path('/partage/Bibliotheque_Cout_securite')
        postes = ['190_Transport', '191_Debroussaillage']
        fichier = racine / '191_Debroussaillage' / 'fiche.pdf'
        self.assertEqual(_poste_pour_fichier_bibliotheque(fichier, racine, postes),
                         '191_Debroussaillage')
        self.assertEqual(_poste_pour_fichier_bibliotheque(racine / 'divers.pdf', racine, postes), '')

    def test_entrees_bibliotheques_utilisent_le_meme_parcours(self):
        for methode, categorie in (
            ('ouvrir_fiches_techniques', 'Technique'),
            ('choisir_fiche_technique_chantier', 'Technique'),
            ('ouvrir_bibliotheque_administrative', 'Administratif'),
            ('choisir_document_administratif_chantier', 'Administratif'),
            ('ouvrir_bibliotheque_cout', 'Documents_installation'),
        ):
            app = SimpleNamespace(_importer_depuis_bibliotheque=Mock())
            getattr(HorizonChantierApp, methode)(app)
            app._importer_depuis_bibliotheque.assert_called_once_with(categorie)


if __name__ == '__main__':
    unittest.main()
