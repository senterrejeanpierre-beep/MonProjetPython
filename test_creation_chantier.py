import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from main import _creer_documents_chantier, _modeles_creation_chantier


class TestCreationChantier(unittest.TestCase):
    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.base = Path(self.tmp.name)
        self.modeles = self.base / 'Modele'
        self.modeles.mkdir()
        (self.modeles / 'Etat_avancement_02.xlsm').write_bytes(b'ancien modele')
        (self.modeles / 'prix_de_revient model.xlsx').write_bytes(b'prix')

    def test_reprise_dossier_vide_et_documents(self):
        (self.base / 'Chantier').mkdir()
        resultat = _creer_documents_chantier(self.base, 'Chantier', {'type_etat': 'Modèle actuel'})
        self.assertTrue(resultat.is_file())
        self.assertEqual((self.base / 'Chantier/data/prix_de_revient.xlsx').read_bytes(), b'prix')
        self.assertEqual((self.base / 'Chantier/Etat_avancement_02.xlsm').read_bytes(), b'ancien modele')

    def test_pas_de_substitution_public_prive(self):
        import json
        resultat = _creer_documents_chantier(self.base, 'Public', {'type_etat': 'Public'})
        fiche = json.loads(resultat.read_text())
        self.assertEqual(fiche['type_etat'], 'Public')
        self.assertTrue(fiche['modele_etat_manquant'])
        self.assertFalse(list((self.base / 'Public').glob('Etat_avancement*')))
        self.assertTrue((self.base / 'Public/data/prix_de_revient.xlsx').exists())

    def test_selection_accents_et_casse(self):
        (self.modeles / 'Etat_avancement_prive\u0301.xlsm').write_bytes(b'prive')
        (self.modeles / 'Etat_avancement_public.xlsm').write_bytes(b'public')
        self.assertIn('Privé', _modeles_creation_chantier(self.base)[2])
        _creer_documents_chantier(self.base, 'Prive', {'type_etat': 'Privé'})
        self.assertEqual(len(list((self.base / 'Prive').glob('Etat_avancement*'))), 2)
        self.assertEqual((self.base / 'Prive/Etat_avancement_public.xlsm').read_bytes(), b'public')
        self.assertEqual((self.base / 'Prive/Etat_avancement_prive\u0301.xlsm').read_bytes(), b'prive')

    def test_echec_copie_ne_laisse_aucun_chantier(self):
        with patch('main.shutil.copy2', side_effect=OSError('copie interrompue')):
            with self.assertRaises(OSError):
                _creer_documents_chantier(self.base, 'Erreur', {})
        self.assertEqual(list(self.base.iterdir()), [self.modeles])

    def test_dossier_existant_preserve(self):
        dossier = self.base / 'Existant'
        dossier.mkdir()
        document = dossier / 'important.txt'
        document.write_text('conserver')
        with self.assertRaises(FileExistsError):
            _creer_documents_chantier(self.base, 'Existant', {})
        self.assertEqual(document.read_text(), 'conserver')

    def test_nom_invalide(self):
        for nom in ('.', '..', '../ailleurs'):
            with self.assertRaises(ValueError):
                _creer_documents_chantier(self.base, nom, {})


if __name__ == '__main__':
    unittest.main()
