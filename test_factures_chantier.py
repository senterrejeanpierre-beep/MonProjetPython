import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import Mock

from factures_chantier import DOSSIERS_FACTURES, preparer_dossiers_factures
from main import HorizonChantierApp


class ClassementFacturesTests(unittest.TestCase):
    def test_modele_et_tous_chantiers_sans_deplacer_les_documents(self):
        with TemporaryDirectory() as temporaire:
            base = Path(temporaire)
            modele = base / 'Modele'
            modele.mkdir()
            for nom in ('001_Chapelette', '002_Saint-Symphorien',
                        '003_Frederic_melin urgence', '004_Mariemont'):
                chantier = base / nom
                (chantier / 'Factures').mkdir(parents=True)
                (base / f'{nom}.json').write_text('{}')
                (chantier / 'Factures' / 'ancien.pdf').write_bytes(nom.encode())
            prepares, erreurs = preparer_dossiers_factures(base, modele)
            self.assertEqual(erreurs, [])
            self.assertEqual(len(prepares), 5)
            for dossier in prepares:
                self.assertTrue(all((dossier / nom).is_dir() for nom in DOSSIERS_FACTURES))
            for nom in ('001_Chapelette', '002_Saint-Symphorien',
                        '003_Frederic_melin urgence', '004_Mariemont'):
                self.assertEqual((base / nom / 'Factures' / 'ancien.pdf').read_bytes(), nom.encode())
            self.assertEqual(preparer_dossiers_factures(base, modele)[1], [])

    def test_boutons_ouvrent_directement_chaque_categorie(self):
        with TemporaryDirectory() as temporaire:
            chantier = Path(temporaire) / 'Chantier'
            chantier.mkdir()
            app = SimpleNamespace(_nom_chantier_selectionne=lambda: 'Chantier',
                                  _dossier_chantier_selectionne=lambda: chantier,
                                  _ouvrir_documents_chantier=Mock())
            for categorie in DOSSIERS_FACTURES:
                HorizonChantierApp.ouvrir_factures_chantier(app, categorie)
                app._ouvrir_documents_chantier.assert_called_with('Factures', factures_type=categorie)
            self.assertEqual(app._ouvrir_documents_chantier.call_count, len(DOSSIERS_FACTURES))


if __name__ == '__main__':
    unittest.main()
