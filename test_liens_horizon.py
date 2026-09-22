import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import liens_horizon as liens


class LiensTests(unittest.TestCase):
    def test_identite_et_poste_sur_partage_deplace(self):
        with tempfile.TemporaryDirectory() as d:
            base = Path(d)
            chantier = base / 'Chantier A'
            chantier.mkdir()
            lien = liens.lien_poste(chantier, 176)
            self.assertEqual(liens.resoudre(base, lien), (chantier, '176'))
            nouveau = base / 'Chantier renommé'
            chantier.rename(nouveau)
            self.assertEqual(liens.resoudre(base, lien), (nouveau, '176'))
            self.assertNotIn(str(base), lien)

    def test_requete_locale_recuperee_une_fois(self):
        with tempfile.TemporaryDirectory() as d, patch.object(liens, 'dossier_local', return_value=Path(d) / 'local'):
            lien = liens.lien_poste(Path(d), 175)
            self.assertFalse(liens.deposer(lien))
            self.assertEqual(list(liens.commandes()), [lien])
            self.assertEqual(list(liens.commandes()), [])
            self.assertTrue(liens.deposer(lien))

    def test_refuse_liens_arbitraires(self):
        for lien in ('https://example.org', 'horizonchantier://documents/../../fichier',
                     'horizonchantier://documents/' + 'a'*24 + '/175?commande=autre'):
            with self.assertRaises(ValueError):
                liens.analyser(lien)


if __name__ == '__main__':
    unittest.main()
