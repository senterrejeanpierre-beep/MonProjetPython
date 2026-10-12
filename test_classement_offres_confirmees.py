import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from factures_chantier import categorie_offre, classer_offres_confirmees
from bibliotheque_prix_materiaux import NOM


class ClassementOffresConfirmeesTests(unittest.TestCase):
    def test_categories_confirmees(self):
        for nom, categorie in (
            ('Échafaudage/location.pdf', 'Offres_sous_traitants'),
            ('Essais/Offre.pdf', 'Offres_sous_traitants'),
            ('Stone Assistance/Etude des mortiers.pdf', 'Offres_sous_traitants'),
            ('Lietar/tarif.pdf', 'Offres_location_materiel'),
            ('Loiselet/tarif.pdf', 'Offres_location_materiel'),
            ('Couvre-mur végétal.pdf', 'Offres_materiaux'),
            ('Inconnu/Offre.pdf', None),
        ):
            with self.subTest(nom=nom):
                self.assertEqual(categorie_offre(Path(nom)), categorie)

    def test_reprise_sans_doublon_ecrasement_ni_modification_original(self):
        with TemporaryDirectory() as tmp:
            chantier = Path(tmp) / 'Chantier'
            ancien = chantier / NOM
            for fournisseur in ('Lietar', 'Loiselet', 'Essais', 'Couvre-mur', 'Inconnu'):
                source = ancien / fournisseur / 'offre.pdf'
                source.parent.mkdir(parents=True)
                source.write_bytes(fournisseur.encode())
            self.assertEqual(classer_offres_confirmees(chantier), 4)
            self.assertEqual(classer_offres_confirmees(chantier), 0)
            locations = chantier / 'Offres_location_materiel'
            self.assertEqual({p.read_bytes() for p in locations.iterdir()}, {b'Lietar', b'Loiselet'})
            for source in ancien.rglob('*.pdf'):
                self.assertEqual(source.read_bytes(), source.parent.name.encode())
            self.assertEqual(len(list(ancien.rglob('*.pdf'))), 5)


if __name__ == '__main__':
    unittest.main()
