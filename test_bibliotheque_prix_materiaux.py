import tempfile
import json
import unittest
from pathlib import Path

from bibliotheque_prix_materiaux import (
    NOM, actualiser_bibliotheque_chantier, creer_bibliotheque,
    importer_offres, trouver_offres,
)


class BibliothequePrixMateriauxTests(unittest.TestCase):
    def test_generale_reprend_toutes_categories_et_anciennes_offres(self):
        from factures_chantier import DOSSIERS_OFFRES
        with tempfile.TemporaryDirectory() as temporaire:
            base = Path(temporaire)
            chantiers = base / "Chantiers"
            sources = []
            for chantier in ("001_Existant", "004_Mariemont", "005_Futur"):
                for categorie in (NOM, *DOSSIERS_OFFRES):
                    source = chantiers / chantier / categorie / "Offre.pdf"
                    source.parent.mkdir(parents=True)
                    source.write_bytes(f"{chantier}/{categorie}".encode())
                    sources.append(source)
            generale = creer_bibliotheque(base)
            offres = trouver_offres(chantiers)
            self.assertEqual(set(offres), set(sources))
            self.assertEqual(importer_offres(offres, chantiers, generale), (12, 0))
            self.assertEqual(importer_offres(offres, chantiers, generale), (0, 12))
            self.assertEqual({p.read_bytes() for p in generale.rglob("*.pdf")},
                             {p.read_bytes() for p in sources})
            sources[0].write_bytes(b"offre revisee")
            self.assertEqual(importer_offres([sources[0]], chantiers, generale), (1, 0))
            self.assertEqual(len(list(generale.rglob("*.pdf"))), 13)

    def test_deux_offres_essais_hors_dossier_demande_prix_standard(self):
        with tempfile.TemporaryDirectory() as temporaire:
            base = Path(temporaire)
            chantiers = base / "Chantiers"
            dossier = chantiers / "004_Test" / "Demande de prix sous-traitants" / "Contrôles et essais"
            dossier.mkdir(parents=True)
            for nom in ("Offre de prix 20260914.pdf", "Offre de prix 20260925 avec 6 pcs.pdf",
                        "Cahier_technique.pdf", "Bordereau contrôles.xlsx", "Plans.pdf"):
                (dossier / nom).write_bytes(nom.encode())
            offres = trouver_offres(chantiers)
            self.assertEqual(len(offres), 2)
            generale = creer_bibliotheque(base)
            self.assertEqual(importer_offres(offres, chantiers, generale), (2, 0))
            self.assertEqual(len(list(generale.rglob("*.pdf"))), 2)

    def test_capture_prix_classee_et_renommee(self):
        with tempfile.TemporaryDirectory() as temporaire:
            base = Path(temporaire)
            chantiers = base / "Chantiers"
            chantier = chantiers / "004_Test"
            source = chantier / "Demande de prix" / "Superplastic" / "Geotextile.png"
            source.parent.mkdir(parents=True)
            source.write_bytes(b"capture originale")
            locale = actualiser_bibliotheque_chantier(chantier)
            copie = locale / "Superplastic" / source.name
            self.assertEqual(copie.read_bytes(), source.read_bytes())
            copie.rename(copie.with_name("Mon geotextile.png"))
            actualiser_bibliotheque_chantier(chantier)
            self.assertFalse(copie.exists())
            generale = creer_bibliotheque(base)
            importer_offres(trouver_offres(chantiers), chantiers, generale)
            self.assertEqual((generale / "Superplastic" / chantier.name / source.name).read_bytes(), source.read_bytes())

    def test_rangement_source_reprise_et_provenance(self):
        with tempfile.TemporaryDirectory() as temporaire:
            base = Path(temporaire)
            chantiers = base / "Chantiers"
            sources = []
            for nom in ("001_A", "002_B", "Modele"):
                source = chantiers / nom / "Demande de prix" / "Échafaudage" / "ACT" / "Offre.pdf"
                source.parent.mkdir(parents=True)
                source.write_bytes(b"meme offre")
                sources.append(source)
            generale = creer_bibliotheque(base)
            (generale / ".classement_fournisseurs.json").write_text(json.dumps({
                source.relative_to(chantiers).as_posix(): "ACT" for source in sources
            }), encoding="utf-8")
            ancien = generale / "23_Echafaudage" / "001_A__Offre.pdf"
            ancien.parent.mkdir()
            ancien.write_bytes(b"meme offre")
            offres = trouver_offres(chantiers)
            self.assertEqual(len(offres), 2)
            self.assertEqual(importer_offres(offres, chantiers, generale), (2, 0))
            self.assertFalse(ancien.exists())
            for nom in ("001_A", "002_B"):
                self.assertEqual((generale / "ACT" / nom / "Offre.pdf").read_bytes(), b"meme offre")
            self.assertEqual(importer_offres(offres, chantiers, generale), (0, 2))
            locale = actualiser_bibliotheque_chantier(chantiers / "001_A")
            self.assertTrue((locale / "ACT" / "Offre.pdf").is_file())
            with self.assertRaises(ValueError):
                importer_offres([sources[1]], chantiers, locale)
            original = generale / "ACT/001_A/Offre.pdf"
            original.rename(original.with_name("Briques.pdf"))
            importer_offres(offres, chantiers, generale)
            self.assertFalse(original.exists())
            self.assertEqual(len(list(original.parent.glob("*.pdf"))), 1)
            original.with_name("Briques.pdf").rename(original)
            sources[0].write_bytes(b"offre revisee")
            importer_offres([sources[0]], chantiers, generale)
            self.assertEqual((generale / "ACT/001_A/Offre.pdf").read_bytes(), b"meme offre")
            self.assertEqual((generale / "ACT/001_A/Offre_2.pdf").read_bytes(), b"offre revisee")

    def test_isolation_actualisation_et_bibliotheque_generale(self):
        with tempfile.TemporaryDirectory() as temporaire:
            base = Path(temporaire)
            chantiers = base / "Chantiers"
            a, b = chantiers / "001_Test", chantiers / "002_Test"
            sources = []
            for chantier, contenu in ((a, b"offre A"), (b, b"offre B")):
                dossier = chantier / "Demande de prix" / "Briques"
                dossier.mkdir(parents=True)
                source = dossier / "Offre fournisseur.pdf"
                source.write_bytes(contenu)
                sources.append(source)
            locale = actualiser_bibliotheque_chantier(a)
            self.assertFalse((b / NOM).exists())
            self.assertEqual([p.read_bytes() for p in locale.rglob("*.pdf")], [b"offre A"])
            actualiser_bibliotheque_chantier(a)
            self.assertEqual(len(list(locale.rglob("*.pdf"))), 1)
            nouvelle = sources[0].with_name("Offre complementaire.pdf")
            nouvelle.write_bytes(b"nouvelle offre")
            actualiser_bibliotheque_chantier(a)
            self.assertEqual(len(list(locale.rglob("*.pdf"))), 2)
            generale = creer_bibliotheque(base)
            self.assertEqual(importer_offres(trouver_offres(chantiers), chantiers, generale), (3, 2))
            self.assertEqual(importer_offres(trouver_offres(chantiers), chantiers, generale), (0, 5))
            self.assertEqual([p.read_bytes() for p in sources], [b"offre A", b"offre B"])
            # La même fonction prend en charge un chantier créé plus tard.
            futur = chantiers / "003_Futur"
            futur.mkdir()
            self.assertTrue(actualiser_bibliotheque_chantier(futur).is_dir())


if __name__ == "__main__":
    unittest.main()
