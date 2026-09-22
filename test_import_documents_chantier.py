import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from main import copier_fichier_importe, convertir_word_en_pdf_chantier


class ImportDocumentsChantierTests(unittest.TestCase):
    def test_import_reutilise_le_dossier_et_preserve_les_fichiers(self):
        with tempfile.TemporaryDirectory() as temporaire:
            base = Path(temporaire)
            source = base / "formulaire offre.docx"
            source.write_bytes(b"premier")
            dossier = base / "Chantiers" / "004_Mariemont" / "Administratif"
            premier, ajoute = copier_fichier_importe(source, dossier)
            self.assertTrue(ajoute)
            self.assertEqual(premier.read_bytes(), b"premier")
            source.write_bytes(b"modifie")
            second, ajoute = copier_fichier_importe(source, dossier)
            self.assertTrue(ajoute)
            self.assertEqual(premier.read_bytes(), b"premier")
            self.assertEqual(second.name, "formulaire offre_2.docx")
            self.assertEqual(second.read_bytes(), b"modifie")
            self.assertEqual(copier_fichier_importe(premier, dossier), (premier, False))

    def test_conversion_pdf_reste_dans_le_chantier_sans_ecraser(self):
        with tempfile.TemporaryDirectory() as temporaire:
            dossier_word = Path(temporaire) / "Library" / "Containers" / "com.microsoft.Word" / "Data" / "Documents"
            dossier_word.mkdir(parents=True)
            def convertir(commande, **_options):
                self.assertEqual(commande[:3], ["osascript", "-l", "JavaScript"])
                self.assertIn("word.open(argv[0])", commande[4])
                self.assertEqual(_options["timeout"], 30)
                self.assertNotEqual(Path(commande[-2]).parent, Path(commande[-1]).parent)
                Path(commande[-1]).write_bytes(b"nouveau PDF")
                return SimpleNamespace(returncode=0, stdout="", stderr="")

            with patch("main.sys.platform", "darwin"), patch("main.Path.home", return_value=Path(temporaire)), \
                    patch("main.subprocess.run", side_effect=convertir):
                for chantier, categorie in (("004_Mariemont", "Administratif"),
                                            ("027_Autre chantier", "Technique")):
                    dossier = Path(temporaire) / "Chantiers" / chantier / categorie
                    dossier.mkdir(parents=True)
                    source = dossier / "formulaire offre.docx"
                    source.write_bytes(b"word")
                    (dossier / "formulaire offre.pdf").write_bytes(b"ancien")
                    pdf = convertir_word_en_pdf_chantier(source)
                    self.assertEqual(pdf, dossier / "formulaire offre_2.pdf")
                    self.assertEqual(pdf.read_bytes(), b"nouveau PDF")
                    self.assertEqual((dossier / "formulaire offre.pdf").read_bytes(), b"ancien")

    def test_conversion_windows_utilise_word_et_le_dossier_du_chantier(self):
        with tempfile.TemporaryDirectory() as temporaire:
            dossier = Path(temporaire) / "Chantiers" / "027_Autre chantier" / "Administratif"
            dossier.mkdir(parents=True)
            source = dossier / "offre.docx"
            source.write_bytes(b"word")

            def convertir(commande, **_options):
                self.assertEqual(commande[0], "powershell.exe")
                self.assertIn("ExportAsFixedFormat", Path(commande[-3]).read_text(encoding="utf-8-sig"))
                Path(commande[-1]).write_bytes(b"PDF")
                return SimpleNamespace(returncode=0, stdout="", stderr="")

            with patch("main.sys.platform", "win32"), patch("main.subprocess.run", side_effect=convertir):
                self.assertEqual(convertir_word_en_pdf_chantier(source), dossier / "offre.pdf")

    def test_conversion_mac_exporte_une_seule_fois_sur_le_poste(self):
        with tempfile.TemporaryDirectory() as temporaire:
            dossier_word = Path(temporaire) / "Library" / "Containers" / "com.microsoft.Word" / "Data" / "Documents"
            dossier_word.mkdir(parents=True)
            dossier = Path(temporaire) / "Chantiers" / "031_Chantier déplacé" / "Administratif"
            dossier.mkdir(parents=True)
            source = dossier / "offre.docx"
            source.write_bytes(b"word")
            appels = []

            def convertir(commande, **_options):
                appels.append(commande[-1])
                Path(commande[-1]).write_bytes(b"PDF")
                return SimpleNamespace(returncode=0, stdout="", stderr="")

            with patch("main.sys.platform", "darwin"), patch("main.Path.home", return_value=Path(temporaire)), \
                    patch("main.subprocess.run", side_effect=convertir):
                self.assertEqual(convertir_word_en_pdf_chantier(source), dossier / "offre.pdf")
            self.assertEqual(len(appels), 1)
            self.assertEqual((dossier / "offre.pdf").read_bytes(), b"PDF")


if __name__ == "__main__":
    unittest.main()
