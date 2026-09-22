import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from preparer_fiches_techniques import preparer_chantier


class PreparationFichesTechniques(unittest.TestCase):
    def test_dossiers_et_liens_relatifs_sans_effacer_les_documents(self):
        with tempfile.TemporaryDirectory() as temp:
            chantier = Path(temp)
            chaux = chantier / "Fiches_techniques/014_Chaux"
            chaux.mkdir(parents=True)
            document = chaux / "fiche.pdf"
            document.write_bytes(b"document")
            index = preparer_chantier(chantier, [("014", "Chaux"), ("015", "Couverture")])
            self.assertTrue((chantier / "Fiches_techniques/015_Couverture").is_dir())
            classeur = load_workbook(index)
            try:
                self.assertEqual(classeur.active["C2"].hyperlink.target, "014_Chaux/")
            finally:
                classeur.close()
            preparer_chantier(chantier, [("014", "Chaux")])
            self.assertEqual(document.read_bytes(), b"document")


if __name__ == "__main__":
    unittest.main()
