import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from index_fiches_techniques import creer_index_fiches


class IndexFichesTechniques(unittest.TestCase):
    def test_categories_et_liens_relatifs_sans_deplacer_les_documents(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            bibliotheque = base / "FICHES TECHNIQUES"
            chaux = bibliotheque / "Fiches techniques 1" / "Chaux"
            chaux.mkdir(parents=True)
            fiche = chaux / "Fiche chaux.pdf"
            fiche.write_bytes(b"fiche")
            direct = bibliotheque / "Catalogue.pdf"
            direct.write_bytes(b"catalogue")
            index = base / "Index_fiches_techniques.xlsx"

            resultat = creer_index_fiches(base, index)

            self.assertEqual(resultat, {"categories": 1, "dossiers": 2,
                                         "fichiers_racine": 1})
            wb = load_workbook(index)
            try:
                categorie = wb["Catégories"]
                self.assertEqual(categorie["B5"].value, "Chaux")
                self.assertEqual(categorie["E5"].hyperlink.target,
                                 "FICHES TECHNIQUES/Fiches techniques 1/Chaux/")
                self.assertEqual(wb["Fichiers à la racine"]["D5"].hyperlink.target,
                                 "FICHES TECHNIQUES/Catalogue.pdf")
            finally:
                wb.close()
            self.assertEqual(fiche.read_bytes(), b"fiche")
            self.assertEqual(direct.read_bytes(), b"catalogue")


if __name__ == "__main__":
    unittest.main()
