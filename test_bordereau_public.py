import hashlib
import tempfile
import unittest
import zipfile
from pathlib import Path
from openpyxl import Workbook, load_workbook
from bordereau_public import preparer, exporter_prix, feuille, habiller_avancement
from main import _corriger_moins_etat_public, _selectionner_etat, _recalculer_etat, _reporter_quantites_cloture

class BordereauPublic(unittest.TestCase):
    def test_ancienne_synthese_publique_retire_moins_sans_effacer_saisie(self):
        ws = Workbook().active
        ws["P20"] = "Total des avenants cumulé"
        ws["T20"] = "=SUM('[1]Feuil1'!$D$3:$D$60)-SUM('[1]Feuil1'!$E$3:$E$60)"
        ws["P18"] = "Total état cumulé hors TVA"
        ws["U19"] = '="Montant de l’avenant "&Q9'
        ws["W19"] = ("=N(INDEX('[1]Feuil1'!$D$3:$D$60,1+3*($Q$9-1)))"
                     "-N(INDEX('[1]Feuil1'!$E$3:$E$60,1+3*($Q$9-1)))")
        ws["W20"] = "=SUM(W18:W19)"
        ws["P21"] = "Avenants cumulés en moins"
        ws["T21"] = 0
        _corriger_moins_etat_public(ws)
        self.assertEqual(ws["T20"].value, "=SUM('[1]Feuil1'!$D$3:$D$60)")
        self.assertIsNone(ws["U19"].value)
        self.assertIsNone(ws["W19"].value)
        self.assertEqual(ws["W20"].value, "=W18")
        self.assertIsNone(ws["P21"].value)
        self.assertIsNone(ws["T21"].value)
        ws["P21"], ws["T21"] = "Avenants cumulés en moins", 200
        with self.assertRaisesRegex(ValueError, "Vérifiez"):
            _corriger_moins_etat_public(ws)
        self.assertEqual(ws["T21"].value, 200)

    def test_circuit_prix_et_avancement(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            original = d / 'original.xlsx'
            w = Workbook(); s = w.active; s.title = 'MARIEMONT'
            for cell, v in {'B1':'N°', 'K1':'QUANTITÉ', 'L1':'PRIX UNITAIRES',
                            'B13':'02.21.4a.01', 'C13':'Essais', 'D13':'QP', 'E13':'PCE',
                            'K13':3, 'L13':0, 'M13':'=L13*K13', 'O13':'Note officielle'}.items():
                s[cell] = v
            w.save(original); w.close()
            avant = original.read_bytes()
            copie = preparer(original, d)
            self.assertEqual(_selectionner_etat(d), copie)
            w = load_workbook(copie); s = feuille(w)
            self.assertEqual(s['P16'].value, 'Total état cumulé hors TVA')
            self.assertEqual(s['T16'].value, '=SUM(T13)')
            self.assertIsNone(s['U17'].value)
            self.assertIsNone(s['W17'].value)
            self.assertEqual(s['W18'].value, '=W16')
            self.assertEqual(s['W19'].value, '=W18*$W$9')
            self.assertEqual(s['W21'].value, '=SUM(W18:W20)')
            self.assertEqual(s['T22'].value, '=T16')  # Le mois n'est pas ajouté deux fois.
            self.assertNotIn('$E$3', s['T20'].value)  # Les moins restent hors de l'état exécuté.
            self.assertIsNone(s['P27'].value)  # Aucune ligne « avenants en moins » dans la synthèse.
            self.assertEqual(s['U21'].fill.fgColor.rgb[-6:], 'FFFF00')
            self.assertEqual([l.file_link.Target for l in w._external_links],
                             ['Avenants.xlsx', 'Revision_global.xlsx'])
            self.assertIn('$Q$9', s['W20'].value)
            self.assertNotIn('676', s['W21'].value)
            s['L13'] = 12.5; s['P13'] = 1; s['Q13'] = 1
            s['Q9'] = 3
            habiller_avancement(w)
            self.assertEqual(s['Q9'].value, 3)
            self.assertEqual(s['L13'].value, 12.5)
            self.assertEqual(s['P13'].value, 1)
            self.assertEqual(s['Q13'].value, 1)
            _recalculer_etat(s)
            self.assertEqual(s['T13'].value, 25)
            self.assertEqual(s['W13'].value, 12.5)
            _reporter_quantites_cloture(s)
            self.assertEqual(s['P13'].value, 2)
            self.assertEqual(s['Q13'].value, 0)
            w.save(copie); w.close()
            export = d / 'export.xlsx'
            self.assertEqual(exporter_prix(original, copie, export, hashlib.sha256(avant).hexdigest()), 1)
            self.assertEqual(original.read_bytes(), avant)
            with zipfile.ZipFile(original) as a, zipfile.ZipFile(export) as b:
                self.assertEqual(a.namelist(), b.namelist())
                for name in a.namelist():
                    if name != 'xl/worksheets/sheet1.xml':
                        self.assertEqual(a.read(name), b.read(name))
            w = load_workbook(export)
            self.assertEqual(w.active['L13'].value, 12.5)
            self.assertEqual(w.active['M13'].value, '=L13*K13')
            self.assertEqual(w.active['O13'].value, 'Note officielle')
            w.close()
            with self.assertRaises(FileExistsError):
                preparer(original, d)
            copie.unlink()
            with self.assertRaises(FileNotFoundError):
                _selectionner_etat(d)

    def test_refus_original_modifie(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / 'original.xlsx'; p.write_bytes(b'changed')
            with self.assertRaisesRegex(ValueError, 'original a changé'):
                exporter_prix(p, p, Path(tmp) / 'export.xlsx', 'ancienne empreinte')

if __name__ == '__main__':
    unittest.main()
