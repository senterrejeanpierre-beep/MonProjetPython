"""Parcours de sauvegarde ; sélection et automatisation Excel simulées."""
from contextlib import ExitStack
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import hashlib
import shutil
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook
from reportlab.pdfgen.canvas import Canvas
from main import HorizonChantierApp, _export_excel_to_pdf
from pypdf import PdfReader


class PDFHistoriqueTests(unittest.TestCase):
    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.chantier = Path(self.tmp.name) / 'Partage été' / '006_Nouveau chantier'
        self.source = self.chantier / 'data' / 'prix_de_revient.xlsx'
        self.source.parent.mkdir(parents=True)
        wb = Workbook()
        wb.active['A1'] = 'État à conserver'
        wb.save(self.source)
        wb.close()
        self.app = SimpleNamespace(
            dernier_fichier_excel_ouvert=None,
            _dossier_chantier_selectionne=Mock(side_effect=lambda: self.chantier))
        self.app._memoriser_fichier_excel = (
            lambda chemin, **kwargs: HorizonChantierApp._memoriser_fichier_excel(self.app, chemin, **kwargs))
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.actif = self.stack.enter_context(patch('main._chemin_classeur_excel_actif', return_value=None))
        self.choix = self.stack.enter_context(patch('main.filedialog.askopenfilename', return_value=str(self.source)))
        self.info = self.stack.enter_context(patch('main.messagebox.showinfo'))
        self.erreur = self.stack.enter_context(patch('main.messagebox.showerror'))
        self.racine = self.stack.enter_context(patch('main._chantier_root_for_path', side_effect=lambda _: self.chantier))
        self.export = self.stack.enter_context(patch('main._export_excel_to_pdf', side_effect=self.creer_pdf))

    @staticmethod
    def creer_pdf(source, destination):
        pdf = Canvas(str(destination))
        pdf.drawString(40, 800, source.name)
        pdf.save()

    def sauvegarder(self):
        return HorizonChantierApp.sauvegarde_pdf_historique(self.app)

    def test_sans_excel_ouvert_choisit_et_archive_sans_modifier_le_classeur(self):
        avant = hashlib.sha256(self.source.read_bytes()).digest()
        self.assertTrue(self.sauvegarder())
        self.assertEqual(self.choix.call_args.kwargs['initialdir'], str(self.chantier))
        self.assertEqual(self.app.dernier_fichier_excel_ouvert, self.source)
        pdf = next((self.chantier / 'Historique_Sauvegardes').glob('*.pdf'))
        self.assertIn(self.source.name, PdfReader(pdf).pages[0].extract_text())
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).digest(), avant)
        self.erreur.assert_not_called()
        self.assertTrue(self.sauvegarder())
        self.assertEqual(len(list(pdf.parent.glob('*.pdf'))), 2)
        self.choix.assert_called_once()

    def test_fichier_actif_prioritaire_et_aucun_dialogue_supplementaire(self):
        self.actif.return_value = self.source
        self.app.dernier_fichier_excel_ouvert = self.chantier / 'autre.xlsx'
        self.assertTrue(self.sauvegarder())
        self.export.assert_called_once()
        self.assertEqual(self.export.call_args.args[0], self.source)
        self.choix.assert_not_called()

    def test_dernier_fichier_retrouve_sans_excel_ouvert(self):
        self.app.dernier_fichier_excel_ouvert = self.source
        self.assertTrue(self.sauvegarder())
        self.choix.assert_not_called()

    def test_chantier_deplace_et_ancienne_reference_absente(self):
        self.app.dernier_fichier_excel_ouvert = self.source
        destination = Path(self.tmp.name) / 'Autre partage' / self.chantier.name
        destination.parent.mkdir()
        shutil.move(self.chantier, destination)
        self.chantier = destination
        self.source = destination / 'data' / self.source.name
        self.choix.return_value = str(self.source)
        self.assertTrue(self.sauvegarder())
        self.assertEqual(self.export.call_args.args[1].parent, destination / 'Historique_Sauvegardes')

    def test_annulation_sans_fichier_ni_message_erreur(self):
        self.choix.return_value = ''
        self.assertFalse(self.sauvegarder())
        self.export.assert_not_called()
        self.erreur.assert_not_called()
        self.assertFalse((self.chantier / 'Historique_Sauvegardes').exists())

    def test_selection_hors_chantier_refusee(self):
        autre = Path(self.tmp.name) / 'autre.xlsx'
        shutil.copy2(self.source, autre)
        self.choix.return_value = str(autre)
        self.assertFalse(self.sauvegarder())
        self.export.assert_not_called()
        self.assertIn('chantier sélectionné', self.erreur.call_args.args[1])

    def test_fichier_selectionne_disparu(self):
        self.source.unlink()
        self.assertFalse(self.sauvegarder())
        self.export.assert_not_called()
        self.erreur.assert_called_once()

    def test_echec_excel_ne_signale_pas_une_sauvegarde_reussie(self):
        self.export.side_effect = RuntimeError('Excel indisponible')
        self.assertFalse(self.sauvegarder())
        self.info.assert_not_called()
        self.assertIn('Excel indisponible', self.erreur.call_args.args[1])

    def test_export_ouvre_et_referme_uniquement_le_classeur_necessaire(self):
        for plateforme in ('darwin', 'win32'):
            for deja_ouvert in (False, True):
                with self.subTest(plateforme=plateforme, deja_ouvert=deja_ouvert):
                    livre = SimpleNamespace(fullname=str(self.source), to_pdf=Mock(), close=Mock())
                    xw = SimpleNamespace(apps=[SimpleNamespace(books=[livre])] if deja_ouvert else [],
                                         Book=Mock(return_value=livre))
                    pdf = self.chantier / 'test.pdf'
                    with patch('main.sys.platform', plateforme), patch.dict('sys.modules', {'xlwings': xw}):
                        _export_excel_to_pdf(self.source, pdf)
                    livre.to_pdf.assert_called_once_with(str(pdf))
                    self.assertEqual(livre.close.call_count, int(not deja_ouvert))
                    self.assertEqual(xw.Book.call_count, int(not deja_ouvert))


if __name__ == '__main__':
    unittest.main()
