"""PDF réels ; Word, Excel, navigateur et commandes système simulés."""
from contextlib import ExitStack
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, Mock, patch
from urllib.parse import parse_qs, urlparse

from main import HorizonChantierApp
from dossier_administratif_pdf import generer_pdf_dossier
from mail_soumission import ouvrir_gmail_soumission
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen.canvas import Canvas
from PIL import Image


def creer_pdf(chemin, *textes):
    chemin.parent.mkdir(parents=True, exist_ok=True)
    c = Canvas(str(chemin))
    for texte in textes:
        c.drawString(50, 700, texte)
        c.showPage()
    c.save()


class DossierAdministratifTests(unittest.TestCase):
    def generer(self, dossier, destination, **options):
        return generer_pdf_dossier(
            dossier, destination,
            convertir_word=options.get("word", Mock(side_effect=AssertionError("Word inattendu"))),
            convertir_excel=options.get("excel", Mock(side_effect=AssertionError("Excel inattendu"))))

    def test_dossier_complet_sous_dossiers_images_et_deplacement(self):
        with TemporaryDirectory() as tmp:
            for nom in ("Ancien chantier", "Nouveau chantier déplacé"):
                with self.subTest(chantier=nom):
                    chantier = Path(tmp) / nom
                    dossier = chantier / "Administratif"
                    creer_pdf(dossier / "01 Offre.pdf", "OFFRE page 1", "OFFRE page 2")
                    creer_pdf(dossier / "02 References" / "preuve.PDF", "REFERENCE")
                    Image.new("RGB", (180, 100), "red").save(dossier / "03 Photo.png")
                    (dossier / ".DS_Store").write_bytes(b"ignore")
                    (dossier / "~$offre.docx").write_bytes(b"ignore")
                    (dossier / ".cache").mkdir()
                    (dossier / ".cache" / "invalide.pdf").write_bytes(b"ignore")
                    avant = {p.relative_to(dossier): p.read_bytes() for p in dossier.rglob("*") if p.is_file()}
                    destination = chantier / "soumission.pdf"
                    self.assertEqual(self.generer(dossier, destination), 3)
                    pdf = PdfReader(destination)
                    self.assertEqual(len(pdf.pages), 4)
                    self.assertEqual([p.extract_text().strip() for p in pdf.pages[:3]],
                                     ["OFFRE page 1", "OFFRE page 2", "REFERENCE"])
                    self.assertEqual(len(pdf.pages[3].images), 1)
                    self.assertEqual(avant, {p.relative_to(dossier): p.read_bytes()
                                            for p in dossier.rglob("*") if p.is_file()})
                    self.assertEqual(self.generer(dossier, chantier / "second.pdf"), 3)
                    self.assertEqual(len(PdfReader(chantier / "second.pdf").pages), 4)

    def test_word_converti_sur_copie_et_excel_depuis_source(self):
        with TemporaryDirectory() as tmp:
            dossier = Path(tmp) / "Administratif"
            dossier.mkdir()
            word = dossier / "01 Offre.docx"
            excel = dossier / "02 Liste.xlsx"
            word.write_bytes(b"word original")
            excel.write_bytes(b"excel original")

            def conversion_word(source):
                self.assertNotEqual(source, word)
                self.assertEqual(source.read_bytes(), word.read_bytes())
                source.write_bytes(b"simulation modification Word")
                creer_pdf(source.with_suffix(".pdf"), "WORD")
                return source.with_suffix(".pdf")

            def conversion_excel(source, cible):
                self.assertEqual(source, excel.resolve())
                creer_pdf(cible, "EXCEL")

            resultat = Path(tmp) / "resultat.pdf"
            self.generer(dossier, resultat, word=conversion_word, excel=conversion_excel)
            self.assertEqual([p.extract_text().strip() for p in PdfReader(resultat).pages], ["WORD", "EXCEL"])
            self.assertEqual(word.read_bytes(), b"word original")
            self.assertEqual(excel.read_bytes(), b"excel original")
            self.assertEqual(sorted(p.name for p in dossier.iterdir()), [word.name, excel.name])

    def test_aucun_pdf_incomplet_sur_piece_invalide_ou_non_convertible(self):
        for nom, contenu in (("preuve.pdf", b"%PDF-1.4\nincomplet"), ("piece.odt", b"ODT")):
            with self.subTest(piece=nom), TemporaryDirectory() as tmp:
                dossier = Path(tmp) / "Administratif"
                creer_pdf(dossier / "01 Offre.pdf", "OFFRE")
                (dossier / nom).write_bytes(contenu)
                destination = Path(tmp) / "resultat.pdf"
                with self.assertRaisesRegex(ValueError, nom):
                    self.generer(dossier, destination)
                self.assertFalse(destination.exists())
                self.assertFalse(list(Path(tmp).glob(".horizon_soumission_*")))

    def test_pdf_protege_bloque_le_dossier_complet(self):
        with TemporaryDirectory() as tmp:
            dossier = Path(tmp) / "Administratif"
            dossier.mkdir()
            with PdfWriter() as pdf:
                pdf.add_blank_page(width=100, height=100)
                pdf.encrypt("secret")
                pdf.write(dossier / "protege.pdf")
            with self.assertRaisesRegex(ValueError, "mot de passe"):
                self.generer(dossier, Path(tmp) / "resultat.pdf")

    def test_aucun_ecrasement_et_aucune_reinclusion(self):
        with TemporaryDirectory() as tmp:
            dossier = Path(tmp) / "Administratif"
            creer_pdf(dossier / "offre.pdf", "OFFRE")
            destination = Path(tmp) / "existant.pdf"
            destination.write_bytes(b"precedent envoi")
            with self.assertRaises(FileExistsError):
                self.generer(dossier, destination)
            self.assertEqual(destination.read_bytes(), b"precedent envoi")
            with self.assertRaisesRegex(ValueError, "hors du dossier"):
                self.generer(dossier, dossier / "resultat.pdf")

    def test_dossier_vide_et_modification_pendant_conversion(self):
        with TemporaryDirectory() as tmp:
            dossier = Path(tmp) / "Administratif"
            dossier.mkdir()
            destination = Path(tmp) / "resultat.pdf"
            with self.assertRaisesRegex(ValueError, "aucun document"):
                self.generer(dossier, destination)
            source = dossier / "offre.docx"
            source.write_bytes(b"original")

            def conversion(copie):
                source.write_bytes(b"nouvelle version")
                creer_pdf(copie.with_suffix(".pdf"), "ancienne version")
                return copie.with_suffix(".pdf")

            with self.assertRaisesRegex(ValueError, "changé"):
                self.generer(dossier, destination, word=conversion)
            self.assertFalse(destination.exists())

    def test_gmail_objet_corps_accents_et_selection_mac_windows(self):
        with TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "Dossier soumission é.pdf"
            creer_pdf(pdf, "OFFRE")
            for plateforme in ("darwin", "win32"):
                with self.subTest(plateforme=plateforme), patch("mail_soumission.sys.platform", plateforme), \
                        patch("mail_soumission.subprocess.run") as run, \
                        patch("mail_soumission.subprocess.Popen") as popen, \
                        patch("mail_soumission.webbrowser.open", return_value=True) as navigateur:
                    nom = "École & bâtiment #2"
                    ouvrir_gmail_soumission(pdf, nom)
                    url = urlparse(navigateur.call_args.args[0])
                    self.assertEqual(url.netloc, "mail.google.com")
                    donnees = parse_qs(url.query)
                    self.assertIn(nom, donnees["su"][0])
                    self.assertIn(nom, donnees["body"][0])
                    self.assertNotIn("to", donnees)
                    if plateforme == "darwin":
                        self.assertEqual(run.call_args.args[0], ["open", "-R", str(pdf.resolve())])
                    else:
                        self.assertEqual(popen.call_args.args[0], ["explorer.exe", "/select,", str(pdf.resolve())])

    def test_bouton_admin_regroupe_tout_meme_si_recherche_active(self):
        with TemporaryDirectory() as tmp, ExitStack() as stack:
            chantier = Path(tmp) / "Chantier déplacé"
            dossier = chantier / "Administratif"
            creer_pdf(dossier / "Offre.pdf", "OFFRE")
            creer_pdf(dossier / "Attestation.pdf", "ATTESTATION")
            secours = dossier / 'Liste de références' / 'Copie de secours de Liste de références.docx'
            secours.parent.mkdir()
            secours.write_bytes(b'sauvegarde Word a ne pas envoyer')
            archive = dossier / 'Copies_de_secours' / secours.relative_to(dossier)
            app = SimpleNamespace(_nom_chantier_selectionne=lambda: chantier.name,
                                  _dossier_chantier_selectionne=lambda: chantier,
                                  choisir_document_administratif_chantier=Mock())
            fenetre = MagicMock()
            fenetre.winfo_screenwidth.return_value = 1440
            fenetre.winfo_screenheight.return_value = 900
            stack.enter_context(patch("main.tk.Toplevel", return_value=fenetre))
            stack.enter_context(patch("main.tk.Menu"))
            variables = []

            def variable(value=""):
                var = Mock()
                var.get.return_value = value
                var.set.side_effect = lambda val: setattr(var.get, "return_value", val)
                variables.append(var)
                return var

            stack.enter_context(patch("main.tk.StringVar", side_effect=variable))
            for widget in ("Frame", "Label", "Entry", "Scrollbar"):
                stack.enter_context(patch("main.ttk." + widget))
            liste = stack.enter_context(patch("main.ttk.Treeview")).return_value
            boutons = {}

            def bouton(*args, **kwargs):
                boutons[kwargs["text"]] = kwargs["command"]
                return Mock()

            stack.enter_context(patch("main.ttk.Button", side_effect=bouton))
            stack.enter_context(patch("main.threading.Thread", side_effect=lambda target, **_: SimpleNamespace(start=target)))
            erreur = stack.enter_context(patch("main.messagebox.showerror"))
            avertissement = stack.enter_context(patch("main.messagebox.showwarning"))
            gmail = stack.enter_context(patch("mail_soumission.ouvrir_gmail_soumission"))
            HorizonChantierApp._ouvrir_documents_chantier(app, "Administratif")
            self.assertFalse(secours.exists())
            self.assertEqual(archive.read_bytes(), b'sauvegarde Word a ne pas envoyer')
            self.assertEqual({c.kwargs['values'][0] for c in liste.insert.call_args_list},
                             {'Attestation.pdf', 'Offre.pdf'})
            variables[2].set("Offre")
            for i in range(2):
                if i == 1:
                    # Le deuxième clic simule Excel dans le fil de préparation Windows.
                    (dossier / "Liste.xlsx").write_bytes(b"excel")
                    com = MagicMock()
                    stack.enter_context(patch.dict("sys.modules", {"pythoncom": com}))
                    stack.enter_context(patch("main.sys.platform", "win32"))
                    stack.enter_context(patch("main._export_excel_to_pdf",
                                              side_effect=lambda source, pdf: creer_pdf(pdf, "LISTE")))
                boutons["Envoyer le dossier par mail"]()
                fenetre.after.call_args.args[1]()
                pdf, nom = gmail.call_args.args
                self.assertEqual(len(PdfReader(pdf).pages), 2 + i)
                self.assertEqual(nom, chantier.name)
                if i == 1:
                    com.CoInitialize.assert_called_once_with()
                    com.CoUninitialize.assert_called_once_with()
            self.assertEqual(len(list(chantier.glob("Dossier_administratif*.pdf"))), 2)
            erreur.assert_not_called()
            avertissement.assert_not_called()

    def test_export_ignore_secours_ranges_ou_encore_disperses(self):
        with TemporaryDirectory() as temp:
            dossier = Path(temp) / 'Administratif'
            creer_pdf(dossier / 'Offre.pdf', 'OFFRE COURANTE')
            for relatif in ('Copie de secours de Offre.docx', 'Liste.WBK',
                            'Copies_de_secours/Liste.docx',
                            'Ancien/Copie de secours de Liste.docx'):
                fichier = dossier / relatif
                fichier.parent.mkdir(parents=True, exist_ok=True)
                fichier.write_bytes(b'sauvegarde non convertible')
            avant = {p: p.read_bytes() for p in dossier.rglob('*') if p.is_file()}
            resultat = Path(temp) / 'envoi.pdf'
            self.assertEqual(self.generer(dossier, resultat), 1)
            self.assertEqual(PdfReader(resultat).pages[0].extract_text().strip(), 'OFFRE COURANTE')
            self.assertEqual({p: p.read_bytes() for p in dossier.rglob('*') if p.is_file()}, avant)


if __name__ == "__main__":
    unittest.main()
