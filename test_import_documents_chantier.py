import tempfile
import json
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from main import copier_fichier_importe, convertir_word_en_pdf_chantier, HorizonChantierApp


class ImportDocumentsChantierTests(unittest.TestCase):
    def test_metre_ouvre_original_configure_sans_choix_parmi_versions_ou_lots(self):
        with tempfile.TemporaryDirectory() as temporaire:
            for nom_chantier, original in (
                    ('004_Mariemont', '2420_MARIEMONT_METRE RECAP_LOT 1-MAC_250924.xlsx'),
                    ('006_Nouveau chantier déplacé', 'Métré Lot 4 dernière publication.xlsx')):
                dossier = Path(temporaire) / nom_chantier
                dossier.mkdir()
                copie = Path(original).stem + '_avancement.xlsx'
                for nom in (original, copie, 'Metre_detaille.pdf',
                            '2420_MARIEMONT_METRE RECAP_LOT 1-MAC_250917.xlsx',
                            '2420_MARIEMONT_METRE RECAP_LOT 2-BETON_250924.xlsx',
                            '2420_MARIEMONT_METRE RECAP_LOT 3-FER_250917.xlsx'):
                    (dossier / nom).write_bytes(nom.encode())
                (dossier / 'bordereau_public.json').write_text(
                    json.dumps({'original': original, 'copie': copie}), encoding='utf-8')
                avant = {p.name: p.read_bytes() for p in dossier.iterdir()}
                app = SimpleNamespace(_nom_chantier_selectionne=lambda: nom_chantier,
                                      _dossier_chantier_selectionne=lambda: dossier,
                                      _preparer_ouverture_document=lambda p: True,
                                      _doit_rappeler_pdf_historique=lambda p: False)
                app._ouvrir_fichier_chantier = lambda nom: HorizonChantierApp._ouvrir_fichier_chantier(app, nom)
                with patch('main.ouvrir_chemin') as ouvrir, patch('main.messagebox.showerror') as erreur:
                    HorizonChantierApp.ouvrir_metre_original(app)
                    ouvrir.assert_called_once_with(dossier / original)
                    erreur.assert_not_called()
                self.assertEqual({p.name: p.read_bytes() for p in dossier.iterdir()}, avant)
                # Un original manquant ne doit jamais ouvrir un autre lot ou une ancienne version.
                (dossier / original).unlink()
                with patch('main.ouvrir_chemin') as ouvrir, patch('main.messagebox.showerror') as erreur:
                    HorizonChantierApp.ouvrir_metre_original(app)
                    ouvrir.assert_not_called()
                    erreur.assert_called_once()

    def test_metre_sans_configuration_conserve_acces_pdf(self):
        with tempfile.TemporaryDirectory() as temporaire:
            app = SimpleNamespace(_nom_chantier_selectionne=lambda: 'Chantier',
                                  _dossier_chantier_selectionne=lambda: Path(temporaire),
                                  _ouvrir_fichier_chantier=Mock())
            HorizonChantierApp.ouvrir_metre_original(app)
            app._ouvrir_fichier_chantier.assert_called_once_with('Metre_detaille.pdf')

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
