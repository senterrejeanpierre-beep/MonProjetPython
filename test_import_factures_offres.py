"""Vérifie les actions d'import avec dialogues simulés et copies réelles."""
import unittest
from contextlib import ExitStack
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock, patch

from main import HorizonChantierApp, DOSSIER_OFFRES_PRIX, DOSSIERS_OFFRES


class ImportFacturesOffresTests(unittest.TestCase):
    def test_bouton_importe_dossier_fournisseur_et_double_clic_le_parcourt(self):
        from main import choisir_fichier_bibliotheque
        for fournisseur in ('Gobert', 'Wienberg'):
            for parcourir in (False, True):
                with self.subTest(fournisseur=fournisseur, parcourir=parcourir), TemporaryDirectory() as tmp, ExitStack() as stack:
                    racine = Path(tmp).resolve()
                    dossier = racine / fournisseur
                    dossier.mkdir()
                    offre = dossier / 'Briques.pdf'
                    offre.write_bytes(b'offre de test')
                    fenetre = MagicMock()
                    fenetre.winfo_screenwidth.return_value = 1440
                    fenetre.winfo_screenheight.return_value = 900
                    stack.enter_context(patch('main.tk.Toplevel', return_value=fenetre))
                    stack.enter_context(patch('main.tk.Menu'))
                    def variable(value=''):
                        var = Mock()
                        var.get.return_value = value
                        var.set.side_effect = lambda valeur: setattr(var.get, 'return_value', valeur)
                        return var
                    stack.enter_context(patch('main.tk.StringVar', side_effect=variable))
                    for widget in ('Frame', 'Label', 'Entry', 'Scrollbar'):
                        stack.enter_context(patch('main.ttk.' + widget))
                    liste = MagicMock()
                    lignes = {}
                    liste.get_children.side_effect = lambda: tuple(lignes)
                    liste.delete.side_effect = lambda identifiant: lignes.pop(identifiant)
                    def inserer(*args, **kwargs):
                        identifiant = str(len(lignes))
                        lignes[identifiant] = kwargs['values']
                        return identifiant
                    liste.insert.side_effect = inserer
                    liste.selection.return_value = ('0',)
                    stack.enter_context(patch('main.ttk.Treeview', return_value=liste))
                    boutons = {}
                    def bouton(*args, **kwargs):
                        boutons[kwargs['text']] = kwargs['command']
                        return Mock()
                    stack.enter_context(patch('main.ttk.Button', side_effect=bouton))
                    def agir():
                        self.assertEqual(lignes['0'][1], 'Dossier')
                        if parcourir:
                            evenements = dict(c.args for c in liste.bind.call_args_list)
                            evenements['<Double-1>']()
                            fenetre.destroy.assert_not_called()
                            self.assertEqual(lignes['0'][1], 'PDF')
                        boutons['Importer dans le chantier']()
                        fenetre.destroy.assert_called_once()
                    fenetre.wait_window.side_effect = agir
                    choix = choisir_fichier_bibliotheque(
                        Mock(), 'Importer dans 004', racine, racine, ('.pdf',), dossiers=True)
                    self.assertEqual(Path(choix), offre if parcourir else dossier)

    def test_dossier_fournisseur_copie_independante_et_conflit(self):
        from main import copier_offre_bibliotheque
        with TemporaryDirectory() as tmp:
            base = Path(tmp)
            source = base / "Generale" / "Domi Plan"
            (source / "Variantes").mkdir(parents=True)
            (source / "Offre.pdf").write_bytes(b"offre")
            (source / "Variantes" / "Plan.pdf").write_bytes(b"plan")
            destination = base / "004" / "Offres_sous_traitants"
            copie = copier_offre_bibliotheque(source, destination)
            self.assertEqual((copie / "Variantes" / "Plan.pdf").read_bytes(), b"plan")
            (copie / "Offre.pdf").unlink()
            self.assertEqual((source / "Offre.pdf").read_bytes(), b"offre")
            seconde = copier_offre_bibliotheque(source, destination)
            self.assertNotEqual(copie, seconde)
            self.assertFalse((copie / "Offre.pdf").exists())
            self.assertEqual((seconde / "Offre.pdf").read_bytes(), b"offre")

    def test_import_pdf_externe_et_document_du_chantier(self):
        for categorie, filtre in (("Factures", None), ("Factures", "Achats_et_sous_traitants"),
                                  *((nom, None) for nom in DOSSIERS_OFFRES)):
            with self.subTest(categorie=categorie, filtre=filtre), TemporaryDirectory() as tmp, ExitStack() as stack:
                base = Path(tmp)
                chantier = base / "Chantier déplacé"
                chantier.mkdir()
                externe = base / "Sous-traitant.pdf"
                externe.write_bytes(b"prix externe")
                interne = chantier / "Devis.pdf"
                interne.write_bytes(b"prix interne")
                app = SimpleNamespace(_nom_chantier_selectionne=lambda: chantier.name,
                                      _dossier_chantier_selectionne=lambda: chantier,
                                      _factures_glisser_deposer=False,
                                      choisir_offre_prix_chantier=Mock(),
                                      choisir_fiche_technique_chantier=Mock())
                fenetre = MagicMock()
                fenetre.winfo_screenwidth.return_value = 1440
                fenetre.winfo_screenheight.return_value = 900
                stack.enter_context(patch('main.tk.Toplevel', return_value=fenetre))
                stack.enter_context(patch('main.tk.Menu'))
                def variable(value=""):
                    var = Mock()
                    var.get.return_value = value
                    var.set.side_effect = lambda valeur: setattr(var.get, 'return_value', valeur)
                    return var
                stack.enter_context(patch('main.tk.StringVar', side_effect=variable))
                for widget in ('Frame', 'Label', 'Entry', 'Treeview', 'Scrollbar', 'Combobox'):
                    stack.enter_context(patch('main.ttk.' + widget))
                boutons = {}
                def bouton(*args, **kwargs):
                    boutons[kwargs['text']] = kwargs['command']
                    return Mock()
                stack.enter_context(patch('main.ttk.Button', side_effect=bouton))
                stack.enter_context(patch('main.filedialog.askopenfilenames', return_value=(str(externe),)))
                stack.enter_context(patch('main.choisir_fichier_a_importer', return_value=interne))
                erreur = stack.enter_context(patch('main.messagebox.showerror'))
                HorizonChantierApp._ouvrir_documents_chantier(app, categorie, factures_type=filtre)
                if categorie in DOSSIERS_OFFRES:
                    self.assertNotIn('Importer un fichier', boutons)
                    self.assertNotIn('Choisir dans ce chantier', boutons)
                    boutons['Ajouter depuis la bibliothèque']()
                    app.choisir_offre_prix_chantier.assert_called_once_with(categorie)
                    self.assertEqual(list((chantier / categorie).iterdir()), [])
                    continue
                boutons['Importer un fichier']()
                boutons['Choisir dans ce chantier']()
                destination = chantier / categorie
                if filtre:
                    destination /= filtre
                self.assertEqual((destination / externe.name).read_bytes(), b"prix externe")
                self.assertEqual((destination / interne.name).read_bytes(), b"prix interne")
                self.assertEqual(externe.read_bytes(), b"prix externe")
                self.assertEqual(interne.read_bytes(), b"prix interne")
                erreur.assert_not_called()

    def test_anciens_acces_ouvrent_uniquement_la_generale(self):
        app = SimpleNamespace(ouvrir_bibliotheque_prix_materiaux=Mock())
        HorizonChantierApp.ouvrir_prix_materiaux_chantier(app)
        HorizonChantierApp._ouvrir_documents_chantier(app, DOSSIER_OFFRES_PRIX)
        self.assertEqual(app.ouvrir_bibliotheque_prix_materiaux.call_count, 2)

    def test_reprise_generale_puis_import_uniquement_sur_choix_utilisateur(self):
        from bibliotheque_prix_materiaux import NOM
        with TemporaryDirectory() as tmp, ExitStack() as stack:
            base = Path(tmp)
            chantiers = base / 'Chantiers'
            sources = []
            for nom in ('001_Existant', '004_Mariemont', '006_Futur'):
                chantier = chantiers / nom
                for origine in ('Demande de prix/Gobert', NOM + '/Wienberg', 'Offres_materiaux'):
                    source = chantier / origine / 'Offre.pdf'
                    source.parent.mkdir(parents=True)
                    source.write_bytes(f'{nom}/{origine}'.encode())
                    sources.append(source)
            avant = {p.relative_to(chantiers): p.read_bytes()
                     for p in chantiers.rglob('*') if p.is_file()}
            stack.enter_context(patch('main.dossier_chantiers', return_value=chantiers))
            erreur = stack.enter_context(patch('main.messagebox.showerror'))
            dialogue_supplementaire = stack.enter_context(patch('main.tk.Toplevel'))
            chantier = chantiers / '004_Mariemont'
            app = SimpleNamespace(_nom_chantier_selectionne=lambda: chantier.name,
                                  _dossier_chantier_selectionne=lambda: chantier,
                                  _ouvrir_documents_chantier=Mock())
            choisir = stack.enter_context(patch('main.choisir_fichier_bibliotheque', return_value=''))
            HorizonChantierApp.choisir_offre_prix_chantier(app)
            generale = base / NOM
            self.assertEqual({p.read_bytes() for p in generale.rglob('*.pdf')}, set(avant.values()))
            self.assertEqual({p.relative_to(chantiers): p.read_bytes()
                              for p in chantiers.rglob('*') if p.is_file()}, avant)
            app._ouvrir_documents_chantier.assert_not_called()

            # Le même parcours fonctionne pour un chantier existant et un futur chantier.
            (chantiers / '007_Nouveau').mkdir()
            for nom in ('004_Mariemont', '007_Nouveau'):
                chantier = chantiers / nom
                for categorie in DOSSIERS_OFFRES:
                    source = generale / 'Gobert/001_Existant/Offre.pdf'
                    def selectionner(*args, **kwargs):
                        self.assertEqual(kwargs['offre_selection']['chantier'], nom)
                        kwargs['offre_selection']['categorie'] = categorie
                        return str(source)
                    choisir.side_effect = selectionner
                    HorizonChantierApp.choisir_offre_prix_chantier(app)
                    copies = list((chantier / categorie).rglob('*.pdf'))
                    self.assertIn(source.read_bytes(), [p.read_bytes() for p in copies])
                    app._ouvrir_documents_chantier.assert_called_with(categorie)
            for relatif, contenu in avant.items():
                self.assertEqual((chantiers / relatif).read_bytes(), contenu)
            dialogue_supplementaire.assert_not_called()
            erreur.assert_not_called()


if __name__ == '__main__':
    unittest.main()
