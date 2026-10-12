"""Préparation du dossier de soumission, sans modifier ses documents sources."""
from io import BytesIO
import os
from pathlib import Path
import shutil
import tempfile
from bibliotheque_administrative import est_copie_secours


IMAGES = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
WORD = {".doc", ".docx"}
EXCEL = {".xls", ".xlsx", ".xlsm"}


def documents_du_dossier(dossier):
    """Inclut les pièces courantes, hors sauvegardes, fichiers cachés ou temporaires."""
    dossier = Path(dossier).resolve()
    if not dossier.is_dir():
        raise ValueError("Le dossier administratif est introuvable.")
    fichiers = []

    def erreur_parcours(erreur):
        raise erreur

    for parent, dossiers, noms in os.walk(dossier, onerror=erreur_parcours, followlinks=False):
        dossiers[:] = [n for n in dossiers if not n.startswith((".", "~$"))
                       and not est_copie_secours(Path(n))]
        noms = [n for n in noms if not est_copie_secours(Path(n))]
        for nom in dossiers + [n for n in noms if not n.startswith((".", "~$"))]:
            fichier = Path(parent) / nom
            if fichier.is_symlink():
                raise ValueError(f"Remplacez le lien par une copie du document : {fichier.relative_to(dossier)}")
        fichiers.extend(Path(parent) / n for n in noms if not n.startswith((".", "~$")))
    fichiers.sort(key=lambda p: (p.relative_to(dossier).as_posix().casefold(), p.as_posix()))
    if not fichiers:
        raise ValueError("Le dossier administratif ne contient aucun document.")
    non_convertibles = [p.relative_to(dossier).as_posix() for p in fichiers
                        if p.suffix.lower() not in {".pdf"} | IMAGES | WORD | EXCEL]
    if non_convertibles:
        raise ValueError("Enregistrez ces documents en PDF et remplacez leur version source "
                         "dans le dossier administratif avant de préparer le mail :\n" +
                         "\n".join(non_convertibles))
    return fichiers


def _image_en_pdf(source, destination):
    from PIL import Image, ImageOps, ImageSequence
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen.canvas import Canvas

    pdf = Canvas(str(destination))
    with Image.open(source) as original:
        for page in ImageSequence.Iterator(original):
            page = ImageOps.exif_transpose(page).convert("RGBA")
            fond = Image.new("RGB", page.size, "white")
            fond.paste(page, mask=page.getchannel("A"))
            largeur, hauteur = landscape(A4) if fond.width > fond.height else A4
            pdf.setPageSize((largeur, hauteur))
            rapport = min((largeur - 36) / fond.width, (hauteur - 36) / fond.height)
            w, h = fond.width * rapport, fond.height * rapport
            pdf.drawImage(ImageReader(fond), (largeur - w) / 2, (hauteur - h) / 2, w, h)
            pdf.showPage()
    pdf.save()


def generer_pdf_dossier(dossier, destination, *, convertir_word, convertir_excel):
    """Regroupe toutes les pièces dans l'ordre de la liste, publication atomique.

    Le résultat reste hors d'Administratif pour ne pas se réinclure à l'envoi suivant.
    Toute pièce illisible bloque la génération : aucun dossier incomplet n'est publié.
    """
    from pypdf import PdfReader, PdfWriter

    dossier, destination = Path(dossier).resolve(), Path(destination).resolve()
    if destination.is_relative_to(dossier):
        raise ValueError("Le PDF regroupé doit être enregistré hors du dossier Administratif.")
    fichiers = documents_du_dossier(dossier)
    signatures = {p: (p.stat().st_size, p.stat().st_mtime_ns) for p in fichiers}
    with tempfile.TemporaryDirectory(prefix=".horizon_soumission_", dir=destination.parent) as temporaire:
        travail = Path(temporaire)
        with PdfWriter() as assembleur:
            for numero, source in enumerate(fichiers):
                try:
                    extension = source.suffix.lower()
                    pdf = source
                    if extension != ".pdf":
                        conversion = travail / str(numero)
                        conversion.mkdir()
                        pdf = conversion / "document.pdf"
                        if extension in WORD:
                            copie = conversion / source.name
                            shutil.copy2(source, copie)
                            pdf = convertir_word(copie)
                        elif extension in EXCEL:
                            convertir_excel(source, pdf)
                        else:
                            _image_en_pdf(source, pdf)
                    lecteur = PdfReader(BytesIO(Path(pdf).read_bytes()))
                    if lecteur.is_encrypted and not lecteur.decrypt(""):
                        raise ValueError("PDF protégé par un mot de passe")
                    if not lecteur.pages:
                        raise ValueError("PDF sans page")
                    assembleur.append(lecteur, outline_item=source.relative_to(dossier).as_posix())
                except Exception as erreur:
                    raise ValueError(f"Impossible d'inclure {source.relative_to(dossier)} :\n{erreur}") from erreur
            if (documents_du_dossier(dossier) != fichiers or
                    any((p.stat().st_size, p.stat().st_mtime_ns) != signatures[p] for p in fichiers)):
                raise ValueError("Le dossier a changé pendant la préparation. Relancez la préparation du mail.")
            resultat = travail / "dossier.pdf"
            assembleur.write(resultat)
        # Réserver le nom : un autre poste peut préparer un envoi simultanément.
        with destination.open("xb"):
            pass
        try:
            os.replace(resultat, destination)
        except Exception:
            destination.unlink(missing_ok=True)
            raise
    return len(fichiers)
