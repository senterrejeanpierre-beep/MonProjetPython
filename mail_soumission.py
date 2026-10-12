"""Ouvre Gmail et présente le PDF à joindre, sans envoyer de message."""
from pathlib import Path
import subprocess
import sys
from urllib.parse import urlencode
import webbrowser


def ouvrir_gmail_soumission(pdf, nom_chantier):
    pdf = Path(pdf).resolve()
    if not pdf.is_file():
        raise ValueError("Le PDF du dossier administratif est introuvable.")
    # Sélectionner le fichier permet de le glisser directement dans Gmail.
    if sys.platform == "darwin":
        subprocess.run(["open", "-R", str(pdf)], check=True, capture_output=True, text=True)
    elif sys.platform == "win32":
        subprocess.Popen(["explorer.exe", "/select,", str(pdf)])
    else:
        subprocess.Popen(["xdg-open", str(pdf.parent)])
    parametres = urlencode({
        "view": "cm", "fs": "1",
        "su": f"Soumission — {nom_chantier} — Dossier administratif",
        "body": ("Bonjour,\n\nVeuillez trouver en pièce jointe le dossier administratif "
                 f"pour la soumission du chantier {nom_chantier}.\n\nCordialement,"),
    })
    if not webbrowser.open("https://mail.google.com/mail/?" + parametres, new=2):
        raise RuntimeError("Le navigateur n'a pas pu ouvrir Gmail. "
                           "Ouvrez Gmail et joignez le PDF préparé.")
