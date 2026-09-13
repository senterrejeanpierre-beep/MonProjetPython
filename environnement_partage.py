"""Emplacement des données : commun au projet ou choisi sur chaque poste."""
import json
import os
import sys
from pathlib import Path


def dossier_application():
    return Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent


def preferences_poste():
    return Path.home() / '.horizon_chantier' / 'emplacement.json'


def trouver_base(application=None, preferences=None):
    application = Path(application) if application is not None else dossier_application()
    preferences = Path(preferences) if preferences is not None else preferences_poste()
    explicite = os.environ.get('HORIZON_CHANTIER_DATA')
    if explicite:
        base = Path(explicite).expanduser()
        if not base.is_dir():
            raise FileNotFoundError(f'Emplacement des chantiers inaccessible : {base}')
        return base
    if preferences.exists():
        base = Path(json.loads(preferences.read_text(encoding='utf-8'))['dossier'])
        if not base.is_dir():
            raise FileNotFoundError(f'Le dossier choisi est inaccessible : {base}. Vérifiez la connexion au partage.')
        return base
    for parent in (application, application.parent):
        candidate = parent / 'Horizon_Chantier_Data'
        if candidate.is_dir():
            return candidate
        if (parent / 'Chantiers').is_dir() or (parent / 'Chantier').is_dir():
            return parent
    return Path.home() / 'Desktop' / 'Horizon_Chantier_Data'


def memoriser_base(base):
    base = Path(base)
    if not base.is_dir():
        raise FileNotFoundError(base)
    fichier = preferences_poste()
    fichier.parent.mkdir(parents=True, exist_ok=True)
    fichier.write_text(json.dumps({'dossier': str(base)}, ensure_ascii=False), encoding='utf-8')
