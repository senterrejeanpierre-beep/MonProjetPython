"""Emplacement des données : commun au projet ou choisi sur chaque poste."""
import json
import os
from pathlib import Path


class EmplacementNonConfigure(FileNotFoundError):
    """Aucun dossier de chantiers n'a encore été choisi sur ce poste."""


def preferences_poste():
    return Path.home() / '.horizon_chantier' / 'emplacement.json'


def trouver_base(preferences=None):
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
    raise EmplacementNonConfigure('Choisissez le dossier des chantiers pour ce poste.')


def memoriser_base(base):
    base = Path(base)
    if not base.is_dir():
        raise FileNotFoundError(base)
    fichier = preferences_poste()
    fichier.parent.mkdir(parents=True, exist_ok=True)
    fichier.write_text(json.dumps({'dossier': str(base)}, ensure_ascii=False), encoding='utf-8')
