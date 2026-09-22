"""Classe les offres de fournisseurs dans une bibliothèque commune aux chantiers."""

import hashlib
import re
import shutil
import unicodedata
from pathlib import Path


NOM = "Bibliotheque_Prix_Materiaux"
RUBRIQUES = (
    "Briques", "Chaux", "Acier_et_metaux", "Bois", "Pierre",
    "Beton_et_granulats", "Ciment", "Mortiers", "Sable", "Isolants",
    "Toiture_et_ardoises", "Zinc_et_gouttieres", "Etancheite", "Peintures",
    "Enduits", "Menuiserie", "Vitrage", "Quincaillerie", "Tuyaux_et_drainage",
    "Electricite", "Outillage_et_consommables", "Autres_materiaux",
)
EXTENSIONS = {".pdf", ".xlsx", ".xls", ".xlsm", ".ods", ".doc", ".docx", ".odt", ".eml", ".msg"}
MOTS = {
    "Briques": ("brique", "terre cuite"), "Chaux": ("chaux",),
    "Acier_et_metaux": ("acier", "ferraill", "metal", "armature", "inox"),
    "Bois": ("bois", "chene", "osb", "panneau", "charpente"),
    "Pierre": ("pierre", "marbre", "granit"),
    "Beton_et_granulats": ("beton", "gravier", "granulat"),
    "Ciment": ("ciment",), "Mortiers": ("mortier", "jointoi"),
    "Sable": ("sable",), "Isolants": ("isolant", "laine de roche", "foamglas"),
    "Toiture_et_ardoises": ("toiture", "ardoise", "tuile", "couverture"),
    "Zinc_et_gouttieres": ("zinc", "gouttiere"),
    "Etancheite": ("etanche", "membrane"), "Peintures": ("peinture", "lasure"),
    "Enduits": ("enduit", "platre"), "Menuiserie": ("menuiserie", "chassis", "porte"),
    "Vitrage": ("vitrage", "verre", "vitre"),
    "Quincaillerie": ("quincaillerie", "visserie", "fixation"),
    "Tuyaux_et_drainage": ("tuyau", "drain", "pvc"),
    "Electricite": ("electri", "cable"),
    "Outillage_et_consommables": ("outillage", "outil", "consommable"),
}


def _normaliser(texte):
    return "".join(c for c in unicodedata.normalize("NFKD", texte.casefold())
                   if not unicodedata.combining(c))


def creer_bibliotheque(base: Path) -> Path:
    racine = Path(base) / NOM
    racine.mkdir(exist_ok=True)
    ancien_bois = racine / "04_Bois_et_panneaux"
    bois = racine / "04_Bois"
    if ancien_bois.is_dir() and not bois.exists():
        ancien_bois.rename(bois)
    for numero, rubrique in enumerate(RUBRIQUES, 1):
        (racine / f"{numero:02d}_{rubrique}").mkdir(exist_ok=True)
    return racine


def categorie(chemin: Path) -> str:
    texte = _normaliser(" ".join(Path(chemin).parts))
    for rubrique, mots in MOTS.items():
        if any(mot in texte for mot in mots):
            return rubrique
    return "Autres_materiaux"


def est_offre_recue(chemin: Path) -> bool:
    """Retient les réponses rangées dans Demande de prix, pas les demandes ni les pièces du marché."""
    chemin = Path(chemin)
    if chemin.suffix.casefold() not in EXTENSIONS or chemin.is_symlink():
        return False
    parties = [_normaliser(p) for p in chemin.parts]
    if any(p.startswith((".", "~$")) for p in chemin.parts):
        return False
    if "demande de prix" not in parties[:-1]:
        return False
    if (re.match(r"\d+(?:\.\d+)?\.[^ ]+_sou_", parties[-1])
            or any(mot in parties[-1] for mot in ("demande de prix", "demande prix", "cahier des charges",
                                                 "cahier_technique", "cdc ", "metre detaille",
                                                 "soumission", "facture", "fiche technique", "plans_", "plan "))):
        return False
    if any(p in {"archives", "archive", "sauvegarde", "backup", "modeles", "modèles"}
           for p in parties[:-1]):
        return False
    return True


def trouver_offres(chantiers: Path) -> list[Path]:
    if not Path(chantiers).is_dir():
        return []
    offres = []
    for chantier in Path(chantiers).iterdir():
        if not chantier.is_dir() or chantier.is_symlink():
            continue
        for dossier in chantier.rglob("*"):
            if (dossier.is_dir() and not dossier.is_symlink()
                    and _normaliser(dossier.name) == "demande de prix"
                    and not any(_normaliser(p) in {"archives", "archive", "sauvegarde", "backup"}
                                for p in dossier.relative_to(chantier).parts)):
                offres.extend(p for p in dossier.rglob("*") if p.is_file() and est_offre_recue(p))
    return sorted(set(offres), key=lambda p: str(p).casefold())


def importer_offres(sources: list[Path], chantiers: Path, bibliotheque: Path) -> tuple[int, int]:
    """Copie les originaux, conserve le nom du chantier et évite les doublons binaires."""
    chantiers = Path(chantiers).resolve()
    bibliotheque = Path(bibliotheque).resolve()
    ajoutes = ignores = 0
    empreintes = set()
    for ancien in bibliotheque.rglob("*"):
        if ancien.is_file() and not ancien.is_symlink():
            empreintes.add(hashlib.sha256(ancien.read_bytes()).digest())
    for source in sources:
        source = Path(source)
        relatif = source.resolve().relative_to(chantiers)
        if not est_offre_recue(source):
            ignores += 1
            continue
        empreinte = hashlib.sha256(source.read_bytes()).digest()
        if empreinte in empreintes:
            ignores += 1
            continue
        rubrique = categorie(relatif)
        numero = RUBRIQUES.index(rubrique) + 1
        dossier = bibliotheque / f"{numero:02d}_{rubrique}"
        nom = "__".join((relatif.parts[0], *relatif.parts[2:]))
        nom = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nom).rstrip(" .")
        if len(nom) > 180:
            nom = nom[:180-len(source.suffix)].rstrip(" .") + source.suffix
        cible = dossier / nom
        index = 2
        while any(p.name.casefold() == cible.name.casefold() for p in dossier.iterdir()):
            cible = dossier / f"{Path(nom).stem}_{index}{source.suffix}"
            index += 1
        shutil.copy2(source, cible)
        empreintes.add(empreinte)
        ajoutes += 1
    return ajoutes, ignores
