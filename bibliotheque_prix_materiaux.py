"""Classe les offres de fournisseurs dans une bibliothèque commune aux chantiers."""

import hashlib
import json
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
    "Electricite", "Outillage_et_consommables", "Autres_materiaux", "Echafaudage",
)
EXTENSIONS = {".pdf", ".xlsx", ".xls", ".xlsm", ".ods", ".doc", ".docx", ".odt", ".eml", ".msg",
              ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".heic", ".webp"}
MOTS = {
    "Echafaudage": ("echafaud",),
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
        if ("demande de prix sous-traitants" not in parties[:-1]
                or not re.search(r"\b(offre|devis)\b", parties[-1])):
            return False
    if (re.match(r"\d+(?:\.\d+)?\.[^ ]+_sou_", parties[-1])
            or any(mot in parties[-1] for mot in ("demande de prix", "demande prix", "cahier des charges",
                                                 "cahier_technique", "cdc ", "metre detaille",
                                                 "soumission", "facture", "fiche technique", "plans_", "plan "))):
        return False
    if any(p in {"archives", "archive", "sauvegarde", "backup", "modele", "modeles", "harchives", "historique_sauvegardes", _normaliser(NOM)}
           for p in parties[:-1]):
        return False
    return True


def trouver_offres_chantier(chantier: Path) -> list[Path]:
    """Recherche uniquement les offres du chantier demandé."""
    chantier = Path(chantier)
    if not chantier.is_dir() or chantier.is_symlink():
        return []
    offres = []
    for dossier in chantier.rglob("*"):
        if (dossier.is_dir() and not dossier.is_symlink()
                and _normaliser(dossier.name) in {"demande de prix", "demande de prix sous-traitants"}
                and not any(_normaliser(p) in {"archives", "archive", "sauvegarde", "backup"}
                            for p in dossier.relative_to(chantier).parts)):
            offres.extend(p for p in dossier.rglob("*") if p.is_file() and est_offre_recue(p))
    return sorted(set(offres), key=lambda p: str(p).casefold())


def actualiser_bibliotheque_chantier(chantier: Path) -> Path:
    """Crée ou complète la bibliothèque locale sans modifier les originaux."""
    chantier = Path(chantier)
    if not chantier.is_dir() or chantier.is_symlink():
        raise ValueError("Dossier du chantier introuvable ou non valide.")
    racine = creer_bibliotheque(chantier)
    importer_offres(trouver_offres_chantier(chantier), chantier.parent, racine)
    return racine


def trouver_offres(chantiers: Path) -> list[Path]:
    from factures_chantier import DOSSIERS_OFFRES
    if not Path(chantiers).is_dir():
        return []
    offres = []
    for chantier in Path(chantiers).iterdir():
        if (not chantier.is_dir() or chantier.is_symlink()
                or _normaliser(chantier.name) in {"modele", "modeles", "archives", "harchives", "sauvegarde", "backup"}):
            continue
        offres.extend(trouver_offres_chantier(chantier))
        for nom in (NOM, *DOSSIERS_OFFRES):
            dossier = chantier / nom
            if dossier.is_dir() and not dossier.is_symlink():
                offres.extend(p for p in dossier.rglob("*")
                              if p.is_file() and p.suffix.casefold() in EXTENSIONS
                              and not any(part.startswith((".", "~$")) for part in p.relative_to(dossier).parts)
                              and not any((chantier / Path(*p.relative_to(chantier).parts[:i])).is_symlink()
                                      for i in range(1, len(p.relative_to(chantier).parts) + 1)))
    return sorted(set(offres), key=lambda p: str(p).casefold())


def _destination_offre(relatif: Path, locale: bool) -> Path:
    """Reprend les dossiers source, avec le chantier sous la rubrique en général."""
    parties = relatif.parts
    indices = [i for i, nom in enumerate(parties[:-1])
               if _normaliser(nom) in {"demande de prix", "demande de prix sous-traitants"}]
    suite = parties[indices[-1] + 1:]
    if len(suite) == 1:
        suite = ("Offres diverses", *suite)
    return Path(*suite) if locale else Path(suite[0], parties[0], *suite[1:])


CONFIG_FOURNISSEURS = ".classement_fournisseurs.json"


def _fournisseur(relatif, configuration):
    valeur = configuration.get(relatif.as_posix(), "")
    if valeur is None:
        return None
    if not valeur:
        ancienne = _destination_offre(relatif, True)
        valeur = ancienne.parts[0]
        generiques = {_normaliser(r).replace("_", " ") for r in RUBRIQUES}
        if (_normaliser(valeur).replace("_", " ") in generiques
                or _normaliser(valeur) in {"offres diverses", "escalier", "peinture"}
                or "terrassement" in _normaliser(valeur)):
            valeur = "Fournisseur à préciser"
    if (not isinstance(valeur, str) or valeur in {".", ".."}
            or any(c in valeur for c in '<>:"/\\|?*') or valeur.endswith((".", " "))):
        raise ValueError("Nom de fournisseur non valide.")
    return valeur


def importer_offres(sources: list[Path], chantiers: Path, bibliotheque: Path) -> tuple[int, int]:
    """Copie selon le rangement source et reprend les anciennes copies identifiées."""
    chantiers = Path(chantiers).resolve()
    bibliotheque = Path(bibliotheque).resolve()
    locale = bibliotheque.parent.parent == chantiers
    config_path = chantiers.parent / NOM / CONFIG_FOURNISSEURS
    configuration = json.loads(config_path.read_text(encoding="utf-8")) if config_path.is_file() else {}
    ajoutes = ignores = 0
    from factures_chantier import DOSSIERS_OFFRES
    empreintes_generales = None
    anciennes_rubriques = {f"{i:02d}_{r}" for i, r in enumerate(RUBRIQUES, 1)}
    anciennes_rubriques.add("04_Bois_et_panneaux")
    for source in sorted(sources, key=lambda p: (len(Path(p).resolve().relative_to(chantiers).parts) >= 3
                         and Path(p).resolve().relative_to(chantiers).parts[1] in (NOM, *DOSSIERS_OFFRES))):
        source = Path(source)
        relatif = source.resolve().relative_to(chantiers)
        # Les offres déjà rangées manuellement sont reprises telles quelles.
        if len(relatif.parts) >= 3 and relatif.parts[1] in (NOM, *DOSSIERS_OFFRES):
            if locale:
                raise ValueError("La reprise des dossiers d'offres cible la bibliothèque générale.")
            if (source.suffix.casefold() not in EXTENSIONS
                    or any((chantiers / Path(*relatif.parts[:i])).is_symlink()
                           for i in range(1, len(relatif.parts) + 1))):
                ignores += 1
                continue
            if empreintes_generales is None:
                empreintes_generales = {hashlib.sha256(p.read_bytes()).digest()
                                       for p in bibliotheque.rglob("*") if p.is_file()
                                       and not any((bibliotheque / Path(*p.relative_to(bibliotheque).parts[:i])).is_symlink()
                                               for i in range(1, len(p.relative_to(bibliotheque).parts) + 1))}
            empreinte = hashlib.sha256(source.read_bytes()).digest()
            if empreinte in empreintes_generales:
                ignores += 1
                continue
            cible = bibliotheque / relatif
            cible.resolve().relative_to(bibliotheque)
            cible.parent.mkdir(parents=True, exist_ok=True)
            initiale = cible
            index = 2
            while any(_normaliser(p.name) == _normaliser(cible.name) for p in cible.parent.iterdir()):
                cible = initiale.with_name(f"{initiale.stem}_{index}{initiale.suffix}")
                index += 1
            with source.open("rb") as entree, cible.open("xb") as sortie:
                shutil.copyfileobj(entree, sortie)
            empreintes_generales.add(empreinte)
            ajoutes += 1
            continue
        if not est_offre_recue(source):
            ignores += 1
            continue
        if locale and relatif.parts[0] != bibliotheque.parent.name:
            raise ValueError("Cette offre appartient à un autre chantier.")
        fournisseur = _fournisseur(relatif, configuration)
        if fournisseur is None:
            ignores += 1
            continue
        empreinte = hashlib.sha256(source.read_bytes()).digest()
        dossier_fournisseur = bibliotheque / fournisseur
        if not locale:
            dossier_fournisseur /= relatif.parts[0]
        cible = dossier_fournisseur / source.name
        ancienne_cible = bibliotheque / _destination_offre(relatif, locale)
        cible.resolve().relative_to(bibliotheque)
        cible.parent.mkdir(parents=True, exist_ok=True)
        # Un fichier renommé par l'utilisateur garde son nom à l'actualisation.
        identique = next((p for p in cible.parent.iterdir()
                          if p.is_file() and not p.is_symlink()
                          and hashlib.sha256(p.read_bytes()).digest() == empreinte), None)
        if identique is not None:
            cible = identique
        initiale = cible
        index = 2
        while True:
            existant = next((p for p in cible.parent.iterdir()
                             if _normaliser(p.name) == _normaliser(cible.name)), None)
            if existant is None:
                shutil.copy2(source, cible)
                if empreintes_generales is not None:
                    empreintes_generales.add(empreinte)
                ajoutes += 1
                break
            if (existant.is_file() and not existant.is_symlink()
                    and hashlib.sha256(existant.read_bytes()).digest() == empreinte):
                cible = existant
                ignores += 1
                break
            cible = initiale.with_name(f"{initiale.stem}_{index}{initiale.suffix}")
            index += 1
        if (ancienne_cible.is_file() and not ancienne_cible.samefile(cible)
                and not ancienne_cible.is_symlink()
                and hashlib.sha256(ancienne_cible.read_bytes()).digest() == empreinte):
            ancienne_cible.unlink()
            parent = ancienne_cible.parent
            while parent != bibliotheque and not any(parent.iterdir()):
                parent.rmdir()
                parent = parent.parent
        # Retirer seulement les anciennes copies identiques de ce même chantier.
        for rubrique in anciennes_rubriques:
            dossier = bibliotheque / rubrique
            if not dossier.is_dir() or dossier.is_symlink():
                continue
            for ancien in dossier.iterdir():
                if (ancien.is_file() and not ancien.is_symlink()
                        and ancien.name.startswith(relatif.parts[0] + "__")
                        and hashlib.sha256(ancien.read_bytes()).digest() == empreinte):
                    ancien.unlink()
    for rubrique in anciennes_rubriques:
        dossier = bibliotheque / rubrique
        if dossier.is_dir() and not dossier.is_symlink() and not any(dossier.iterdir()):
            dossier.rmdir()
    return ajoutes, ignores
