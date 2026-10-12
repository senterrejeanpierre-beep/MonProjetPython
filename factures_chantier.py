"""Classement commun des factures de chaque chantier."""

from pathlib import Path


DOSSIERS_OFFRES = {
    "Offres_sous_traitants": "Sous-traitants",
    "Offres_materiaux": "Matériaux",
    "Offres_location_materiel": "Location de matériel",
}

DOSSIERS_FACTURES = ("Entreprise", "Achats_et_sous_traitants", "Achats_materiaux")


def assurer_dossiers_factures(dossier_chantier: Path) -> Path:
    """Crée seulement la structure ; aucun document existant n'est déplacé."""
    dossier_chantier = Path(dossier_chantier)
    if not dossier_chantier.is_dir() or dossier_chantier.is_symlink():
        raise ValueError(f"Dossier de chantier indisponible : {dossier_chantier}")
    factures = dossier_chantier / "Factures"
    if factures.is_symlink() or (factures.exists() and not factures.is_dir()):
        raise ValueError(f"Dossier Factures non standard : {factures}")
    factures.mkdir(exist_ok=True)
    for nom in DOSSIERS_FACTURES:
        sous_dossier = factures / nom
        if sous_dossier.is_symlink() or (sous_dossier.exists() and not sous_dossier.is_dir()):
            raise ValueError(f"Dossier de factures non standard : {sous_dossier}")
        sous_dossier.mkdir(exist_ok=True)
    for nom in DOSSIERS_OFFRES:
        offres = dossier_chantier / nom
        if offres.is_symlink() or (offres.exists() and not offres.is_dir()):
            raise ValueError(f"Dossier d'offres non standard : {offres}")
        offres.mkdir(exist_ok=True)
    return factures


def preparer_dossiers_factures(base: Path, modele: Path | None = None) -> tuple[list[Path], list[str]]:
    """Prépare les chantiers enregistrés et le modèle sans toucher aux pièces."""
    base = Path(base)
    dossiers = [base / fichier.stem for fichier in sorted(base.glob("*.json"))]
    if modele is not None:
        dossiers.append(Path(modele))
    prepares, erreurs = [], []
    for dossier in dossiers:
        try:
            prepares.append(assurer_dossiers_factures(dossier))
        except (OSError, ValueError) as erreur:
            erreurs.append(str(erreur))
    return prepares, erreurs


def categorie_offre(chemin: Path) -> str | None:
    """Classe seulement les familles explicitement confirmées par l'utilisateur."""
    import re
    import unicodedata
    texte = unicodedata.normalize("NFKD", str(chemin).casefold())
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    mots = set(re.findall(r"[a-z0-9]+", texte))
    if mots & {"echafaudage", "echafaudages", "essai", "essais"} or "stone assistance" in texte:
        return "Offres_sous_traitants"
    if mots & {"lietar", "loiselet"}:
        return "Offres_location_materiel"
    if (mots & {"vegetal", "vegetaux", "sedum"}
            or {"couvre", "mur"} <= mots or {"couvre", "murs"} <= mots):
        return "Offres_materiaux"
    return None


def classer_offres_confirmees(chantier: Path) -> int:
    """Copie les offres reconnues du seul chantier, sans écrasement ni doublon."""
    import hashlib
    import shutil
    from bibliotheque_prix_materiaux import NOM, EXTENSIONS, trouver_offres_chantier

    chantier = Path(chantier)
    assurer_dossiers_factures(chantier)
    sources = set(trouver_offres_chantier(chantier))
    ancienne = chantier / NOM
    if ancienne.is_dir() and not ancienne.is_symlink():
        sources.update(p for p in ancienne.rglob("*")
                       if p.is_file() and p.suffix.casefold() in EXTENSIONS)
    empreintes = {nom: {hashlib.sha256(p.read_bytes()).digest()
                       for p in (chantier / nom).rglob("*")
                       if p.is_file() and not p.is_symlink()}
                 for nom in DOSSIERS_OFFRES}
    ajouts = 0
    for source in sorted(sources):
        relatif = source.relative_to(chantier)
        if (source.is_symlink() or any((chantier / Path(*relatif.parts[:i])).is_symlink()
                                      for i in range(1, len(relatif.parts)))
                or any(p.startswith((".", "~$")) for p in relatif.parts)):
            continue
        categorie = categorie_offre(relatif)
        if categorie is None:
            continue
        empreinte = hashlib.sha256(source.read_bytes()).digest()
        if empreinte in empreintes[categorie]:
            continue
        cible = chantier / categorie / source.name
        numero = 2
        while cible.exists():
            cible = chantier / categorie / f"{source.stem}_{numero}{source.suffix}"
            numero += 1
        cree = False
        try:
            with source.open("rb") as entree, cible.open("xb") as sortie:
                cree = True
                shutil.copyfileobj(entree, sortie)
        except Exception:
            if cree:
                cible.unlink(missing_ok=True)
            raise
        empreintes[categorie].add(empreinte)
        ajouts += 1
    return ajouts
