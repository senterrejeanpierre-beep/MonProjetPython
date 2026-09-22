"""Classement commun des factures de chaque chantier."""

from pathlib import Path


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
