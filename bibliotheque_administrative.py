"""Bibliothèque administrative commune, plate et portable entre macOS et Windows."""

import hashlib
import re
import shutil
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile


DOSSIER_BIBLIOTHEQUE = "Bibliotheque_administrative"
DOSSIER_COMMUN = "Document prêt/Dossier à envoyer/JT BATI"
DOSSIERS_COMMUNS = (
    "Liste du personnel et diplôme",
    "Liste de référence JT BATI copie",
    "Méthodologie",
    "Document prêt/Chiffre d'affaire",
    "Document prêt/Agréation",
    "Document prêt/Liste de référence",
    "Document prêt/Méthodologies",
    "Document prêt/Attestation B sous-traitant",
    "Document prêt/Attestation C engagement des tiers",
    "Document prêt/Attestation D dumping social",
    DOSSIER_COMMUN,
)
EXCLUS = {".ds_store", "thumbs.db"}
MODIFIABLES = {".doc", ".docx", ".odt", ".xls", ".xlsx", ".xlsm", ".ods"}
DOSSIERS_CATEGORIES = (
    "01_Personnel_et_diplomes",
    "02_References_et_attestations",
    "03_Methodologie_et_reportages",
    "04_Agreation_et_chiffre_affaires",
    "05_Securite_et_autres_documents",
)


def _nom_normalise(nom: str) -> str:
    return unicodedata.normalize("NFC", nom).casefold()


def _sans_accents(nom: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", nom.casefold())
                   if not unicodedata.combining(c))


def categorie_document(fichier: Path) -> str:
    """Classe une pièce commune selon son usage, quel que soit son format."""
    nom = _sans_accents(Path(fichier).name.removeprefix("COMMUN__"))
    if any(mot in nom for mot in ("personnel", "diplom", "joffrey", "jouannet", "loic", "lemaire",
                                  "renauld", "masset", "malori", "pector", "timy", "trinchesse", "annexe")):
        return DOSSIERS_CATEGORIES[0]
    if any(mot in nom for mot in ("reference", "attestation ferme", "attestation st roch",
                                  "attestation-des-travaux", "attestation chateau", "attestation chapelle",
                                  "attestation acienne", "attestation sig")):
        return DOSSIERS_CATEGORIES[1]
    if any(mot in nom for mot in ("methodologie", "reportage", "photograph", "mortier", "boutisse",
                                  "installation de chantier", "instalation de chantier")):
        return DOSSIERS_CATEGORIES[2]
    if any(mot in nom for mot in ("agreation", "chiffre d'affaire", "compte 70", "compte 41",
                                  "chiffres en lettres")) or re.match(r"^20\d\d-\d+", nom):
        return DOSSIERS_CATEGORIES[3]
    return DOSSIERS_CATEGORIES[4]


def _nom_portable(nom: str) -> str:
    nom = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nom).rstrip(" .")
    if Path(nom).stem.upper() in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(?:COM|LPT)[1-9]", Path(nom).stem.upper()):
        nom = "_" + nom
    return nom or "Document"


def _chemin_existant(base: Path, relatif: str) -> Path | None:
    courant = base
    for morceau in Path(relatif).parts:
        if not courant.is_dir():
            return None
        courant = next((p for p in courant.iterdir() if _nom_normalise(p.name) == _nom_normalise(morceau)), None)
        if courant is None:
            return None
    return courant


def _fichiers_valides(dossier: Path):
    for fichier in dossier.rglob("*"):
        if fichier.is_file() and not fichier.is_symlink() and not any(
            part.startswith(".") or part.startswith("~$") or part.casefold() in EXCLUS
            for part in fichier.relative_to(dossier).parts
        ):
            yield fichier


def _copier_sans_remplacer(source: Path, destination: Path, nom: str) -> tuple[Path, bool]:
    nom = _nom_portable(nom)
    cible = destination / nom
    numero = 2
    while (existante := _chemin_existant(destination, cible.name)) is not None:
        if existante.is_file() and hashlib.sha256(existante.read_bytes()).digest() == hashlib.sha256(source.read_bytes()).digest():
            return existante, False
        cible = destination / f"{Path(nom).stem}_{numero}{Path(nom).suffix}"
        numero += 1
    cree = False
    try:
        with source.open("rb") as entree, cible.open("xb") as sortie:
            cree = True
            shutil.copyfileobj(entree, sortie)
        shutil.copystat(source, cible)
    except Exception:
        if cree:
            cible.unlink(missing_ok=True)
        raise
    return cible, True


def ajouter_pieces_communes(bibliotheque: Path, source: Path, dossiers: tuple[str, ...]) -> int:
    """Ajoute les documents réutilisables dans cinq dossiers."""
    bibliotheque = Path(bibliotheque)
    for nom in DOSSIERS_CATEGORIES:
        (bibliotheque / nom).mkdir(exist_ok=True)
    ajoutes = 0
    for relatif in dossiers:
        dossier = _chemin_existant(Path(source), relatif)
        if dossier is None or not dossier.is_dir():
            continue
        prefixe = "COMMUN__" if relatif == DOSSIER_COMMUN else ""
        for fichier in _fichiers_valides(dossier):
            nom = prefixe + fichier.name
            categorie = bibliotheque / categorie_document(nom)
            ajoutes += _copier_sans_remplacer(fichier, categorie, nom)[1]
    return ajoutes


def creer_bibliotheque_administrative(base: Path, source: Path | None = None,
                                     alimenter_depuis_source: bool = False) -> Path:
    """Crée la bibliothèque ; l'alimentation depuis le modèle est explicite."""
    base = Path(base)
    bibliotheque = base / DOSSIER_BIBLIOTHEQUE
    if bibliotheque.exists():
        if not bibliotheque.is_dir():
            raise ValueError(f"La bibliothèque n'est pas un dossier : {bibliotheque}")
        return bibliotheque
    if alimenter_depuis_source and (source is None or not Path(source).is_dir()):
        raise FileNotFoundError(source)
    with tempfile.TemporaryDirectory(prefix=".admin_", dir=base) as temporaire:
        preparation = Path(temporaire) / DOSSIER_BIBLIOTHEQUE
        preparation.mkdir()
        for categorie in DOSSIERS_CATEGORIES:
            (preparation / categorie).mkdir()
        if alimenter_depuis_source:
            ajouter_pieces_communes(preparation, source, DOSSIERS_COMMUNS)
            for fichier in _fichiers_valides(Path(source)):
                if fichier.suffix.casefold() in MODIFIABLES:
                    _copier_sans_remplacer(fichier, preparation / categorie_document(fichier), fichier.name)
        preparation.rename(bibliotheque)
    return bibliotheque


def importer_dans_bibliotheque(sources: list[Path], bibliotheque: Path) -> tuple[int, int]:
    """Importe des fichiers ou tout le contenu d'un dossier dans la bibliothèque unique."""
    bibliotheque = Path(bibliotheque).resolve()
    if not bibliotheque.is_dir():
        raise FileNotFoundError(bibliotheque)
    ajoutes = ignores = 0
    for source in map(Path, sources):
        source_resolue = source.resolve()
        if source_resolue == bibliotheque or bibliotheque in source_resolue.parents:
            raise ValueError("Ce document est déjà dans la bibliothèque.")
        if source.is_symlink() or (not source.is_file() and not source.is_dir()):
            raise FileNotFoundError(source)
        fichiers = [source] if source.is_file() else list(_fichiers_valides(source))
        for fichier in fichiers:
            if fichier.name.startswith((".", "~$")) or fichier.name.casefold() in EXCLUS:
                continue
            nom = (source.name + "__" + fichier.name) if source.is_dir() else fichier.name
            if _copier_sans_remplacer(fichier, bibliotheque / categorie_document(nom), nom)[1]:
                ajoutes += 1
            else:
                ignores += 1
    return ajoutes, ignores


def copier_documents_administratifs(documents: list[Path], bibliotheque: Path, chantier: Path,
                                     nom_copie: str | None = None) -> tuple[int, int]:
    """Copie les documents choisis à plat dans le chantier, avec nom adapté si demandé."""
    if nom_copie is not None and len(documents) != 1:
        raise ValueError("Le renommage concerne un seul document à la fois.")
    ajoutes = ignores = 0
    for document in documents:
        _, ajoute = copier_document_administratif(document, bibliotheque, chantier, nom_copie)
        if ajoute:
            ajoutes += 1
        else:
            ignores += 1
    return ajoutes, ignores


def copier_document_administratif(document: Path, bibliotheque: Path, chantier: Path,
                                  nom_copie: str | None = None) -> tuple[Path, bool]:
    """Retourne le chemin exact de la copie à ouvrir et indique si elle est nouvelle."""
    document = Path(document)
    document.resolve().relative_to(Path(bibliotheque).resolve())
    if not document.is_file() or document.name.casefold() in EXCLUS:
        raise ValueError(f"Document administratif invalide : {document}")
    if nom_copie is not None:
        nom = nom_copie.strip()
        if not nom or nom in {".", ".."} or "/" in nom or "\\" in nom:
            raise ValueError("Nom de copie invalide.")
        if nom.casefold().endswith(document.suffix.casefold()):
            nom = nom[:-len(document.suffix)]
        nom = _nom_portable(nom) + document.suffix
    else:
        nom = document.name.removeprefix("COMMUN__")
    destination = Path(chantier) / "Administratif"
    destination.mkdir(parents=True, exist_ok=True)
    return _copier_sans_remplacer(document, destination, nom)


def aplatir_bibliotheque(bibliotheque: Path, supplements: list[Path]) -> tuple[int, Path]:
    """Remplace l'ancien classement par un dossier plat après archive ZIP."""
    bibliotheque = Path(bibliotheque)
    base = bibliotheque.parent
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive = base / f"Archive_bibliotheque_administrative_{horodatage}.zip"
    with ZipFile(archive, "x") as zip_archive:
        for fichier in _fichiers_valides(bibliotheque):
            zip_archive.write(fichier, fichier.relative_to(bibliotheque).as_posix())
    with tempfile.TemporaryDirectory(prefix=".admin_plat_", dir=base) as temporaire:
        preparation = Path(temporaire) / DOSSIER_BIBLIOTHEQUE
        preparation.mkdir()
        fichiers = list(_fichiers_valides(bibliotheque))
        fichiers.sort(key=lambda p: (0 if "Dossier à envoyer/JT BATI" in p.relative_to(bibliotheque).as_posix() else 1,
                                     p.relative_to(bibliotheque).as_posix().casefold()))
        for fichier in fichiers:
            relatif = fichier.relative_to(bibliotheque)
            commun = len(relatif.parts) >= 4 and all(
                _nom_normalise(a) == _nom_normalise(b)
                for a, b in zip(relatif.parts[:3], Path(DOSSIER_COMMUN).parts)
            )
            prefixe = "COMMUN__" if commun else ""
            _copier_sans_remplacer(fichier, preparation, prefixe + fichier.name)
        for fichier in supplements:
            _copier_sans_remplacer(Path(fichier), preparation, Path(fichier).name)
        ancien = Path(temporaire) / "ancien"
        bibliotheque.rename(ancien)
        try:
            preparation.rename(bibliotheque)
        except Exception:
            ancien.rename(bibliotheque)
            raise
        compte = sum(p.is_file() for p in bibliotheque.iterdir())
    return compte, archive


def classer_bibliotheque_cinq_dossiers(bibliotheque: Path) -> tuple[dict[str, int], Path]:
    """Classe les fichiers du dossier unique en cinq dossiers après archivage."""
    bibliotheque = Path(bibliotheque)
    base = bibliotheque.parent
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive = base / f"Archive_bibliotheque_administrative_avant_classement_{horodatage}.zip"
    fichiers = list(_fichiers_valides(bibliotheque))
    with ZipFile(archive, "x") as zip_archive:
        for fichier in fichiers:
            zip_archive.write(fichier, fichier.relative_to(bibliotheque).as_posix())
    with tempfile.TemporaryDirectory(prefix=".admin_5_dossiers_", dir=base) as temporaire:
        preparation = Path(temporaire) / DOSSIER_BIBLIOTHEQUE
        preparation.mkdir()
        comptes = {nom: 0 for nom in DOSSIERS_CATEGORIES}
        empreintes = {nom: set() for nom in DOSSIERS_CATEGORIES}
        for nom in DOSSIERS_CATEGORIES:
            (preparation / nom).mkdir()
        fichiers.sort(key=lambda p: (p.name.startswith("COMMUN__"), " copie" in _sans_accents(p.stem),
                                     p.name.casefold()))
        for fichier in fichiers:
            categorie = categorie_document(fichier)
            empreinte = hashlib.sha256(fichier.read_bytes()).digest()
            if empreinte in empreintes[categorie]:
                continue
            empreintes[categorie].add(empreinte)
            nom = fichier.name.removeprefix("COMMUN__")
            if _copier_sans_remplacer(fichier, preparation / categorie, nom)[1]:
                comptes[categorie] += 1
        ancien = Path(temporaire) / "ancien"
        bibliotheque.rename(ancien)
        try:
            preparation.rename(bibliotheque)
        except Exception:
            ancien.rename(bibliotheque)
            raise
    return comptes, archive
