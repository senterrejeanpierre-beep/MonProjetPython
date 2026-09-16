#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
import subprocess
import shutil
import re
import time
import tempfile
import unicodedata
import math
import hashlib
from copy import copy
from datetime import datetime
from pathlib import Path
import tkinter as tk
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from tkinter import ttk, messagebox, filedialog

from openpyxl import load_workbook
from listes_pr import reparer_listes_pr

from bordereau_public import (feuille as _feuille_etat, est_metre as _est_metre_public,
                               lignes_postes as _postes_publics, lire_config as _config_public,
                               exporter_prix as _exporter_prix_public)

from environnement_partage import trouver_base, memoriser_base

APP_NAME = "Horizon Chantier"
_DOSSIER_BASE_MEMOIRE: Path | None = None

# =========================
# FICHIERS (STRICT / SANS REFONTE)
# =========================
PV_SOURCE = "soumission_chapelette__MAJ_PV_2026-01-29_13-52-57.xlsx"
PV_BACKUP = "soumission_chapelette__backup_avant_injectionPV_2026-01-29_13-52-57.xlsx"
PV_CORRIGE = "soumission_chapelette__MAJ_injection_CORRIGE_PV_2026-01-29_13-52-57.xlsx"
FICHIER_PV = PV_SOURCE


# =========================
# AppleScript helpers (Excel ouvert)
# =========================
def _osascript(script: str) -> str:
    if sys.platform != "darwin":
        raise RuntimeError("Cette commande d'automatisation Excel est disponible uniquement sur macOS.")
    p = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(p.stderr.strip() or "Erreur AppleScript")
    return (p.stdout or "").strip()


def _a1_to_col_row(a1: str) -> tuple[str, int]:
    """
    "C12" -> ("C", 12)
    """
    m = re.fullmatch(r"([A-Z]+)(\d+)", a1.strip().upper().replace("$", ""))
    if not m:
        raise ValueError(f"Adresse cellule invalide : {a1}")
    return m.group(1), int(m.group(2))


def _is_allowed_target(cell_a1: str) -> bool:
    """
    Sécurité absolue : n'autorise que
    - P3:P12  (matière)
    - P17:P26 (main-d'œuvre)
    """
    col, row = _a1_to_col_row(cell_a1)
    if col == "P" and 3 <= row <= 12:
        return True
    if col == "P" and 17 <= row <= 26:
        return True
    return False


def _normalize_excel_addr(addr: str) -> str:
    """
    Normalise:
    - "$P$3" -> "P3"
    - "P3"   -> "P3"
    """
    if not addr:
        return ""
    a = addr.strip().replace("$", "").upper()
    if re.fullmatch(r"[A-Z]+[0-9]+", a):
        return a
    return ""


def excel_find_first_empty_cell(workbook_hint: str, sheet_name: str, cell_range: str) -> str:
    """
    1ère cellule vide dans une plage (colonne unique) ex: P3:P12
    Retour "P3" ou "" si aucune.
    Robuste: retrouve le classeur par "contient" (Excel Mac change parfois le nom).
    """
    hint = (workbook_hint or "").strip()
    if not hint:
        return ""

    if sys.platform != "darwin":
        try:
            import xlwings as xw
            for app in xw.apps:
                for livre in app.books:
                    if hint.casefold() not in livre.name.casefold():
                        continue
                    feuille = livre.sheets[sheet_name]
                    for cellule in feuille.range(cell_range):
                        formule = cellule.formula
                        valeur = cellule.value
                        if not (isinstance(formule, str) and formule.startswith("=")) and valeur in (None, ""):
                            return cellule.address.replace("$", "")
        except Exception:
            return ""
        return ""

    for _ in range(8):
        script = f'''
        tell application "Microsoft Excel"
            set wbRef to missing value
            repeat with w in workbooks
                try
                    set wn to (name of w as string)
                    if wn contains "{hint}" then
                        set wbRef to w
                        exit repeat
                    end if
                end try
            end repeat

            if wbRef is missing value then return ""

            tell wbRef
                activate
                try
                    tell worksheet "{sheet_name}"
                        activate
                        set rng to range "{cell_range}"
                        repeat with c in cells of rng
                            try
                                set v to value of c
                                set f to formula of c

                                try
                                    if (f as string) is not "" then
                                        -- ne jamais proposer une cellule contenant une formule
                                    else
                                        if v is missing value then return (address of c) as string
                                        try
                                            if (v as string) is "" then return (address of c) as string
                                        end try
                                    end if
                                end try

                            on error
                                return (address of c) as string
                            end try
                        end repeat
                    end tell
                end try
            end tell
        end tell
        return ""
        '''
        try:
            addr_raw = _osascript(script)
            addr = _normalize_excel_addr(addr_raw)
            if addr:
                return addr
        except Exception:
            pass

        time.sleep(0.2)

    return ""


def excel_set_cell_value(workbook_hint: str, sheet_name: str, cell_a1: str, value: str) -> None:
    """
    Écrit une valeur dans Excel (classeur déjà ouvert) sans fermer.
    Sécurité : n'écrit QUE dans P3:P12 ou P17:P26.
    Robuste: retrouve le classeur par "contient".
    """
    cell_a1 = (cell_a1 or "").strip().replace("$", "").upper()
    if not _is_allowed_target(cell_a1):
        raise ValueError(f"Sécurité: écriture interdite hors zones autorisées : {cell_a1}")

    hint = (workbook_hint or "").strip()
    if not hint:
        raise ValueError("Nom de classeur vide")

    if sys.platform != "darwin":
        try:
            import xlwings as xw
            for app in xw.apps:
                for livre in app.books:
                    if hint.casefold() not in livre.name.casefold():
                        continue
                    cellule = livre.sheets[sheet_name].range(cell_a1)
                    if isinstance(cellule.formula, str) and cellule.formula.startswith("="):
                        raise ValueError("Cellule formule protégée")
                    cellule.value = value
                    livre.activate()
                    return
        except ValueError:
            raise
        except Exception as e:
            raise RuntimeError(f"Impossible de communiquer avec Microsoft Excel : {e}") from e
        raise RuntimeError("Le classeur Excel ouvert est introuvable")

    safe_val = (value or "").replace('"', '\\"')

    script = f'''
    tell application "Microsoft Excel"
        set wbRef to missing value
        repeat with w in workbooks
            try
                set wn to (name of w as string)
                if wn contains "{hint}" then
                    set wbRef to w
                    exit repeat
                end if
            end try
        end repeat

        if wbRef is missing value then return

        tell wbRef
            if not (exists worksheet "{sheet_name}") then return
            tell worksheet "{sheet_name}"
                if ((formula of range "{cell_a1}") as string) is not "" then error "Cellule formule protégée"
                set value of range "{cell_a1}" to "{safe_val}"
            end tell
        end tell
        activate
    end tell
    '''
    _osascript(script)


def _is_archive_or_history_path(path: Path) -> bool:
    ignored = {"archives", "archive", "harchives", "harchive", "historique_sauvegardes"}
    return any(part.lower() in ignored for part in path.parts)


def _chantier_root_for_path(path: str | Path) -> Path:
    src = Path(path).resolve()
    try:
        rel = src.relative_to(dossier_chantiers().resolve())
        if rel.parts:
            return dossier_chantiers() / rel.parts[0]
    except Exception:
        pass
    return src.parent.parent if src.parent.name == "data" else src.parent


def _historique_key(chantier_dir: Path, path: str | Path) -> str:
    src = Path(path)
    try:
        rel = src.relative_to(chantier_dir)
    except ValueError:
        rel = Path(src.name)
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", "__".join(rel.parts))


def _backup_excel_before_write(path: str | Path) -> Path:
    src = Path(path)
    chantier_dir = _chantier_root_for_path(src)
    backup_dir = chantier_dir / "Historique_Sauvegardes"
    backup_dir.mkdir(exist_ok=True)

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    key = _historique_key(chantier_dir, src)
    dst = backup_dir / f"{stamp}_{key}"
    if dst.exists():
        i = 2
        while True:
            candidate = backup_dir / f"{stamp}_{i}_{key}"
            if not candidate.exists():
                dst = candidate
                break
            i += 1

    shutil.copy2(src, dst)
    return dst


def _normaliser_libelle_excel(valeur):
    txt = str(valeur or "").strip().lower()
    txt = txt.replace("é", "e").replace("è", "e").replace("ê", "e").replace("ë", "e")
    txt = txt.replace("à", "a").replace("â", "a").replace("ä", "a")
    txt = txt.replace("ù", "u").replace("û", "u").replace("ü", "u")
    txt = txt.replace("î", "i").replace("ï", "i")
    txt = txt.replace("ô", "o").replace("ö", "o")
    txt = txt.replace("ç", "c")
    txt = re.sub(r"\s+", " ", txt)
    return txt


def _premiere_ligne_synthese_etat(ws) -> int:
    libelles_synthese = (
        "total soumission hors tva",
        "total etat cumule hors tva",
        "total du mois hors tva",
        "total des avenants cumule",
        "montant global a facturer",
        "total execute",
    )

    for ligne in range(2 if _est_metre_public(ws) else 24, ws.max_row + 1):
        contenus = [_normaliser_libelle_excel(ws.cell(ligne, col).value) for col in range(16, 24)]
        if any(contenu.startswith(libelle) for contenu in contenus for libelle in libelles_synthese):
            return ligne

    return ws.max_row + 1


def _montant_diminution_avenant(montant: float) -> float:
    return abs(montant)


def _soumission_avant_cloture(chemin, convertir_nombre):
    """Retrouve le total sauvegardé avant l'effacement du libellé en Q."""
    chemin = Path(chemin)
    sauvegardes = sorted(
        (chemin.parent / "Sauvegarde").glob(f"{chemin.stem}_*{chemin.suffix}"),
        key=lambda fichier: fichier.stat().st_mtime,
        reverse=True,
    )
    for sauvegarde in sauvegardes:
        wb = load_workbook(sauvegarde, data_only=True)
        try:
            if "Bordereau" not in wb.sheetnames:
                continue
            ws = _feuille_etat(wb)
            for row in reversed(list(ws.iter_rows())):
                for cellule in row:
                    if _normaliser_libelle_excel(cellule.value) != "total soumission hors tva":
                        continue
                    valeur = ws.cell(cellule.row, cellule.column + 1).value
                    if valeur is not None:
                        return convertir_nombre(valeur)
        finally:
            wb.close()
    return None


def _reporter_quantites_cloture(ws):
    # Une synthèse intermédiaire ne doit pas masquer les postes suivants.
    reports = []
    for ligne in (_postes_publics(ws) if _est_metre_public(ws) else range(24, ws.max_row + 1)):
        if _ligne_synthese_etat(ws, ligne):
            continue
        valeurs = [ws[f"{col}{ligne}"].value for col in ("P", "Q", "R")]
        if all(v in (None, "") for v in valeurs):
            continue
        # Les en-têtes textuels ne sont pas des quantités.
        if not _article_pr(ws[f"B{ligne}"].value) and any(
            isinstance(v, str) and v.strip() and not re.fullmatch(r"[-+\d.,\s]+", v)
            for v in valeurs
        ):
            continue
        cumul = _nombre_metier(valeurs[2], f"R{ligne}")
        reports.append((ligne, cumul))
    for ligne, cumul in reports:
        ws[f"P{ligne}"] = cumul
        ws[f"Q{ligne}"] = 0


def _nombre_metier(valeur, cellule=""):
    texte = re.sub(r"\s+", "", str(valeur if valeur is not None else ""))
    texte = texte.replace("€", "").replace("−", "-").replace(",", ".")
    if texte in ("", "-", "—"):
        return 0.0
    try:
        resultat = float(texte)
        if not math.isfinite(resultat):
            raise ValueError
        return resultat
    except (TypeError, ValueError):
        raise ValueError(f"Valeur numérique indisponible ou invalide en {cellule} : {valeur!r}. "
                         "Vérifiez et enregistrez le classeur dans Excel.") from None


def _presentation_ecarts_pilotage(soumission, marche_total, production_cumulee):
    reste = marche_total - production_cumulee
    ecart_marche = -reste
    if reste < 0:
        libelle_marche = "Production au-delà du marché total"
        valeur_marche = -reste
    else:
        libelle_marche = "Reste à facturer sur le marché total"
        valeur_marche = reste
    ecart_initial = production_cumulee - soumission
    if ecart_initial > 0:
        texte_initial = f"{ecart_initial:+,.2f} € au-dessus"
    elif ecart_initial < 0:
        texte_initial = f"{ecart_initial:+,.2f} € en dessous"
    else:
        texte_initial = "0.00 € : égal à la soumission"
    texte_marche = f"{ecart_marche:+,.2f} €" if ecart_marche else "0.00 €"
    return libelle_marche, valeur_marche, texte_marche, texte_initial


def _article_pr(val):
    if val is None:
        return ""
    txt = re.sub(r"\s+", "", str(val).strip().replace("\xa0", " ").replace(",", ".").lower())
    parts = []
    for part in txt.split("."):
        if not part:
            continue
        match = re.fullmatch(r"(\d+)([a-z]+)?", part)
        if match:
            parts.append(str(int(match.group(1))) + (match.group(2) or ""))
        else:
            cleaned = re.sub(r"[^0-9a-z]+", "", part)
            if cleaned:
                parts.append(cleaned)
    article = ".".join(parts)
    return article if any(c.isdigit() for c in article) else ""


def _ligne_synthese_etat(ws, ligne):
    # Ne pas examiner les anciens états placés à droite du bordereau courant.
    libelles = [ws[f"C{ligne}"].value] + [ws.cell(ligne, col).value for col in range(16, 24)]
    return any(
        _normaliser_libelle_excel(v).startswith(("total", "montant", "synthese", "sous-total"))
        for v in libelles if isinstance(v, str) and not v.startswith("=")
    )


def _corriger_moins_etat_public(ws):
    """Retire les montants contractuels d'avenants des totaux exécutés publics connus."""
    ancienne_formule_cumulee = "=SUM('[1]Feuil1'!$D$3:$D$60)-SUM('[1]Feuil1'!$E$3:$E$60)"
    formule_plus_cumulee = "=SUM('[1]Feuil1'!$D$3:$D$60)"
    ancienne_formule_mois = ("=N(INDEX('[1]Feuil1'!$D$3:$D$60,1+3*($Q$9-1)))"
                             "-N(INDEX('[1]Feuil1'!$E$3:$E$60,1+3*($Q$9-1)))")
    ancienne_formule_plus_mois = "=N(INDEX('[1]Feuil1'!$D$3:$D$60,1+3*($Q$9-1)))"
    ancienne_etiquette_mois = '=\"Montant de l’avenant \"&Q9'
    for ligne in range(1, ws.max_row + 1):
        libelle = _normaliser_libelle_excel(ws[f"P{ligne}"].value)
        if libelle == "total des avenants cumule" and ws[f"T{ligne}"].value == ancienne_formule_cumulee:
            ws[f"T{ligne}"] = formule_plus_cumulee
        elif libelle == "avenants cumules en moins":
            valeur = ws[f"T{ligne}"].value
            if valeur not in (None, 0, "=SUM('[1]Feuil1'!$E$3:$E$60)"):
                raise ValueError(f"Avenants en moins saisis dans l'état en T{ligne}. "
                                 "Vérifiez cette ancienne ligne avant de recalculer.")
            ws[f"P{ligne}"] = None
            ws[f"T{ligne}"] = None
        elif libelle == "total etat cumule hors tva" and ligne + 2 <= ws.max_row:
            etiquette = ws[f"U{ligne + 1}"].value
            if etiquette == ancienne_etiquette_mois:
                valeur = ws[f"W{ligne + 1}"].value
                if valeur not in (None, 0, ancienne_formule_mois, ancienne_formule_plus_mois):
                    raise ValueError(f"Montant d'avenant saisi en W{ligne + 1}. "
                                     "Vérifiez cette ancienne ligne avant de recalculer.")
                ws[f"U{ligne + 1}"] = None
                ws[f"W{ligne + 1}"] = None
            ancien_total = f"=SUM(W{ligne}:W{ligne + 1})"
            if ws[f"W{ligne + 2}"].value == ancien_total:
                ws[f"W{ligne + 2}"] = f"=W{ligne}"


def _corriger_avenants_externes_etat_prive(ws):
    """Écarte les liens contractuels connus sans toucher aux postes d'avenant exécutés."""
    for ligne in range(24, ws.max_row):
        libelle = _normaliser_libelle_excel(ws[f"V{ligne}"].value)
        if not libelle.startswith("montant de l'avenant"):
            continue
        valeur = ws[f"W{ligne}"].value
        if not isinstance(valeur, str) or not valeur.startswith("="):
            continue
        formule = valeur.casefold()
        lien_externe = "[1]feuil1!" in formule or "avenants!" in formule
        if not lien_externe:
            cellule_source = re.fullmatch(r"=sum\((m\d+)\)", formule)
            if cellule_source:
                source = ws[cellule_source.group(1).upper()].value
                lien_externe = isinstance(source, str) and "avenants!" in source.casefold()
        if not lien_externe:
            continue
        ws[f"V{ligne}"] = None
        ws[f"W{ligne}"] = None
        if ws[f"W{ligne + 1}"].value == f"=SUM(W{ligne - 1}:W{ligne})":
            ws[f"W{ligne + 1}"] = f"=W{ligne - 1}"


def _selectionner_etat(dossier):
    dossier = Path(dossier)
    configuration = _config_public(dossier)
    if configuration:
        copie = dossier / configuration["copie"]
        if not copie.is_file():
            raise FileNotFoundError(f"Copie du bordereau public introuvable : {copie}. Aucun ancien modèle utilisé.")
        return copie
    fichiers = sorted(p for p in dossier.glob("Etat_avancement*.xlsm") if not p.name.startswith("~$"))
    if not fichiers:
        raise FileNotFoundError(f"Aucun état d'avancement dans : {dossier}")
    fiche = dossier.parent / f"{dossier.name}.json"
    type_etat = lire_json(fiche).get("type_etat") if fiche.exists() else None
    attendu = {"Public": "Etat_avancement_Public.xlsm", "Privé": "Etat_avancement_Privé.xlsm",
               "Modèle actuel": "Etat_avancement_02.xlsm"}.get(type_etat)
    if len(fichiers) == 1:
        nom = unicodedata.normalize("NFC", fichiers[0].name).casefold()
        if attendu and nom in {"etat_avancement_public.xlsm", "etat_avancement_privé.xlsm"} and nom != attendu.casefold():
            raise ValueError(f"L'état présent ne correspond pas au type {type_etat} de la fiche chantier.")
        return fichiers[0]
    if attendu:
        for p in fichiers:
            if unicodedata.normalize("NFC", p.name).casefold() == attendu.casefold():
                return p
    raise ValueError("Plusieurs états d'avancement sont présents sans choix identifiable Public/Privé. "
                     "Conservez l'état à utiliser ou renseignez le type dans la fiche chantier.")


def _signature_excel(chemin, verifier_verrou=True):
    chemin = Path(chemin)
    if verifier_verrou and chemin.with_name("~$" + chemin.name).exists():
        raise ValueError(f"Fermez {chemin.name} dans Excel après l'avoir enregistré, puis réessayez.")
    stat = chemin.stat()
    empreinte = hashlib.sha256()
    with chemin.open("rb") as fichier:
        for bloc in iter(lambda: fichier.read(1024 * 1024), b""):
            empreinte.update(bloc)
    stat_fin = chemin.stat()
    if (stat.st_mtime_ns, stat.st_size, stat.st_ino) != (stat_fin.st_mtime_ns, stat_fin.st_size, stat_fin.st_ino):
        raise ValueError("Le classeur a changé pendant la lecture. Relancez l'opération.")
    return stat.st_mtime_ns, stat.st_size, stat.st_ino, empreinte.digest()


def _fermer_classeur(wb):
    wb.close()
    archive_vba = getattr(wb, "vba_archive", None)
    if archive_vba is not None:
        archive_vba.close()


def _sauver_excel_atomique(wb, chemin, signature):
    chemin = Path(chemin)
    if _signature_excel(chemin) != signature:
        raise ValueError("Le classeur a été modifié pendant l'opération. Aucun remplacement effectué.")
    fd, nom = tempfile.mkstemp(prefix=".horizon_", suffix=chemin.suffix, dir=chemin.parent)
    os.close(fd)
    temporaire = Path(nom)
    try:
        wb.save(temporaire)
        shutil.copymode(chemin, temporaire)
        if _signature_excel(chemin) != signature:
            raise ValueError("Le classeur a été modifié pendant l'opération. Aucun remplacement effectué.")
        _backup_excel_before_write(chemin)
        if _signature_excel(chemin) != signature:
            raise ValueError("Le classeur a été modifié pendant la sauvegarde. Aucun remplacement effectué.")
        os.replace(temporaire, chemin)
    finally:
        temporaire.unlink(missing_ok=True)


def _chemin_pdf_temporaire(destination):
    destination = Path(destination)
    fd, nom = tempfile.mkstemp(prefix=".horizon_pdf_", suffix=".pdf", dir=destination.parent)
    os.close(fd)
    return Path(nom)


def _publier_pdf_atomique(doc, elements, destination):
    temporaire = _chemin_pdf_temporaire(destination)
    doc.filename = str(temporaire)
    try:
        doc.build(elements)
        os.replace(temporaire, destination)
    finally:
        temporaire.unlink(missing_ok=True)


def _charger_valeurs_excel(chemin, feuille, libelle, max_col=None, lecture_seule=True,
                           colonnes_formules_facultatives=()):
    """Contrôle les résultats enregistrés, sans recalculer ni écrire le classeur."""
    signature = _signature_excel(chemin, verifier_verrou=False)
    formules = load_workbook(chemin, read_only=True, data_only=False)
    valeurs = None
    try:
        valeurs = load_workbook(chemin, read_only=lecture_seule, data_only=True)
        ws_f = _feuille_etat(formules) if feuille == "Bordereau" else (formules[feuille] if feuille else formules.active)
        ws_v = valeurs[ws_f.title]
        manquantes = []
        for row_f, row_v in zip(ws_f.iter_rows(max_col=max_col), ws_v.iter_rows(max_col=max_col)):
            for f, v in zip(row_f, row_v):
                if v.data_type == "e" or (f.data_type == "f" and v.value is None
                                          and f.column_letter not in colonnes_formules_facultatives):
                    manquantes.append(f.coordinate)
        if manquantes:
            if libelle == "État d'avancement":
                # Montrer d'abord les montants T/W requis par le pilotage.
                manquantes.sort(key=lambda adresse: (adresse[0] not in "TW", int(re.search(r"\d+", adresse).group())))
            raise ValueError(f"{libelle} : {len(manquantes)} résultat(s) Excel indisponible(s) ou en erreur "
                             f"({', '.join(manquantes[:8])}). Recalculez et enregistrez le classeur dans Excel "
                             "avant de recommencer. Aucun résultat ne peut être validé.")
        if _signature_excel(chemin, verifier_verrou=False) != signature:
            raise ValueError(f"{libelle} : le classeur a changé pendant la lecture. Relancez l'opération.")
        return valeurs
    except Exception:
        if valeurs is not None:
            valeurs.close()
        raise
    finally:
        formules.close()


def _charger_pr_controle(chemin):
    # Lecture seulement : aucune formule, mise en page ou valeur du PR n'est écrite.
    return _charger_valeurs_excel(chemin, "Chiffrage", "PR")


def _verifier_blocs_pr(ws, ws_formules=None):
    lignes = list(ws.iter_rows(max_col=16, values_only=True))
    formules_pr = None
    if ws_formules is not None:
        formules_pr = {ligne: tuple(cellule.value for cellule in cellules)
                       for ligne, cellules in enumerate(ws_formules.iter_rows(min_col=9, max_col=12), 1)}
    debuts = [i for i, row in enumerate(lignes, 1) if i >= 3 and _article_pr(row[1])]
    if not debuts:
        raise ValueError("Aucun bloc article détecté dans le PR.")
    erreurs = []
    def nombre(ligne, colonne):
        return _nombre_metier(lignes[ligne - 1][colonne - 1], f"ligne {ligne}, colonne {colonne}")
    def calcul_simple(ligne, colonne):
        if formules_pr is None:
            return None
        formule = formules_pr[ligne][colonne - 9]
        if not isinstance(formule, str) or not formule.startswith("="):
            return ""
        expression = formule[1:].upper().replace("$", "").replace(" ", "")
        for enveloppe in ("_XLFN.SINGLE", "SUM"):
            prefixe = enveloppe + "("
            while expression.startswith(prefixe) and expression.endswith(")"):
                expression = expression[len(prefixe):-1]
        return expression
    def produit(expression, ligne, gauche, droite):
        return expression in (f"{gauche}{ligne}*{droite}{ligne}",
                              f"{droite}{ligne}*{gauche}{ligne}")
    for debut in debuts:
        if debut + 27 > len(lignes):
            raise ValueError(f"Bloc PR incomplet à la ligne {debut}.")
        for ligne in range(debut, debut + 10):
            attendu_k = nombre(ligne, 9) * (1 + nombre(ligne, 10))
            pu_corrige = nombre(ligne, 11)
            expression_k = calcul_simple(ligne, 11)
            pu_a_calculer = formules_pr is None or expression_k in (
                f"I{ligne}*(1+J{ligne})", f"(1+J{ligne})*I{ligne}")
            if pu_a_calculer and abs(pu_corrige - attendu_k) > 0.01:
                erreurs.append(f"Ligne {ligne} : PU matière à vérifier")
            expression_l = calcul_simple(ligne, 12)
            if formules_pr is None:
                attendu_l = nombre(ligne, 8) * pu_corrige
            elif produit(expression_l, ligne, "H", "K"):
                attendu_l = nombre(ligne, 8) * pu_corrige
            elif produit(expression_l, ligne, "G", "K"):
                attendu_l = nombre(ligne, 7) * pu_corrige
            else:
                attendu_l = None
            if attendu_l is not None and abs(nombre(ligne, 12) - attendu_l) > 0.01:
                erreurs.append(f"Ligne {ligne} : total matière à vérifier")
        for ligne in range(debut + 13, debut + 23):
            if all(lignes[ligne - 1][col - 1] in (None, "") for col in (8, 9, 10, 11)):
                continue  # Commentaire sans heures ni coût.
            attendu_i = nombre(ligne, 3) * nombre(ligne, 8)
            attendu_k = nombre(ligne, 9) * nombre(ligne, 10)
            expression_i = calcul_simple(ligne, 9)
            if (formules_pr is None or produit(expression_i, ligne, "C", "H")) and abs(nombre(ligne, 9) - attendu_i) > 0.01:
                erreurs.append(f"Ligne {ligne} : heures MO à vérifier")
            expression_k = calcul_simple(ligne, 11)
            if produit(expression_k, ligne, "H", "J"):
                attendu_k = nombre(ligne, 8) * nombre(ligne, 10)
            elif formules_pr is not None and not produit(expression_k, ligne, "I", "J"):
                continue
            if abs(nombre(ligne, 11) - attendu_k) > 0.01:
                erreurs.append(f"Ligne {ligne} : coût MO à vérifier")
    return len(debuts), erreurs


def _recalculer_etat(ws, valeurs=None):
    def normaliser_libelle(valeur):
        texte = str(valeur or "").strip().lower()
        texte = texte.replace("é", "e").replace("è", "e").replace("ê", "e").replace("ë", "e")
        texte = texte.replace("à", "a").replace("â", "a").replace("ä", "a")
        texte = texte.replace("ù", "u").replace("û", "u").replace("ü", "u")
        texte = texte.replace("î", "i").replace("ï", "i")
        texte = texte.replace("ô", "o").replace("ö", "o")
        texte = texte.replace("ç", "c")
        texte = texte.replace("\n", " ")
        texte = re.sub(r"\s+", " ", texte)
        return texte

    def trouver_colonne_bordereau(alias):
        alias_normalises = {normaliser_libelle(a) for a in alias}
        max_ligne_entete = min(15, ws.max_row)
        for ligne in range(1, max_ligne_entete + 1):
            for col in range(1, ws.max_column + 1):
                libelle = normaliser_libelle(ws.cell(ligne, col).value)
                if libelle in alias_normalises or any(a in libelle for a in alias_normalises):
                    return ws.cell(ligne, col).column_letter
        return None

    col_quantite = trouver_colonne_bordereau({"quantite", "qte", "qte.", "qte :", "qte/", "qté"})
    col_prix_unitaire = trouver_colonne_bordereau({"prix unitaire", "pu", "p.u.", "prix unit"})

    if not col_quantite or not col_prix_unitaire:
        erreurs = []
        if not col_quantite:
            erreurs.append("colonne quantité introuvable (attendu : quantité, quantite, qté, qte)")
        if not col_prix_unitaire:
            erreurs.append("colonne prix unitaire introuvable (attendu : prix unitaire, pu, p.u., prix unit)")
        raise ValueError("Impossible de recalculer la feuille Bordereau.\n" + "\n".join(erreurs))

    def nombre_cellule(adresse):
        cellule = ws[adresse]
        valeur = cellule.value
        if cellule.data_type == "f":
            valeur = valeurs[adresse].value if valeurs is not None else None
            if valeur is None:
                raise ValueError(f"Résultat Excel absent en {adresse}. Recalculez et enregistrez l'état dans Excel.")
        return _nombre_metier(valeur, adresse)

    public = _est_metre_public(ws)
    if public:
        _corriger_moins_etat_public(ws)
    else:
        _corriger_avenants_externes_etat_prive(ws)
    for ligne in (_postes_publics(ws) if public else range(24, ws.max_row + 1)):
        if _ligne_synthese_etat(ws, ligne):
            continue
        # Un poste ajouté à la main peut ne pas avoir de code numérique en B.
        # Ses quantités saisies et son prix suffisent alors à l'identifier.
        if not _article_pr(ws[f"B{ligne}"].value):
            if public or not any(ws[f"{col}{ligne}"].value not in (None, "", "-", "—") for col in ("P", "Q")):
                continue
            if any(ws[f"{col}{ligne}"].value in (None, "", "-", "—") for col in (col_quantite, col_prix_unitaire)):
                continue
        if all(ws[f"{col}{ligne}"].value in (None, "", "-", "—") for col in (col_quantite, col_prix_unitaire)):
            continue
        j = nombre_cellule(f"{col_quantite}{ligne}")
        l = nombre_cellule(f"{col_prix_unitaire}{ligne}")
        p = nombre_cellule(f"P{ligne}")
        q = nombre_cellule(f"Q{ligne}")

        r = p + q
        s = 0 if j == 0 else r / j
        t = r * l
        u = q
        v = l
        w = u * v

        if j == 0 and l == 0:
            continue

        ws[f"R{ligne}"] = r
        ws[f"S{ligne}"] = f'=IF(AND(OR(P{ligne}="",P{ligne}=0),OR(Q{ligne}="",Q{ligne}=0),OR(R{ligne}="",R{ligne}=0)),"",IFERROR(R{ligne}/{col_quantite}{ligne},0))'
        ws[f"S{ligne}"].number_format = "0.00%"
        ws[f"T{ligne}"] = t
        ws[f"U{ligne}"] = u
        ws[f"V{ligne}"] = v
        ws[f"W{ligne}"] = w




    # Correction limitée à la formule connue qui ajoutait le mois au cumul une seconde fois.
    if col_quantite == "E" and col_prix_unitaire == "H":
        cumul = mois = execute = None
        for ligne in (_postes_publics(ws) if _est_metre_public(ws) else range(24, ws.max_row + 1)):
            libelle = _normaliser_libelle_excel(ws[f"S{ligne}"].value)
            if libelle.startswith("total etat cumule hors tva"):
                cumul = f"T{ligne}"
            elif libelle.startswith("total du mois hors tva"):
                mois = f"T{ligne}"
            elif libelle.startswith("total execute"):
                execute = f"T{ligne}"
        if cumul and mois and execute and ws[execute].value == f"=SUM({cumul}:{mois})":
            ws[execute] = f"={cumul}"


def _normaliser_dossier_data(dossier_chantier: Path) -> Path:
    data_path = dossier_chantier / "data"
    if not dossier_chantier.exists() or data_path.exists():
        return data_path

    for candidat in dossier_chantier.iterdir():
        if candidat.is_dir() and candidat.name.strip() == "data":
            candidat.rename(data_path)
            return data_path

    return data_path


def _chemin_fichier_chantier(dossier_chantier: Path, nom_fichier: str) -> Path:
    if Path(nom_fichier).parts and Path(nom_fichier).parts[0].strip() == "data":
        _normaliser_dossier_data(dossier_chantier)

    chemin = dossier_chantier / nom_fichier
    if dossier_chantier.exists():
        for candidat in dossier_chantier.iterdir():
            if candidat.name == nom_fichier:
                return candidat

        attendu = unicodedata.normalize("NFC", nom_fichier)
        for candidat in dossier_chantier.iterdir():
            if unicodedata.normalize("NFC", candidat.name) == attendu:
                return candidat

    if chemin.exists():
        return chemin

    return chemin


def _montant_a_droite(ws, row, col_depart, nombre, max_ecart=3, requis=False):
    if not row or not col_depart:
        if requis:
            raise ValueError("Libellé de synthèse introuvable dans l'état d'avancement.")
        return 0

    for fusion in ws.merged_cells.ranges:
        if fusion.min_row <= row <= fusion.max_row and fusion.min_col <= col_depart <= fusion.max_col:
            col_depart = fusion.max_col
            break
    col_fin = min(col_depart + max_ecart, ws.max_column)

    for col in range(col_depart + 1, col_fin + 1):
        val = ws.cell(row=row, column=col).value

        # Ignore les cellules vides ou décoratives
        if val in (None, "", "-", "—"):
            continue

        # Si Excel renvoie déjà un nombre, on le prend tel quel
        if isinstance(val, (int, float)):
            return _nombre_metier(val, ws.cell(row=row, column=col).coordinate) if requis else float(val)

        # Sinon on passe par le parseur existant
        num = _nombre_metier(val, ws.cell(row=row, column=col).coordinate) if requis else nombre(val)
        txt = str(val).strip()

        # Conserve aussi un vrai zéro explicite
        if num != 0 or txt in {"0", "0,0", "0.0", "0,00", "0.00"}:
            return num

    if requis:
        raise ValueError(f"Montant de synthèse absent ou illisible à la ligne {row} de l'état d'avancement.")
    return 0


def _lire_synthese_avenants_pilotage(ws, convertir_nombre):
    avenants_plus = _lire_montant_synthese_colonne(ws, "E", convertir_nombre)
    avenants_moins = _montant_diminution_avenant(
        _lire_montant_synthese_colonne(ws, "G", convertir_nombre)
    )
    return avenants_plus, avenants_moins


def _lire_avenants_moins_publics(chemin_etat):
    """Lit les moins dans Avenants.xlsx du chantier, jamais dans l'état exécuté."""
    chemin_avenants = Path(chemin_etat).with_name("Avenants.xlsx")
    if not chemin_avenants.is_file():
        raise FileNotFoundError(f"Avenants.xlsx introuvable pour le pilotage : {chemin_avenants}")
    wb = _charger_valeurs_excel(chemin_avenants, None, "Avenants", max_col=5, lecture_seule=False)
    try:
        ws = wb.active
        for ligne in range(ws.max_row, 0, -1):
            if _normaliser_libelle_excel(ws[f"B{ligne}"].value).startswith("montant cumule a reporter"):
                cellule = ws[f"E{ligne}"]
                if cellule.value in (None, ""):
                    raise ValueError(f"Avenants en moins : résultat absent en {cellule.coordinate}.")
                return abs(_nombre_metier(cellule.value, cellule.coordinate))
        raise ValueError("Synthèse des avenants en moins introuvable dans Avenants.xlsx.")
    finally:
        wb.close()


def _lire_revision_globale_pilotage(ws, convertir_nombre):
    return _lire_montant_synthese_colonne(ws, "E", convertir_nombre)


def _lire_montant_synthese_colonne(ws, colonne: str, convertir_nombre) -> float:
    for ligne in range(ws.max_row, 0, -1):
        valeur = ws[f"{colonne}{ligne}"].value
        if valeur in (None, "", "-", "—"):
            continue
        montant = convertir_nombre(valeur)
        texte_valeur = str(valeur).strip()
        if montant != 0 or texte_valeur in {"0", "0,0", "0.0", "0,00", "0.00"}:
            return montant
    return 0


def _set_cell_protection(cell, locked: bool, hidden: bool = False) -> None:
    prot = copy(cell.protection)
    prot.locked = locked
    prot.hidden = hidden
    cell.protection = prot


def _protect_formula_cells(wb, source_path: str | Path | None = None) -> None:
    filename = Path(source_path).name.lower() if source_path else ""

    for ws in wb.worksheets:
        synthese_debut = None
        if ("etat_avancement" in filename and ws.title == "Bordereau") or _est_metre_public(ws):
            synthese_debut = _premiere_ligne_synthese_etat(ws)

        for row in ws.iter_rows():
            for cell in row:
                if cell.__class__.__name__ == "MergedCell":
                    continue

                in_synthese = (synthese_debut is not None and cell.row >= synthese_debut
                               and (not _est_metre_public(ws) or 16 <= cell.column <= 23))
                has_formula = cell.data_type == "f" or (
                    isinstance(cell.value, str) and cell.value.startswith("=")
                )

                if in_synthese:
                    _set_cell_protection(cell, locked=False, hidden=False)
                else:
                    _set_cell_protection(cell, locked=bool(has_formula), hidden=bool(has_formula))

        ws.protection.sheet = True
        ws.protection.set_password("1234")


def _prepare_excel_file_before_write(path: str | Path, keep_vba: bool = False) -> None:
    path = Path(path)
    if Path(path).name.lower() == "prix_de_revient.xlsx":
        reparer_listes_pr(path, _backup_excel_before_write)
    signature = _signature_excel(path)
    wb = load_workbook(path, keep_vba=keep_vba or path.suffix.lower() == ".xlsm")
    try:
        _protect_formula_cells(wb, path)
        _sauver_excel_atomique(wb, path, signature)
    finally:
        _fermer_classeur(wb)


def _export_excel_to_pdf(excel_path: str | Path, pdf_path: str | Path) -> None:
    import xlwings as xw

    excel_path = Path(excel_path).resolve()
    pdf_path = Path(pdf_path)
    book = None
    opened_here = False

    for app in xw.apps:
        for candidate in app.books:
            try:
                if Path(candidate.fullname).resolve() == excel_path:
                    book = candidate
                    break
            except Exception:
                continue
        if book is not None:
            break

    if book is None:
        book = xw.Book(str(excel_path))
        opened_here = True

    try:
        book.to_pdf(str(pdf_path))
    finally:
        if opened_here:
            book.close()


def _chemin_classeur_excel_actif() -> Path | None:
    import xlwings as xw

    apps = []
    try:
        apps.append(xw.apps.active)
    except Exception:
        pass

    try:
        for app in xw.apps:
            if app not in apps:
                apps.append(app)
    except Exception:
        pass

    for app in apps:
        try:
            book = app.books.active
            chemin = Path(book.fullname)
        except Exception:
            continue
        if chemin.suffix.lower() in {".xlsx", ".xlsm", ".xls"} and chemin.exists():
            return chemin

    return None


# =========================
# Lecture bibliothèques (colonne A) - lecture seule
# =========================
def read_biblio_colA_openpyxl(pr_path: Path, sheet_name: str) -> list[str]:
    wb = load_workbook(pr_path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Onglet introuvable : {sheet_name}")
        ws = wb[sheet_name]
        out: list[str] = []
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v is None:
                continue
            s = str(v).strip()
            if s:
                out.append(s)
        return out
    finally:
        wb.close()


# =========================
# Popup recherche (20 résultats)
# =========================
def popup_search_20(parent: tk.Tk, titre: str, items: list[str], on_pick) -> None:
    win = tk.Toplevel(parent)
    win.title(titre)
    win.geometry("780x520")
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text=titre, font=("Helvetica", 14, "bold")).pack(anchor="w")

    q = tk.StringVar()
    entry = ttk.Entry(frm, textvariable=q)
    entry.pack(fill="x", pady=(10, 8))
    entry.focus_set()

    lb = tk.Listbox(frm, height=18)
    lb.pack(fill="both", expand=True)

    ttk.Label(frm, text="Tape 2 lettres. Double-clic ou Entrée pour valider.", foreground="#555").pack(
        anchor="w", pady=(8, 0)
    )

    current: list[str] = []

    def refresh(*_):
        text = q.get().strip().lower()
        lb.delete(0, tk.END)

        if not text:
            current[:] = items[:20]
        else:
            filtered = [s for s in items if text in s.lower()]
            current[:] = filtered[:20]

        for s in current:
            lb.insert(tk.END, s)

        if current:
            lb.selection_set(0)

    def choose(_event=None):
        sel = lb.curselection()
        if not sel:
            return
        picked = current[sel[0]]
        try:
            on_pick(picked)
        finally:
            win.destroy()

    q.trace_add("write", refresh)
    lb.bind("<Double-Button-1>", choose)
    win.bind("<Return>", choose)

    refresh()


# ============================================================
# Montant en lettres FR (euros + centimes) - inchangé
# ============================================================
_UNITS = ["zéro", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf"]
_TEENS = ["dix", "onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf"]
_TENS = ["", "", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante", "quatre-vingt", "quatre-vingt"]


def _below_100(n: int) -> str:
    assert 0 <= n < 100
    if n < 10:
        return _UNITS[n]
    if 10 <= n < 20:
        return _TEENS[n - 10]

    ten = n // 10
    unit = n % 10

    if ten == 7:
        return "soixante" + ("-" + _below_100(10 + unit) if unit else "-dix")
    if ten == 9:
        return "quatre-vingt" + ("-" + _below_100(10 + unit) if unit else "-dix")

    base = _TENS[ten]

    if ten == 8 and unit == 0:
        return "quatre-vingts"
    if unit == 0:
        return base

    if unit == 1 and ten in (2, 3, 4, 5, 6):
        return f"{base} et un"

    return f"{base}-{_UNITS[unit]}"


def _below_1000(n: int) -> str:
    assert 0 <= n < 1000
    if n < 100:
        return _below_100(n)

    hundred = n // 100
    rest = n % 100

    if hundred == 1:
        head = "cent"
    else:
        head = f"{_UNITS[hundred]} cent"

    if rest == 0 and hundred > 1:
        return head + "s"
    if rest == 0:
        return head

    return head + " " + _below_100(rest)


def _int_to_words_fr(n: int) -> str:
    if n == 0:
        return "zéro"

    parts: list[str] = []

    billions = n // 1_000_000_000
    n %= 1_000_000_000
    millions = n // 1_000_000
    n %= 1_000_000
    thousands = n // 1000
    n %= 1000
    rest = n

    if billions:
        parts.append("un milliard" if billions == 1 else _below_1000(billions) + " milliards")

    if millions:
        parts.append("un million" if millions == 1 else _below_1000(millions) + " millions")

    if thousands:
        parts.append("mille" if thousands == 1 else _below_1000(thousands) + " mille")

    if rest:
        parts.append(_below_1000(rest))

    return " ".join(parts)


def montant_en_lettres_fr(valeur: float) -> str:
    v = round(float(valeur), 2)
    euros = int(v)
    centimes = int(round((v - euros) * 100))
    if centimes == 100:
        euros += 1
        centimes = 0

    euros_txt = _int_to_words_fr(euros)
    euro_label = "euro" if euros == 1 else "euros"

    if centimes == 0:
        return f"{euros_txt} {euro_label}"

    cent_txt = _int_to_words_fr(centimes)
    cent_label = "centime" if centimes == 1 else "centimes"
    return f"{euros_txt} {euro_label} et {cent_txt} {cent_label}"


# ============================================================
# PR / PV (inchangés)
# ============================================================
def open_pr(chantier_dir: str | Path):
    chantier = Path(chantier_dir)
    pr = _normaliser_dossier_data(chantier) / "prix_de_revient.xlsx"
    if not pr.exists():
        raise FileNotFoundError(f"PR introuvable : {pr}")
    verrouille = pr.with_name("~$" + pr.name).exists()
    if not verrouille:
        reparer_listes_pr(pr, _backup_excel_before_write)
    ouvrir_chemin(pr)
    return not verrouille


def inject_pv(chantier_dir: str | Path) -> int:
    chantier = Path(chantier_dir)
    pr_path = _normaliser_dossier_data(chantier) / "prix_de_revient.xlsx"
    pv_path = chantier / PV_SOURCE

    if not pr_path.exists():
        raise FileNotFoundError(f"PR introuvable : {pr_path}")
    if not pv_path.exists():
        raise FileNotFoundError(f"PV source introuvable : {pv_path}")

    pr_wb = load_workbook(pr_path, data_only=True)
    pr_ws = pr_wb["Chiffrage"] if "Chiffrage" in pr_wb.sheetnames else pr_wb.active


    # trouver la 1ère ligne d'article à partir de B3
    row_article = None
    for r in range(3, pr_ws.max_row + 1):
        v = pr_ws.cell(row=r, column=2).value  # colonne B
        if v is not None and str(v).strip() != "":
            row_article = r
            break

    if row_article is None:
        pr_wb.close()
        raise ValueError("Aucun ARTICLE trouvé dans la colonne B (à partir de B3).")

    article = pr_ws.cell(row=row_article, column=2).value
    designation = pr_ws.cell(row=row_article, column=3).value

    price = pr_ws["F32"].value

    pr_wb.close()



    if article is None or str(article) == "":
        raise ValueError("ARTICLE vide dans PR (B3).")
    if price is None or str(price) == "":
        raise ValueError("PRIX vide dans PR (F32).")

    article = str(article)
    designation = "" if designation is None else str(designation)

    price_f = float(price)
    price_2d = round(price_f, 2)
    letters = montant_en_lettres_fr(price_2d)

    _backup_excel_before_write(pv_path)
    wb = load_workbook(pv_path)
    ws = wb["PV"] if "PV" in wb.sheetnames else wb.active

    row_found = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        if str(v) == article:
            row_found = r
            break

    if row_found is None:
        wb.close()
        raise ValueError(f"Article introuvable dans PV source : '{article}'")

    ws.cell(row=row_found, column=2).value = designation

    g = ws.cell(row=row_found, column=7)
    g.value = price_2d
    g.number_format = "#,##0.00"

    ws.cell(row=row_found, column=8).value = letters

    _protect_formula_cells(wb, pv_path)
    wb.save(pv_path)
    wb.close()
    return 1


def sync_pv_to_copies(chantier_dir: str | Path) -> None:
    chantier = Path(chantier_dir)

    src = chantier / PV_SOURCE
    tgt_backup = chantier / PV_BACKUP
    tgt_corrige = chantier / PV_CORRIGE

    for p in (src, tgt_backup, tgt_corrige):
        if not p.exists():
            raise FileNotFoundError(f"Fichier introuvable : {p}")

    src_wb = load_workbook(src, data_only=True)
    src_ws = src_wb["PV"] if "PV" in src_wb.sheetnames else src_wb.active
    data = [row for row in src_ws.iter_rows(values_only=True)]
    src_wb.close()

    def _write(target: Path):
        _backup_excel_before_write(target)
        wb = load_workbook(target)
        ws = wb["PV"] if "PV" in wb.sheetnames else wb.active

        for r_idx, row in enumerate(data, start=1):
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if cell.__class__.__name__ == "MergedCell":
                    continue
                cell.value = val

        _protect_formula_cells(wb, target)
        wb.save(target)
        wb.close()

    _write(tgt_backup)
    _write(tgt_corrige)


# ---------------------------
# Dossiers (auto Chantier/Chantiers)
# ---------------------------
def dossier_base() -> Path:
    return _DOSSIER_BASE_MEMOIRE if _DOSSIER_BASE_MEMOIRE is not None else trouver_base()


def dossier_chantiers() -> Path:
    base = dossier_base()
    if base.name.casefold() in {"chantier", "chantiers"}:
        return base
    d1 = base / "Chantier"
    d2 = base / "Chantiers"
    if d1.exists():
        return d1
    if d2.exists():
        return d2
    d2.mkdir(parents=True, exist_ok=True)
    return d2
def _modeles_creation_chantier(base: Path):
    def cle(nom):
        return unicodedata.normalize("NFC", nom).casefold()

    dossiers = {cle(p.name): p for p in base.iterdir() if p.is_dir()}
    modeles = next((dossiers[cle(n)] for n in ("Modèles", "Modeles", "Modèle", "Modele", "Model")
                    if cle(n) in dossiers), None)
    if modeles is None:
        raise FileNotFoundError(f"Dossier modèles introuvable dans : {base}")
    fichiers = {cle(p.name): p for p in modeles.iterdir() if p.is_file()}
    etats = {}
    for type_etat, nom in (("Privé", "Etat_avancement_Privé.xlsm"),
                           ("Public", "Etat_avancement_Public.xlsm"),
                           ("Modèle actuel", "Etat_avancement_02.xlsm")):
        if cle(nom) in fichiers:
            etats[type_etat] = fichiers[cle(nom)]
    return modeles, fichiers, etats


def _creer_documents_chantier(base: Path, nom: str, chantier: dict) -> Path:
    """Prépare tous les documents avant de publier le nouveau chantier."""
    destination = base / nom
    json_path = base / f"{nom}.json"
    if nom in {"", ".", ".."} or Path(nom).name != nom:
        raise ValueError("Nom de chantier invalide.")
    if json_path.exists() or (destination.exists() and (
        not destination.is_dir() or destination.is_symlink() or any(destination.iterdir())
    )):
        raise FileExistsError("Ce chantier existe déjà.")

    def cle(nom):
        return unicodedata.normalize("NFC", nom).casefold()

    modeles, fichiers, etats = _modeles_creation_chantier(base)
    type_etat = chantier.get("type_etat", "Privé")
    if type_etat not in {"Public", "Privé", "Modèle actuel"}:
        raise ValueError("Type d'état d'avancement invalide.")
    modele_etat = etats.get(type_etat)
    chantier = dict(chantier)
    chantier["type_etat"] = type_etat
    chantier["modeles_etat_manquants"] = [t for t in ("Public", "Privé") if t not in etats]
    chantier["modele_etat_manquant"] = bool(chantier["modeles_etat_manquants"])

    with tempfile.TemporaryDirectory(prefix=".nouveau_chantier_", dir=base) as temporaire:
        preparation = Path(temporaire) / "documents"

        def ignorer_etats(src, names):
            if Path(src) == modeles:
                return [n for n in names if cle(n).startswith("etat_avancement") and cle(n).endswith(".xlsm")]
            return []

        shutil.copytree(modeles, preparation, ignore=ignorer_etats)
        # Chaque chantier reçoit les deux états ; l'utilisateur conserve celui voulu.
        for type_modele in ("Public", "Privé"):
            if type_modele in etats:
                modele = etats[type_modele]
                shutil.copy2(modele, preparation / modele.name)
        if type_etat == "Modèle actuel" and modele_etat is not None:
            shutil.copy2(modele_etat, preparation / modele_etat.name)
        data = _normaliser_dossier_data(preparation)
        data.mkdir(exist_ok=True)
        if not (data / "prix_de_revient.xlsx").exists():
            pr = fichiers.get(cle("prix_de_revient.xlsx")) or fichiers.get(cle("prix_de_revient model.xlsx"))
            if pr is None:
                raise FileNotFoundError(f"Modèle de prix de revient introuvable dans : {modeles}")
            shutil.copy2(pr, data / "prix_de_revient.xlsx")
        json_temp = Path(temporaire) / "chantier.json"
        ecrire_json(json_temp, chantier)
        # Un dossier strictement vide peut provenir d'une ancienne tentative ratée.
        if destination.exists():
            destination.rmdir()
        preparation.rename(destination)
        try:
            json_temp.rename(json_path)
        except Exception:
            destination.rename(preparation)
            raise
    return json_path


def ouvrir_doc(self, nom_fichier):
    sel = self.tree.selection()
    if not sel:
        messagebox.showwarning("Sélection", "Sélectionne un chantier.")
        return

    item = self.tree.item(sel[0])

    # 1) Nom sélectionné (text si présent, sinon values[0])
    chantier_sel = (item.get("text") or "").strip()
    if not chantier_sel:
        vals = item.get("values", [])
        chantier_sel = str(vals[0]).strip() if vals else ""

    base = dossier_chantiers()

    # 2) On cherche le dossier chantier qui existe vraiment
    dossier = base / chantier_sel
    if not dossier.exists():
        def norm(s):
            return s.lower().strip().replace(" ", "_").replace("-", "_")

        cible = norm(chantier_sel)
        dossier = next(
            (p for p in base.iterdir() if p.is_dir() and norm(p.name) == cible),
            None
        )

    if not dossier or not Path(dossier).exists():
        messagebox.showerror(
            "Introuvable",
            f"Dossier chantier introuvable pour :\n{chantier_sel}\n\nBase : {base}"
        )
        return

    chemin = Path(dossier) / nom_fichier

    if not chemin.exists():
        # Si on n'a pas trouvé le fichier, on tente l'autre dossier chantier
        # (ex: "soumission_chapelette" vs "Soumission Chapelette")
        def norm(s):
            return str(s).lower().strip().replace(" ", "_").replace("-", "_")

        cible = norm(chantier_sel)
        for p in base.iterdir():
            if p.is_dir() and norm(p.name) == cible:
                alt = p / nom_fichier
                if alt.exists():
                   chemin = alt
                   break

    if not chemin.exists():
        messagebox.showerror("Introuvable", f"Fichier absent :\n{chemin}")
        return

    if not self._preparer_ouverture_document(chemin):
        return

    ouvrir_chemin(chemin)
    if self._doit_rappeler_pdf_historique(chemin):
        self._planifier_rappel_pdf_historique_ouverture()

def ouvrir_chemin(path: str | Path) -> None:
    chemin = Path(path)
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(chemin)])
        elif os.name == "nt":
            os.startfile(str(chemin))
        else:
            subprocess.Popen(["xdg-open", str(chemin)])
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier ou le dossier.\n{e}")


def ouvrir_dossier(path: Path) -> None:
    ouvrir_chemin(path)


# ---------------------------
# Lecture chantier JSON
# ---------------------------
def lire_json(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def ecrire_json(path: Path, data: dict) -> None:
    path = Path(path)
    contenu = json.dumps(data, ensure_ascii=False, indent=2)
    fd, nom = tempfile.mkstemp(prefix=".fiche_", suffix=".json", dir=path.parent)
    temporaire = Path(nom)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(contenu)
            f.flush()
            os.fsync(f.fileno())
        if path.exists():
            shutil.copymode(path, temporaire)
        os.replace(temporaire, path)
    finally:
        temporaire.unlink(missing_ok=True)


def infos_chantier(chantier: dict) -> dict:
    return {
        "nom": chantier.get("nom", chantier.get("chantier", "")) or "",
        "client": chantier.get("client", "") or "",
        "etat": chantier.get("etat", chantier.get("type_travaux", "")) or "",
        "avancement": chantier.get("avancement", chantier.get("avancement_pct", 0)) or 0,
        "debut": chantier.get("date_debut", chantier.get("debut_travaux", "")) or "",
        "fin": chantier.get("date_fin", chantier.get("fin_prevue", chantier.get("date_fin_prevue", ""))) or "",
    }


# ---------------------------
# Fenêtre Bordereau (visu JSON)
# ---------------------------
def afficher_bordereau(parent: tk.Tk, chantier: dict, titre: str = "Bordereau") -> None:
    bord = chantier.get("bordereau", {})
    articles = bord.get("articles", {})

    if not articles:
        messagebox.showwarning("Bordereau", "Aucun bordereau dans ce chantier.")
        return

    win = tk.Toplevel(parent)
    win.title(titre)
    win.geometry("1200x650")

    cols = ("Article", "Libellé", "Unité", "Qté", "PU", "Total")

    tree = ttk.Treeview(win, columns=cols, show="headings")
    vsb = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(win, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    for c in cols:
        tree.heading(c, text=c)
        if c == "Libellé":
            tree.column(c, width=700, anchor="w")
        elif c == "Article":
            tree.column(c, width=160, anchor="w")
        elif c == "Unité":
            tree.column(c, width=80, anchor="w")
        else:
            tree.column(c, width=120, anchor="e")

    def v(x):
        return "" if x is None else x

    for article_id in sorted(articles.keys()):
        a = articles[article_id] or {}
        tree.insert(
            "",
            "end",
            values=(
                v(a.get("article_id", article_id)),
                v(a.get("libelle", "")),
                v(a.get("unite", "")),
                v(a.get("quantite", 0)),
                v(a.get("pu_vente", 0)),
                v(a.get("total_vente", 0)),
            ),
        )

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    win.grid_rowconfigure(0, weight=1)
    win.grid_columnconfigure(0, weight=1)


# ---------------------------
# App
# ---------------------------
class HorizonChantierApp(tk.Tk):


    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1400x800")
        self.minsize(1300, 950)
        self.dernier_fichier_excel_ouvert: Path | None = None
        self.rappel_pdf_historique_fenetre = None
        self._rappel_pdf_historique_after_id = None

        self._build_ui()
        self.refresh_liste()

    def report_callback_exception(self, exc_type, exc_value, traceback):
        super().report_callback_exception(exc_type, exc_value, traceback)
        messagebox.showerror("Opération interrompue", str(exc_value), parent=self)

    def _nom_chantier_selectionne(self) -> str | None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Chantier", "Sélectionne un chantier dans la liste.")
            return None
        vals = self.tree.item(sel[0]).get("values", [])
        return str(vals[0]).strip() if vals else ""

    def _ouvrir_fichier_chantier(self, nom_fichier: str) -> None:
        nom_chantier = self._nom_chantier_selectionne()
        if not nom_chantier:
            return
        chemin = _chemin_fichier_chantier(self._dossier_chantier_selectionne(), nom_fichier)
        if not chemin.exists():
            messagebox.showerror("Introuvable", f"Fichier absent :\n{chemin}")
            return
        if not self._preparer_ouverture_document(chemin):
            return
        ouvrir_chemin(chemin)
        if self._doit_rappeler_pdf_historique(chemin):
            self._planifier_rappel_pdf_historique_ouverture()

    def _ouvrir_premier_fichier_chantier(self, motif: str) -> None:
        nom_chantier = self._nom_chantier_selectionne()
        if not nom_chantier:
            return
        try:
            chemin = (_selectionner_etat(self._dossier_chantier_selectionne())
                      if motif == "Etat_avancement*.xlsm"
                      else next(self._dossier_chantier_selectionne().glob(motif)))
        except (StopIteration, FileNotFoundError, ValueError) as e:
            messagebox.showerror("Introuvable", str(e) or f"Fichier absent :\n{motif}")
            return
        if not self._preparer_ouverture_document(chemin):
            return
        ouvrir_chemin(chemin)
        if self._doit_rappeler_pdf_historique(chemin):
            self._planifier_rappel_pdf_historique_ouverture()

    def _dossier_depuis_json(self, p: Path) -> Path:
        try:
            chantier = lire_json(p)
            info = infos_chantier(chantier)
            nom = str(info.get("nom") or "").strip()
            if nom:
                dossier = p.parent / nom
                if dossier.is_dir() and Path(nom).name == nom:
                    return dossier
        except Exception:
            pass
        return p.parent / p.stem

    def _memoriser_fichier_excel(self, chemin: Path, afficher_rappel: bool = True) -> None:
        if chemin.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            self.dernier_fichier_excel_ouvert = chemin

    def _preparer_ouverture_document(self, chemin: Path) -> bool:
        if chemin.is_dir():
            if not self._dossier_contient_fichier_excel(chemin):
                return True
            return self._confirmer_ouverture_document_excel()

        if chemin.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            return True
        if not self._confirmer_ouverture_document_excel():
            return False
        self._memoriser_fichier_excel(chemin)
        return True

    def _dossier_contient_fichier_excel(self, dossier: Path) -> bool:
        try:
            return any(
                fichier.is_file() and fichier.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
                for fichier in dossier.rglob("*")
            )
        except Exception:
            return False

    def _doit_rappeler_pdf_historique(self, chemin: Path) -> bool:
        if chemin.is_dir():
            return self._dossier_contient_fichier_excel(chemin)
        return chemin.suffix.lower() in {".xlsx", ".xlsm", ".xls"}

    def _planifier_rappel_pdf_historique_ouverture(self) -> None:
        if self._rappel_pdf_historique_after_id is not None:
            return
        self._rappel_pdf_historique_after_id = self.after(1200, self._afficher_rappel_pdf_historique_ouverture)

    def _afficher_rappel_pdf_historique_ouverture(self) -> None:
        self._rappel_pdf_historique_after_id = None
        if self.rappel_pdf_historique_fenetre and self.rappel_pdf_historique_fenetre.winfo_exists():
            self.rappel_pdf_historique_fenetre.lift()
            return

        win = tk.Toplevel(self)
        win.title("Sauvegarde PDF Historique")
        win.resizable(False, False)
        win.geometry("560x340")
        self.rappel_pdf_historique_fenetre = win
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass

        cadre = ttk.Frame(win, padding=32)
        cadre.pack(fill="both", expand=True)

        tk.Label(
            cadre,
            text="⚠️",
            font=("Helvetica", 54, "bold"),
            fg="#b45309",
        ).pack(pady=(0, 8))
        tk.Label(
            cadre,
            text="ATTENTION - SAUVEGARDE OBLIGATOIRE",
            font=("Helvetica", 19, "bold"),
            wraplength=480,
            justify="center",
            fg="#7c2d12",
        ).pack(pady=(0, 18))
        tk.Label(
            cadre,
            text=(
                "N'oubliez pas d'enregistrer vos modifications dans Excel.\n\n"
                "Après vos modifications importantes, créez votre PDF Historique.\n\n"
                "Ce PDF est votre copie de sécurité en cas de modification ou d'erreur ultérieure."
            ),
            font=("Helvetica", 14, "bold"),
            wraplength=490,
            justify="center",
        ).pack(pady=(0, 26))

        def fermer():
            self.rappel_pdf_historique_fenetre = None
            win.destroy()

        ttk.Button(cadre, text="✓ OK, j'ai compris", command=fermer).pack()

        win.protocol("WM_DELETE_WINDOW", fermer)
        win.update_idletasks()
        largeur_ecran = win.winfo_screenwidth()
        hauteur_ecran = win.winfo_screenheight()
        x = self.winfo_rootx() + self.winfo_width() + 16
        if x + win.winfo_width() > largeur_ecran - 24:
            x = largeur_ecran - win.winfo_width() - 24
        y = self.winfo_rooty() + 80
        if y + win.winfo_height() > hauteur_ecran - 24:
            y = max(24, hauteur_ecran - win.winfo_height() - 24)
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _confirmer_ouverture_document_excel(self) -> bool:
        confirmation = tk.BooleanVar(value=False)

        win = tk.Toplevel(self)
        win.title("Règles de travail obligatoires")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        cadre = ttk.Frame(win, padding=28)
        cadre.pack(fill="both", expand=True)

        tk.Label(
            cadre,
            text="⚠️ ATTENTION - RÈGLES DE TRAVAIL OBLIGATOIRES",
            font=("Helvetica", 18, "bold"),
            fg="#7c2d12",
            wraplength=620,
            justify="center",
        ).pack(pady=(0, 18))
        tk.Label(
            cadre,
            text=(
                "Vous allez ouvrir un document de travail.\n\n"
                "Avant de continuer, vous confirmez avoir compris les règles suivantes :\n\n"
                "• Enregistrer votre fichier Excel après chaque modification.\n"
                "• Créer votre PDF Historique après vos modifications importantes.\n"
                "• Respecter les cellules protégées et les zones prévues dans les modèles Horizon Chantier."
            ),
            font=("Helvetica", 13, "bold"),
            wraplength=620,
            justify="left",
        ).pack(pady=(0, 24))

        def confirmer():
            confirmation.set(True)
            win.destroy()

        ttk.Button(
            cadre,
            text="✓ J'ai compris - Ouvrir le document",
            command=confirmer,
        ).pack()

        win.protocol("WM_DELETE_WINDOW", lambda: win.destroy())
        win.bind("<Escape>", lambda _event: win.destroy())
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - win.winfo_height()) // 2
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.wait_window(win)
        return confirmation.get()

    def sauvegarde_pdf_historique(self) -> bool:
        chemin = _chemin_classeur_excel_actif() or self.dernier_fichier_excel_ouvert
        if not chemin or not chemin.exists():
            messagebox.showwarning(
                "Sauvegarde PDF Historique",
                "Aucun fichier Excel ouvert n'a été trouvé."
            )
            return False
        self._memoriser_fichier_excel(chemin, afficher_rappel=False)

        try:
            dossier_chantier = _chantier_root_for_path(chemin)
            historique = dossier_chantier / "Historique_Sauvegardes"
            historique.mkdir(exist_ok=True)

            nom_source = re.sub(r"\.[^.]+$", "", chemin.name)
            nom_pdf = f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S_%f')}_{nom_source}.pdf"
            destination = historique / nom_pdf

            _export_excel_to_pdf(chemin, destination)
            messagebox.showinfo(
                "Sauvegarde PDF Historique",
                f"PDF historique créé :\n{destination}"
            )
            return True
        except Exception as e:
            messagebox.showerror("Sauvegarde PDF Historique", f"Impossible de créer le PDF.\n{e}")
            return False

    def _confirmer_cloture_etat(self) -> bool:
        confirmation = tk.BooleanVar(value=False)

        win = tk.Toplevel(self)
        win.title("Clôture de l'État")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        cadre = ttk.Frame(win, padding=28)
        cadre.pack(fill="both", expand=True)

        tk.Label(
            cadre,
            text="⚠️",
            font=("Helvetica", 54, "bold"),
            fg="#b42318",
        ).pack(pady=(0, 8))
        tk.Label(
            cadre,
            text="ATTENTION - CLÔTURE IRRÉVERSIBLE",
            font=("Helvetica", 20, "bold"),
            fg="#b42318",
            wraplength=580,
            justify="center",
        ).pack(pady=(0, 18))
        tk.Label(
            cadre,
            text=(
                "Vous allez clôturer l'État d'avancement.\n\n"
                "Cette opération effectue la mise à zéro des quantités du mois "
                "et il sera impossible de revenir à l'état précédent.\n\n"
                "Une copie datée de l'état sera enregistrée automatiquement "
                "dans le dossier Sauvegarde avant la clôture."
            ),
            font=("Helvetica", 13, "bold"),
            wraplength=590,
            justify="center",
        ).pack(pady=(0, 24))

        boutons = ttk.Frame(cadre)
        boutons.pack(fill="x")

        def annuler():
            confirmation.set(False)
            win.destroy()

        def confirmer():
            confirmation.set(True)
            win.destroy()

        bouton_annuler = ttk.Button(boutons, text="❌ Annuler", command=annuler)
        bouton_annuler.pack(side="left")
        ttk.Button(
            boutons,
            text="✅ Oui, je confirme la clôture",
            command=confirmer,
        ).pack(side="right", padx=(20, 0))

        win.protocol("WM_DELETE_WINDOW", annuler)
        bouton_annuler.focus_set()
        win.bind("<Return>", lambda _event: None)
        win.bind("<Escape>", lambda _event: annuler())
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - win.winfo_height()) // 2
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.wait_window(win)
        return confirmation.get()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)
        style = ttk.Style()
        style.configure("TitreBloc.TLabel",
                font=("Helvetica", 12, "bold"),
                foreground="#1f4e79")
                
        title = ttk.Label(root, text=APP_NAME, font=("Helvetica", 22, "bold"))
        title.pack(anchor="w")

        try:
            emplacement = str(dossier_chantiers())
        except (OSError, ValueError) as e:
            emplacement = str(e)
        ttk.Button(root, text="Choisir le dossier des chantiers", command=self.choisir_dossier_chantiers).pack(anchor="w")
        self.lbl_path = ttk.Label(root, text=f"Emplacement: {emplacement}")
        self.lbl_path.pack(fill="x", pady=(4, 8))

        cols = ("Chantier", "Client", "État", "Avancement", "Début", "Fin prévue")

        zone = ttk.Frame(root)
        zone.pack(fill="both", expand=True)

        zone.columnconfigure(0, weight=0)
        zone.columnconfigure(1, weight=1)
        zone.columnconfigure(2, weight=0)
        zone.rowconfigure(0, weight=1)

        frame_gauche = ttk.Frame(zone)
        frame_gauche.grid(row=0, column=0, sticky="ns", padx=(0, 12))

        frame_centre = ttk.Frame(zone)
        frame_centre.grid(row=0, column=1, sticky="nsew")

        frame_droite = ttk.Frame(zone)
        frame_droite.grid(row=0, column=2, sticky="ns", padx=(12, 0))

        self.tree = ttk.Treeview(frame_centre, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=c)
            if c == "Chantier":
                self.tree.column(c, width=260, anchor="w")
            elif c == "Client":
                self.tree.column(c, width=180, anchor="w")
            elif c == "État":
                self.tree.column(c, width=120, anchor="w")
            elif c == "Avancement":
                self.tree.column(c, width=110, anchor="e")
            else:
                self.tree.column(c, width=120, anchor="w")

        self.tree.pack(fill="both", expand=True)

        # -------------------------
        # GAUCHE
        # -------------------------

        ttk.Separator(frame_gauche).pack(fill="x", pady=10)
        tk.Label(frame_gauche, text="AIDE",
            font=("Helvetica", 15, "bold"),
            fg="#C62828").pack(anchor="w", pady=(2, 4))
        ttk.Separator(frame_gauche).pack(fill="x", pady=(0, 8))
        ttk.Button(frame_gauche, text="Aide / Dépannage", command=self.ouvrir_aide).pack(fill="x", pady=3)

        ttk.Separator(frame_gauche).pack(fill="x", pady=10)
        tk.Label(frame_gauche, text="GESTION",
            font=("Helvetica", 14, "bold"),
            fg="#0052cc").pack(anchor="w", pady=(0, 2))
        ttk.Separator(frame_gauche).pack(fill="x", pady=(0, 8))

        ttk.Button(frame_gauche, text="📄 Ouvrir dans le logiciel", command=self.ouvrir_chantier).pack(fill="x", pady=3)
        ttk.Button(frame_gauche, text="➕ Nouveau chantier", command=self.nouveau_chantier).pack(fill="x", pady=3)
        ttk.Button(frame_gauche, text="✏️ Modifier", command=self.modifier_chantier).pack(fill="x", pady=3)
        ttk.Button(frame_gauche, text="👥 Voir les clients", command=self.voir_clients).pack(fill="x", pady=3)
        ttk.Button(frame_gauche, text="🗑️ Supprimer", command=self.supprimer_chantier).pack(fill="x", pady=3)
        ttk.Button(frame_gauche, text="🔄 Rafraîchir", command=self.refresh_liste).pack(fill="x", pady=3)
        tk.Button(
            frame_gauche,
            text="Sauvegarde PDF Historique",
            command=self.sauvegarde_pdf_historique,
            bg="yellow",
            background="yellow",
            fg="black",
            foreground="black",
            activebackground="yellow",
            activeforeground="black",
            font=("Helvetica", 12, "bold"),
            relief="raised",
            bd=2,
            highlightbackground="yellow",
            highlightcolor="yellow",
            highlightthickness=2,
            padx=8,
            pady=6,
        ).pack(fill="x", pady=(6, 4))

        ttk.Separator(frame_gauche).pack(fill="x", pady=10)

        tk.Label(frame_gauche, text="DOCUMENTS",
         font=("Helvetica", 15, "bold"),
         fg="#0052cc").pack(anchor="w", pady=(2, 4))
        ttk.Separator(frame_gauche).pack(fill="x", pady=(0, 8))

        ttk.Button(frame_gauche, text="📁 Ce chantier", command=self.dossier_du_chantier).pack(fill="x", pady=3)

        ttk.Button(
            frame_gauche,
            text="🔎 Dossier administratif",
            command=lambda: self._ouvrir_fichier_chantier("Administratif")
        ).pack(fill="x", pady=3)

        ttk.Button(
            frame_gauche,
            text="📄 Cahier des charges administratif",
            command=lambda: self._ouvrir_fichier_chantier("Cahier_administratif.pdf")
        ).pack(fill="x", pady=3)

        ttk.Button(
           frame_gauche,
            text="🛠 Cahier des charges technique",
            command=lambda: self._ouvrir_fichier_chantier("Cahier_technique.pdf")
        ).pack(fill="x", pady=3)

        ttk.Button(
            frame_gauche,
            text="📐 Postes / Métré",
            command=lambda: self._ouvrir_fichier_chantier("Metre_detaille.pdf")
        ).pack(fill="x", pady=3)

        ttk.Button(
            frame_gauche,
            text="🦺 Plan de sécurité (PSS)",
            command=lambda: self._ouvrir_fichier_chantier("PSS.pdf")
        ).pack(fill="x", pady=3)

        ttk.Button(
            frame_gauche,
            text="📄 Décompte intempéries",
            command=lambda: self._ouvrir_fichier_chantier("Décompte_intempéries.doc")
        ).pack(fill="x", pady=3)

        # -------------------------
        # DROITE
        # -------------------------

        

        ttk.Separator(frame_droite).pack(fill="x", pady=10)

        tk.Label(frame_droite, text="EXECUTION",
         font=("Helvetica", 15, "bold"),
         fg="#0052cc").pack(anchor="w", pady=(2, 4))
        ttk.Separator(frame_droite).pack(fill="x", pady=(0, 8))

        ttk.Button(
            frame_droite,
            text="📅 Planning d’exécution",
            command=lambda: self._ouvrir_fichier_chantier("Planning_execution.xlsx")
        ).pack(fill="x", pady=3)

        ttk.Separator(frame_droite).pack(fill="x", pady=10)

        tk.Label(frame_droite, text="CALCULS",
         font=("Helvetica", 15, "bold"),
         fg="#0052cc").pack(anchor="w", pady=(2, 4))
        ttk.Separator(frame_droite).pack(fill="x", pady=(0, 8))

        ttk.Button(
            frame_droite,
            text="🔎 Coût de la sécurité",
            command=lambda: self._ouvrir_fichier_chantier("Cout_securite.xlsx")
        ).pack(fill="x", pady=3)

        ttk.Button(
            frame_droite,
            text="📐 Formule de révision",
            command=lambda: self._ouvrir_fichier_chantier("Formule_révision.xlsm")
        ).pack(fill="x", pady=3)

        ttk.Button(
            frame_droite,
            text="💰 Prix de revient",
            command=self.prix_revient
        ).pack(fill="x", pady=3)

        ttk.Button(frame_droite, text="🔍 Vérifier PR", command=self.verifier_pr).pack(fill="x", pady=3)
        ttk.Button(frame_droite, text="⚙️ Calcul / Reca", command=self.calcul_reca).pack(fill="x", pady=3)

        ttk.Separator(frame_droite).pack(fill="x", pady=10)

        tk.Label(frame_droite, text="ETAT D'AVANCEMENT",
         font=("Helvetica", 15, "bold"),
         fg="#0052cc").pack(anchor="w", pady=(2, 4))
        ttk.Separator(frame_droite).pack(fill="x", pady=(0, 8))

        ttk.Button(
            frame_droite,
            text="📊 État d'avancement",
            command=lambda: self._ouvrir_premier_fichier_chantier("Etat_avancement*.xlsm")
        ).pack(fill="x", pady=3)

        ttk.Button(frame_droite, text="📊 Calcul état", command=self.calcul_etat_avancement).pack(fill="x", pady=3)
        ttk.Button(frame_droite, text="Exporter les prix vers l’original",
                   command=self.exporter_prix_bordereau_public).pack(fill="x", pady=3)
        btn_cloturer = tk.Label(
            frame_droite,
            text="♻️ Clôturer état",
            bg="#EF8F8F",
            fg="black",
            font=("Helvetica", 14, "bold"),
            relief="flat",
            bd=0,
            padx=8,
            pady=6,
            cursor="hand2",
        )
        btn_cloturer.bind("<Button-1>", lambda _event: self.mise_a_zero_etat())
        btn_cloturer.pack(fill="x", pady=3)

        ttk.Separator(frame_droite).pack(fill="x", pady=10)

        tk.Label(frame_droite, text="PILOTAGE",
         font=("Helvetica", 15, "bold"),
         fg="#0052cc").pack(anchor="w", pady=(2, 4))
        ttk.Separator(frame_droite).pack(fill="x", pady=(0, 8))

        ttk.Button(frame_droite, text="🧭 Pilotage", command=self.pilotage_chantier).pack(fill="x", pady=3)
        ttk.Button(frame_droite, text="📊 Pilotage délai", command=self.pilotage_delai).pack(fill="x", pady=3)
        ttk.Button(frame_droite, text="⏱ Rendements", command=self.rendements_chantier).pack(fill="x", pady=3)
       
       

    def ouvrir_aide(self) -> None:
        existing = getattr(self, "_help_window", None)
        if existing and existing.winfo_exists():
            existing.deiconify()
            existing.lift()
            existing.focus_force()
            return

        win = tk.Toplevel(self)
        self._help_window = win
        win.title("Aide / Mode d’emploi - Horizon Chantier")
        win.geometry("900x600")
        win.minsize(820, 520)
        win.transient(self)
        topmost_var = tk.BooleanVar(value=True)

        def appliquer_toujours_visible() -> None:
            try:
                win.attributes("-topmost", topmost_var.get())
            except Exception:
                pass

        def fermer() -> None:
            self._help_window = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", fermer)
        appliquer_toujours_visible()

        header = ttk.Frame(win, padding=(18, 16, 18, 10))
        header.pack(fill="x")

        ttk.Label(
            header,
            text="Aide / Mode d’emploi - Horizon Chantier",
            font=("Helvetica", 20, "bold"),
            foreground="#163A59",
        ).pack(anchor="w")

        ttk.Label(
            header,
            text="Choisir le guide à afficher puis suivre les consignes sans modifier les formules Excel.",
            font=("Helvetica", 10),
            foreground="#4F6475",
        ).pack(anchor="w", pady=(4, 0))

        zone_choix = ttk.Frame(header)
        zone_choix.pack(anchor="w", pady=(10, 0))
        ttk.Label(zone_choix, text="Aide :", font=("Helvetica", 10, "bold")).pack(side="left", padx=(0, 8))
        aide_var = tk.StringVar(value="Mode d’emploi")
        choix_aide = ttk.Combobox(
            zone_choix,
            textvariable=aide_var,
            values=["Mode d’emploi", "🔧 Dépannage"],
            state="readonly",
            width=24,
        )
        choix_aide.pack(side="left")

        ttk.Separator(win).pack(fill="x")

        corps = ttk.Frame(win, padding=(14, 12, 14, 8))
        corps.pack(fill="both", expand=True)
        corps.columnconfigure(0, weight=1)
        corps.rowconfigure(0, weight=1)

        texte = tk.Text(
            corps,
            wrap="word",
            font=("Helvetica", 11),
            relief="flat",
            bd=0,
            padx=22,
            pady=18,
            spacing1=2,
            spacing3=7,
            background="#FFFFFF",
            foreground="#1E1E1E",
            insertwidth=0,
        )
        texte.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(corps, orient="vertical", command=texte.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        texte.configure(yscrollcommand=scroll.set)

        texte.tag_configure("titre", font=("Helvetica", 19, "bold"), foreground="#163A59", spacing3=14)
        texte.tag_configure("intro", font=("Helvetica", 11), foreground="#405261", spacing3=10)
        texte.tag_configure("section", font=("Helvetica", 14, "bold"), foreground="#1F4E79", spacing1=14, spacing3=8)
        texte.tag_configure("texte", font=("Helvetica", 11), foreground="#1E1E1E", spacing3=6)
        texte.tag_configure("liste", font=("Helvetica", 11), foreground="#1E1E1E", lmargin1=22, lmargin2=42, spacing3=4)
        texte.tag_configure("regle", font=("Helvetica", 12, "bold"), foreground="#8A1C1C", background="#FFF3CD", spacing1=8, spacing3=10, lmargin1=10, lmargin2=10)
        texte.tag_configure("formule", font=("Courier", 11, "bold"), foreground="#163A59", spacing3=4, lmargin1=22, lmargin2=42)
        texte.tag_configure("ok", font=("Helvetica", 11, "bold"), foreground="#1E6B34", spacing3=4)
        texte.tag_configure("attention", font=("Helvetica", 11, "bold"), foreground="#9C4F00", spacing3=4)
        texte.tag_configure("danger", font=("Helvetica", 11, "bold"), foreground="#B42318", spacing3=4)
        texte.tag_configure("aide", font=("Helvetica", 11, "bold"), foreground="#5B3F99", spacing3=4)
        texte.tag_configure("fin", font=("Helvetica", 10, "italic"), foreground="#4F6475", spacing1=8, spacing3=10)

        contenu_mode_emploi = [
            ("titre", "Mode d’emploi par bouton - Horizon Chantier\n"),
            ("intro", "Ce guide explique l’utilisation normale du logiciel Horizon Chantier. Il reprend les boutons comme ils apparaissent dans l’application.\n"),
            ("regle", "RÈGLE PRINCIPALE : ne jamais modifier les formules, validations, feuilles, colonnes ou structures des fichiers Excel.\n"),
            ("section", "\nRègle générale de saisie\n"),
            ("regle", "Dans tous les fichiers Excel utilisés par Horizon Chantier, les cellules grisées sont les zones de saisie prévues pour l’utilisateur.\n"),
            ("liste", "✅ L’utilisateur peut uniquement compléter ou modifier ces zones grisées.\n"),
            ("liste", "❌ Toutes les autres cellules contiennent des calculs, des formules ou des informations générées par le logiciel.\n"),
            ("liste", "❌ Elles ne doivent jamais être modifiées manuellement.\n"),
            ("liste", "✅ L’utilisateur n’a pas à contrôler ou corriger les calculs du logiciel.\n"),
            ("liste", "✅ Il doit uniquement introduire les données demandées dans les zones prévues.\n"),
            ("attention", "⚠️ Cette règle est une base obligatoire pour utiliser correctement Horizon Chantier.\n"),

            ("section", "\nAIDE\n"),
            ("section", "\nAide / Dépannage\n"),
            ("texte", "Ouvre la fenêtre contenant le mode d’emploi.\n"),

            ("section", "\nGESTION\n"),
            ("section", "\n📄 Ouvrir dans le logiciel\n"),
            ("texte", "Ouvre la fiche du chantier sélectionné dans Horizon Chantier.\n"),
            ("liste", "✅ Sélectionner un chantier dans la liste avant de cliquer.\n"),
            ("liste", "✅ Utiliser ce bouton pour consulter les informations du chantier dans le logiciel.\n"),

            ("section", "\n➕ Nouveau chantier\n"),
            ("texte", "Crée un nouveau chantier complet avec les documents de travail prévus.\n"),
            ("liste", "✅ Renseigner le client, le nom du chantier et les informations demandées.\n"),
            ("liste", "✅ Horizon prépare automatiquement le dossier et les documents du chantier.\n"),
            ("liste", "❌ Ne pas créer manuellement un dossier chantier à la place du logiciel.\n"),

            ("section", "\n✏️ Modifier\n"),
            ("texte", "Permet de modifier les informations générales du chantier sélectionné.\n"),
            ("liste", "✅ Utiliser ce bouton pour mettre à jour les informations administratives du chantier.\n"),
            ("liste", "❌ Ne pas utiliser ce bouton pour modifier les fichiers Excel.\n"),

            ("section", "\n🗑️ Supprimer\n"),
            ("texte", "Supprime le chantier sélectionné si l’action est confirmée.\n"),
            ("liste", "⚠️ À utiliser uniquement si la suppression est réellement voulue.\n"),
            ("liste", "❌ Ne pas supprimer un chantier actif ou contenant des documents utiles.\n"),

            ("section", "\n🔄 Rafraîchir\n"),
            ("texte", "Recharge la liste des chantiers affichés dans Horizon Chantier.\n"),
            ("liste", "✅ À utiliser après création, modification ou si un chantier n’apparaît pas immédiatement.\n"),

            ("section", "\nSauvegarde PDF Historique\n"),
            ("texte", "Crée le PDF de sécurité du document Excel sur lequel vous venez de travailler.\n"),
            ("liste", "✅ Ouvrir le fichier Excel du chantier.\n"),
            ("liste", "✅ Modifier uniquement les zones prévues dans Excel.\n"),
            ("liste", "✅ Enregistrer le fichier dans Excel.\n"),
            ("liste", "✅ Revenir dans Horizon Chantier.\n"),
            ("liste", "✅ Cliquer sur Sauvegarde PDF Historique.\n"),
            ("liste", "✅ Le PDF est créé dans Historique_Sauvegardes avec la date, l’heure et le nom du document.\n"),
            ("attention", "⚠️ Si plusieurs fichiers Excel sont ouverts, afficher d’abord le bon document Excel avant de lancer la sauvegarde PDF.\n"),

            ("section", "\nDOCUMENTS\n"),
            ("section", "\n📁 Ce chantier\n"),
            ("texte", "Ouvre le dossier complet du chantier sélectionné.\n"),
            ("liste", "✅ Cliquer sur Ce chantier.\n"),
            ("liste", "✅ Choisir le fichier Excel ou document souhaité dans le dossier.\n"),
            ("liste", "✅ Travailler dans Excel uniquement dans les zones prévues.\n"),
            ("liste", "✅ Enregistrer le fichier.\n"),
            ("liste", "✅ Revenir dans Horizon Chantier et cliquer sur Sauvegarde PDF Historique si le fichier Excel a été modifié.\n"),

            ("section", "\n🔎 Dossier administratif\n"),
            ("texte", "Ouvre le dossier Administratif du chantier.\n"),
            ("liste", "✅ Utiliser ce bouton pour accéder aux documents administratifs du chantier.\n"),

            ("section", "\n📄 Cahier des charges administratif\n"),
            ("texte", "Ouvre le fichier Cahier_administratif.pdf du chantier.\n"),

            ("section", "\n🛠 Cahier des charges technique\n"),
            ("texte", "Ouvre le fichier Cahier_technique.pdf du chantier.\n"),

            ("section", "\n📐 Postes / Métré\n"),
            ("texte", "Ouvre le fichier Metre_detaille.pdf du chantier.\n"),

            ("section", "\n🦺 Plan de sécurité (PSS)\n"),
            ("texte", "Ouvre le fichier PSS.pdf du chantier.\n"),

            ("section", "\n📄 Décompte intempéries\n"),
            ("texte", "Ouvre le document Décompte_intempéries.doc du chantier.\n"),

            ("section", "\nEXECUTION\n"),
            ("section", "\n📅 Planning d’exécution\n"),
            ("texte", "Ouvre le planning d’exécution du chantier.\n"),
            ("liste", "✅ Modifier uniquement les zones prévues dans Excel.\n"),
            ("liste", "✅ Enregistrer puis créer une Sauvegarde PDF Historique si le planning a été modifié.\n"),

            ("section", "\nCALCULS\n"),
            ("section", "\n🔎 Coût de la sécurité\n"),
            ("texte", "Ouvre le document de coût de la sécurité du chantier.\n"),
            ("liste", "✅ Encoder uniquement les zones de saisie prévues.\n"),
            ("liste", "✅ Enregistrer puis créer une Sauvegarde PDF Historique après modification.\n"),

            ("section", "\n📐 Formule de révision\n"),
            ("texte", "Ouvre le document de formule de révision du chantier.\n"),
            ("liste", "✅ Utiliser les zones prévues pour les informations de révision.\n"),
            ("liste", "✅ La révision globale est gérée dans un document séparé, en colonne E.\n"),
            ("liste", "✅ La révision du mois reste encodée manuellement.\n"),
            ("liste", "❌ Aucune automatisation n’est prévue pour la révision mensuelle.\n"),
            ("liste", "❌ Ne pas modifier les formules, la structure ou les protections du document.\n"),

            ("section", "\n💰 Prix de revient\n"),
            ("texte", "Ouvre le Prix de revient du chantier.\n"),
            ("liste", "✅ Encoder les informations de PR uniquement dans les zones prévues.\n"),
            ("liste", "✅ Enregistrer le PR avant Vérifier PR ou Calcul / Reca.\n"),
            ("liste", "❌ Ne pas modifier les formules, les feuilles de bibliothèque ou les listes déroulantes.\n"),

            ("section", "\n📖 Cahier technique dans le PR\n"),
            ("texte", "Dans un PR équipé, les pages du cahier des charges sont intégrées comme images lisibles dans des onglets CSC du même classeur Excel.\n"),
            ("liste", "✅ Sous la désignation du poste, à droite du numéro d’article, cliquer sur le lien Cahier technique.\n"),
            ("liste", "✅ Lire les pages dans l’onglet CSC correspondant, faire défiler ou ajuster le zoom Excel si nécessaire.\n"),
            ("liste", "✅ Cliquer sur Retour au PR en haut de l’onglet pour revenir au poste.\n"),
            ("texte", "Le lien remplace les longs textes copiés-collés manuellement. Les photos et les saisies du poste sont conservées. Le bouton Cahier des charges technique d’Horizon ouvre toujours le PDF complet séparément.\n"),
            ("liste", "✅ Préparer et vérifier les correspondances articles/pages pour chaque chantier dans les dossiers de travail, avant transfert sur le partage.\n"),
            ("liste", "✅ Transférer le classeur complet : les pages et les liens internes restent dans le fichier lors de son déplacement entre Windows et macOS.\n"),
            ("texte", "Mariemont est équipé. Ce fonctionnement est retenu pour les autres chantiers, mais leur préparation reste à effectuer chantier par chantier ; la reconnaissance et l’intégration automatiques ne sont pas encore disponibles.\n"),

            ("section", "\n🔍 Vérifier PR\n"),
            ("texte", "Lance le contrôle du Prix de revient enregistré.\n"),
            ("liste", "✅ À lancer après avoir enregistré le PR dans Excel.\n"),
            ("liste", "✅ Lire le résultat affiché par Horizon.\n"),
            ("liste", "✅ Reprendre uniquement les zones de saisie prévues si une correction métier est nécessaire.\n"),
            ("liste", "❌ Ne pas modifier une formule pour faire disparaître une erreur.\n"),

            ("section", "\n⚙️ Calcul / Reca\n"),
            ("texte", "Lance la mise à jour prévue par Horizon Chantier à partir des documents enregistrés.\n"),
            ("liste", "✅ Enregistrer le PR et les fichiers concernés avant de cliquer.\n"),
            ("liste", "✅ Laisser Horizon lire les documents et effectuer la mise à jour prévue.\n"),

            ("section", "\nETAT D'AVANCEMENT\n"),
            ("section", "\n📊 État d'avancement\n"),
            ("texte", "Ouvre l’État d’avancement du chantier.\n"),
            ("liste", "✅ L’état d’avancement représente uniquement les travaux réellement exécutés.\n"),
            ("liste", "✅ Il contient les quantités réalisées, les montants exécutés et les montants réellement facturables.\n"),
            ("liste", "✅ Encoder les avancements et informations métier dans les zones prévues.\n"),
            ("liste", "✅ La synthèse reste une zone libre pour les ajustements métier prévus.\n"),
            ("liste", "✅ L’état d’avancement doit rester le reflet fidèle de la réalité exécutée du chantier.\n"),
            ("liste", "❌ Les avenants en moins ne font pas partie de l’état d’avancement.\n"),
            ("liste", "❌ Les corrections financières de pilotage ne font pas partie de l’état d’avancement.\n"),
            ("liste", "✅ Enregistrer avant Calcul état ou Clôturer état.\n"),

            ("section", "\n📊 Calcul état\n"),
            ("texte", "Demande à Horizon Chantier de mettre à jour l’état d’avancement enregistré.\n"),
            ("liste", "✅ À lancer après avoir enregistré l’état d’avancement dans Excel.\n"),
            ("liste", "❌ Ne pas modifier les cellules calculées pour forcer un résultat.\n"),

            ("section", "\n♻️ Clôturer état\n"),
            ("texte", "Clôture l’état d’avancement terminé pour préparer l’état suivant.\n"),
            ("liste", "✅ À utiliser uniquement lorsque la période est validée.\n"),
            ("danger", "⚠️⚠️ ATTENTION - OPÉRATION IRRÉVERSIBLE ⚠️⚠️\n"),
            ("danger", "La fonction Clôturer état réalise une mise à zéro de l’état d’avancement pour préparer l’état suivant.\n"),
            ("danger", "Avant de lancer la clôture, il est OBLIGATOIRE de créer une Sauvegarde PDF Historique du document.\n"),
            ("liste", "✅ Vérifier que l’état d’avancement est terminé et enregistré.\n"),
            ("liste", "✅ Créer une Sauvegarde PDF Historique du document.\n"),
            ("liste", "✅ Vérifier que le PDF a bien été créé dans le dossier Historique_Sauvegardes.\n"),
            ("liste", "✅ Seulement après cette vérification, lancer le bouton Clôturer état.\n"),
            ("danger", "Si la clôture est réalisée sans PDF Historique : les quantités du mois sont remises à zéro, l’état précédent n’est plus récupérable sous sa forme de travail, et il est impossible de revenir en arrière.\n"),
            ("danger", "RÈGLE ABSOLUE : aucune clôture ne doit être réalisée sans cette sauvegarde préalable.\n"),

            ("section", "\nPILOTAGE\n"),
            ("section", "\n🧭 Pilotage\n"),
            ("texte", "Affiche ou génère la synthèse financière globale du chantier.\n"),
            ("liste", "✅ Le pilotage représente la situation financière globale du chantier.\n"),
            ("liste", "✅ Le pilotage lit la synthèse de l’état et les avenants en moins dans leur document séparé.\n"),
            ("liste", "✅ Il prend en compte la soumission, les avenants en plus, les avenants en moins, les révisions, la production cumulée, le reste à facturer et l’avancement financier.\n"),
            ("liste", "✅ Les avenants sont gérés dans un document séparé.\n"),
            ("liste", "✅ Dans le modèle privé : avenants en plus en E, avenants en moins en G. Dans le modèle public Mariemont : en plus en D, en moins en E.\n"),
            ("liste", "✅ Les avenants servent au pilotage financier.\n"),
            ("liste", "✅ Les avenants en plus peuvent figurer dans la synthèse de l’état.\n"),
            ("liste", "❌ Les avenants en moins ne figurent jamais dans l’état d’avancement ni dans sa synthèse.\n"),
            ("liste", "❌ Un montant du fichier Avenants n'est pas une quantité exécutée : les travaux réalisés sont saisis dans les postes de l’état.\n"),
            ("liste", "❌ Le pilotage ne modifie jamais les fichiers Excel.\n"),
            ("liste", "✅ Il effectue uniquement une lecture des résultats.\n"),
            ("liste", "✅ À lancer après enregistrement et calcul des fichiers concernés.\n"),

            ("section", "\n📊 Pilotage délai\n"),
            ("texte", "Affiche ou génère la synthèse du délai chantier.\n"),
            ("liste", "✅ Vérifier que le fichier délai ou planning est à jour avant de cliquer.\n"),

            ("section", "\n⏱ Rendements\n"),
            ("texte", "Affiche ou génère la lecture des rendements du chantier.\n"),
            ("liste", "✅ Vérifier que les données de rendement sont enregistrées avant de cliquer.\n"),

            ("section", "\nRègles définitives sur les fichiers Excel\n"),
            ("liste", "❌ Ne pas modifier la structure des modèles Excel.\n"),
            ("liste", "❌ Ne pas modifier les formules.\n"),
            ("liste", "❌ Ne pas modifier les validations de données ou listes déroulantes.\n"),
            ("liste", "❌ Ne pas renommer les feuilles.\n"),
            ("liste", "✅ Les cellules de calcul doivent être préservées.\n"),
            ("liste", "✅ Les zones de saisie prévues par le logiciel sont les seules zones à modifier.\n"),
            ("liste", "✅ Respecter les zones de saisie prévues garantit une utilisation correcte du logiciel.\n"),

            ("section", "\nBon réflexe de fin de travail Excel\n"),
            ("ok", "✅ Enregistrer Excel.\n"),
            ("ok", "✅ Revenir dans Horizon Chantier.\n"),
            ("ok", "✅ Cliquer sur Sauvegarde PDF Historique.\n"),
            ("ok", "✅ Vérifier que le PDF daté est présent dans Historique_Sauvegardes.\n"),
            ("danger", "❌ Ne jamais attendre plusieurs jours avant de créer le PDF Historique d’un document important.\n"),

            ("section", "\nPhilosophie Horizon Chantier\n"),
            ("liste", "✅ Horizon privilégie une architecture simple et robuste.\n"),
            ("liste", "✅ Les calculs restent réalisés dans les documents Excel spécialisés.\n"),
            ("liste", "✅ Horizon lit les résultats et assure le pilotage.\n"),
            ("liste", "✅ Les corrections doivent rester ciblées.\n"),
            ("liste", "✅ Les utilisateurs ne doivent modifier que les zones prévues.\n"),
            ("fin", "\nFin du mode d’emploi.\n"),
        ]

        contenu_depannage_complet = [
            ("titre", "🔧 DÉPANNAGE\n"),

            ("section", "\n📖 CAHIER TECHNIQUE INTÉGRÉ AU PR\n"),
            ("texte", "Le lien sous la désignation doit afficher un onglet CSC dans le même classeur. Retour au PR, en haut de cet onglet, ramène au poste.\n"),
            ("liste", "✅ Si le lien est absent, vérifier que ce PR a été préparé avec les pages intégrées. L’équipement d’un chantier ne prépare pas automatiquement les autres.\n"),
            ("liste", "✅ Si un lecteur PDF, une recherche de fichier ou une demande d’accès apparaît, vérifier que vous utilisez le lien interne du PR équipé, et non un ancien lien vers un PDF ou le bouton Cahier des charges technique d’Horizon.\n"),
            ("liste", "✅ Si un lien ne fonctionne plus, vérifier que l’onglet CSC existe et n’a pas été renommé ou supprimé. Noter le numéro d’article pour faire corriger la correspondance.\n"),
            ("liste", "✅ Si les pages ne correspondent pas au poste, vérifier le numéro d’article et la version du cahier technique. Après modification du PDF source, les pages intégrées doivent être mises à jour : elles ne se rafraîchissent pas automatiquement.\n"),
            ("liste", "✅ Si le texte paraît trop petit, augmenter le zoom Excel. Les pages sont des images de consultation ; leur texte ne se modifie pas dans les cellules.\n"),
            ("liste", "✅ Avant toute correction, enregistrer et fermer le PR, puis conserver une copie complète du classeur avec les dernières saisies et photos.\n"),
            ("danger", "❌ Ne pas supprimer les photos, les onglets CSC ou les formules pour réparer un lien.\n"),

            ("section", "\n💰 PR (PRIX DE REVIENT) - FORMULES & DÉPANNAGE\n"),
            ("intro", "Cette aide rappelle les formules clés du PR et les points de contrôle à ne jamais casser.\n"),

            ("section", "\n📌 Règle importante\n"),
            ("liste", "✅ Si la cellule n’est pas gris clair → TU NE TOUCHES PAS.\n"),
            ("danger", "❌ Ne jamais modifier les formules.\n"),

            ("texte", "\n---\n"),

            ("section", "\n💰 MATIÈRE (P13)\n"),
            ("formule", "=SIERREUR(SOMME(P3:P11)/P12;0)\n"),

            ("texte", "\n---\n"),

            ("section", "\n💰 MAIN-D’ŒUVRE (P26)\n"),
            ("formule", "=SIERREUR(SOMME(P16:P24)/P25;0)\n"),

            ("texte", "\n---\n"),

            ("section", "\n📊 RÉCAPITULATIF\n"),
            ("formule", "C28 = P13\n"),
            ("formule", "C29 = P26\n"),
            ("formule", "C30 = C28 + C29\n"),

            ("texte", "\n---\n"),

            ("section", "\n💰 PRIX DE VENTE\n"),
            ("formule", "J28 = C30 * F28\n"),
            ("formule", "J29 = (C30 + J28) * F29\n"),
            ("formule", "F30 = C30 + J28 + J29\n"),
            ("formule", "F32 = F30 * (1 + F31)\n"),

            ("texte", "\n---\n"),

            ("section", "\n📋 LISTES DÉROULANTES\n"),
            ("danger", "⚠️ Si la flèche disparaît :\n"),
            ("liste", "✅ 1. Données\n"),
            ("liste", "✅ 2. Validation des données\n"),
            ("liste", "✅ 3. Autoriser : Liste\n"),
            ("texte", "\n📌 Source :\n"),
            ("formule", "Matière → Bibliothèque_Matière!A2:A500\n"),
            ("formule", "Main d’œuvre → Bibliothèque_MainOeuvre!A2:A500\n"),

            ("texte", "\n---\n"),

            ("section", "\n📌 IMPORTANT\n"),
            ("danger", "❌ Ne jamais modifier les formules.\n"),

            ("texte", "\n---\n"),

            ("section", "\n💰 AVENANTS EN PLUS\n"),
            ("liste", "✅ Modèle privé : avenants en plus en colonne E. Modèle public Mariemont : colonne D.\n"),
            ("liste", "✅ Vérifier que la saisie est faite dans la bonne ligne et la bonne colonne du modèle.\n"),
            ("danger", "❌ Ne pas déplacer les colonnes.\n"),
            ("danger", "❌ Ne pas modifier les formules liées aux avenants.\n"),

            ("texte", "\n---\n"),

            ("section", "\n💰 AVENANTS EN MOINS\n"),
            ("liste", "✅ Modèle privé : avenants en moins en colonne G. Modèle public Mariemont : colonne E.\n"),
            ("liste", "✅ Vérifier que la saisie est faite dans la bonne ligne et la bonne colonne du modèle.\n"),
            ("danger", "⚠️ Ne pas inverser les avenants en plus et les avenants en moins.\n"),
            ("danger", "❌ Ne pas modifier les formules pour corriger un résultat.\n"),

            ("texte", "\n---\n"),

            ("section", "\n💰 RÉVISION GLOBALE\n"),
            ("liste", "✅ La révision globale est reprise dans la colonne E.\n"),
            ("liste", "✅ La colonne E doit rester la référence pour la révision globale.\n"),
            ("danger", "❌ Ne pas déplacer la révision globale vers une autre colonne.\n"),
            ("danger", "❌ Ne pas modifier les formules liées à la révision globale.\n"),

            ("texte", "\n---\n"),

            ("section", "\n💰 RÉVISION DU MOIS MANUELLE\n"),
            ("liste", "✅ La révision du mois peut être saisie manuellement uniquement dans la zone prévue.\n"),
            ("danger", "❌ Ne pas écrire dans une cellule de calcul.\n"),
            ("danger", "❌ Ne pas forcer un total en remplaçant une formule.\n"),

            ("texte", "\n---\n"),

            ("section", "\n🧭 PILOTAGE\n"),
            ("liste", "✅ Le pilotage sert à contrôler et suivre les écarts.\n"),
            ("liste", "✅ Les documents Excel restent les sources de travail.\n"),
            ("liste", "✅ Horizon Chantier lit, contrôle et regroupe les informations disponibles.\n"),
            ("liste", "✅ En cas d’écart, vérifier d’abord les fichiers sources, les enregistrements Excel et les zones de saisie prévues.\n"),
            ("danger", "❌ Ne pas modifier les structures Excel pour obtenir un résultat attendu.\n"),

            ("texte", "\n---\n"),

            ("section", "\n🔍 VÉRIFIER PR\n"),
            ("liste", "✅ Enregistrer le PR dans Excel avant de cliquer sur Vérifier PR.\n"),
            ("liste", "✅ Si un message apparaît, lire le message et reprendre uniquement les zones de saisie prévues.\n"),
            ("liste", "✅ Si le fichier est verrouillé, fermer Excel puis relancer l’action.\n"),
            ("danger", "❌ Ne pas modifier une formule pour changer un résultat.\n"),

            ("texte", "\n---\n"),

            ("section", "\n⚙️ CALCUL / RECA\n"),
            ("liste", "✅ Vérifier que le bon chantier est sélectionné.\n"),
            ("liste", "✅ Enregistrer le PR et l’état d’avancement avant de cliquer.\n"),
            ("liste", "✅ Si l’action ne se termine pas, fermer Excel puis relancer une seule fois.\n"),
            ("danger", "⚠️ Si le problème continue, noter le message affiché et demander de l’aide.\n"),

            ("texte", "\n---\n"),

            ("section", "\n📊 ÉTAT D’AVANCEMENT\n"),
            ("liste", "✅ Si l’état ne s’ouvre pas, essayer d’abord Ce chantier.\n"),
            ("liste", "✅ Enregistrer Excel avant Calcul état ou Clôturer état.\n"),
            ("liste", "✅ La synthèse reste libre pour les ajustements métier prévus.\n"),
            ("danger", "❌ Ne pas écrire dans les cellules de calcul pour forcer un total.\n"),

            ("texte", "\n---\n"),

            ("section", "\n📊 CALCUL ÉTAT\n"),
            ("liste", "✅ Enregistrer l’état d’avancement dans Excel avant de cliquer.\n"),
            ("liste", "✅ Si un résultat paraît anormal, contrôler uniquement les zones de saisie prévues ou demander de l’aide.\n"),
            ("liste", "✅ Relancer une seule fois après avoir enregistré Excel.\n"),
            ("danger", "❌ Ne pas modifier les formules.\n"),

            ("texte", "\n---\n"),

            ("section", "\n📄 PDF HISTORIQUE\n"),
            ("liste", "✅ Vérifier qu’un fichier Excel est ouvert.\n"),
            ("liste", "✅ Si plusieurs classeurs sont ouverts, afficher le bon document Excel avant de revenir dans Horizon.\n"),
            ("liste", "✅ Enregistrer le fichier dans Excel avant de créer le PDF.\n"),
            ("liste", "✅ Vérifier que le PDF apparaît dans Historique_Sauvegardes du chantier.\n"),
            ("danger", "⚠️ Si Excel affiche une fenêtre, la lire et la fermer correctement avant de relancer la sauvegarde PDF.\n"),
            ("danger", "❌ Ne pas déplacer ou renommer le fichier Excel pendant la génération du PDF.\n"),
            ("texte", "\n📌 Le PDF Historique sert de trace avant calcul, clôture ou modification importante.\n"),

            ("texte", "\n---\n"),

            ("section", "\n📊 CLÔTURE ÉTAT\n"),
            ("danger", "⚠️ ATTENTION - OPÉRATION IRRÉVERSIBLE\n"),
            ("texte", "Clôturer état remet à zéro l’état d’avancement pour préparer l’état suivant.\n"),
            ("liste", "✅ Utiliser ce bouton uniquement quand la période est validée.\n"),
            ("liste", "✅ Avant toute clôture, vérifier que l’état d’avancement est terminé et enregistré.\n"),
            ("liste", "✅ Créer une Sauvegarde PDF Historique du document.\n"),
            ("liste", "✅ Vérifier que le PDF a bien été créé dans Historique_Sauvegardes.\n"),
            ("liste", "✅ Seulement ensuite, lancer Clôturer état.\n"),
            ("liste", "✅ Vérifier que le bon chantier et le bon état sont sélectionnés.\n"),
            ("danger", "\n⚠️ Sans PDF Historique, les quantités du mois sont remises à zéro et l’état précédent n’est plus récupérable sous sa forme de travail.\n"),
            ("danger", "❌ Aucune clôture ne doit être réalisée sans cette sauvegarde préalable.\n"),
            ("danger", "❌ Ne pas clôturer pour tester.\n"),

            ("texte", "\n---\n"),

            ("section", "\n🔗 LIENS EXTERNES EXCEL\n"),
            ("danger", "⚠️ Si Excel signale des liens externes, ne pas supprimer les liens sans consigne.\n"),
            ("liste", "✅ Lire le message Excel avant de répondre.\n"),
            ("liste", "✅ Si le fichier provient d’un modèle ou d’un ancien chantier, noter le message affiché.\n"),
            ("danger", "❌ Ne pas casser les liaisons pour ouvrir plus vite le fichier.\n"),

            ("texte", "\n---\n"),

            ("section", "\n🛠 RÉPARATION XML EXCEL\n"),
            ("danger", "⚠️ Si Excel propose une réparation XML, ne pas continuer sans prudence.\n"),
            ("liste", "✅ Enregistrer une copie avant toute action risquée si Excel le permet.\n"),
            ("liste", "✅ Noter le message exact affiché par Excel.\n"),
            ("danger", "⚠️ Si Excel supprime des éléments pendant la réparation, demander de l’aide avant de poursuivre le travail.\n"),

            ("texte", "\n---\n"),

            ("section", "\n♻️ RÉCUPÉRATION AUTOMATIQUE EXCEL\n"),
            ("liste", "✅ Si Excel propose une récupération automatique, vérifier le nom du fichier récupéré.\n"),
            ("danger", "❌ Ne pas écraser le fichier de travail sans certitude.\n"),
            ("liste", "✅ Comparer le fichier récupéré avec le fichier attendu.\n"),
            ("liste", "✅ Enregistrer correctement le fichier seulement après vérification.\n"),

            ("texte", "\n---\n"),

            ("section", "\n📂 FICHIER DÉJÀ OUVERT\n"),
            ("liste", "✅ Si Excel indique que le fichier est déjà ouvert, vérifier les fenêtres Excel ouvertes.\n"),
            ("liste", "✅ Fermer le fichier déjà ouvert si nécessaire.\n"),
            ("danger", "❌ Ne pas ouvrir une deuxième version en lecture seule pour travailler dessus.\n"),
            ("danger", "⚠️ Si le fichier est utilisé par une autre personne ou une autre session, attendre ou demander confirmation.\n"),

            ("texte", "\n---\n"),

            ("section", "\n📂 DEUX FICHIERS PORTANT LE MÊME NOM\n"),
            ("liste", "✅ Vérifier leur emplacement complet.\n"),
            ("danger", "❌ Ne pas se baser uniquement sur le nom affiché dans Excel.\n"),
            ("liste", "✅ Contrôler le dossier du chantier avant de travailler.\n"),
            ("danger", "❌ Ne pas enregistrer dans le mauvais dossier.\n"),

            ("texte", "\n---\n"),

            ("section", "\n🔒 CELLULES PROTÉGÉES OU FORMULES INVISIBLES\n"),
            ("liste", "✅ Les cellules contenant des formules sont protégées pour éviter les erreurs de manipulation.\n"),
            ("liste", "✅ Les formules peuvent être masquées dans la barre de formule lorsque la feuille est protégée.\n"),
            ("liste", "✅ Les zones de saisie prévues restent modifiables.\n"),
            ("liste", "✅ La synthèse de l’état d’avancement reste libre pour les ajustements métier prévus.\n"),
            ("danger", "❌ Ne pas enlever les protections pour travailler plus vite.\n"),

            ("texte", "\n---\n"),

            ("section", "\n⚠️ ERREURS COURANTES\n"),
            ("danger", "❌ Travailler sur le mauvais chantier sélectionné.\n"),
            ("danger", "❌ Oublier d’enregistrer Excel avant un calcul ou un PDF Historique.\n"),
            ("danger", "❌ Ouvrir plusieurs fichiers Excel et laisser le mauvais classeur actif avant la sauvegarde PDF.\n"),
            ("danger", "❌ Renommer un fichier utilisé par Horizon.\n"),
            ("danger", "❌ Copier-coller sur des cellules de calcul.\n"),
            ("danger", "❌ Modifier le dossier Modèles pour intervenir sur un chantier existant.\n"),

            ("texte", "\n---\n"),

            ("section", "\n🧭 PROCÉDURE GÉNÉRALE EN CAS DE BLOCAGE\n"),
            ("liste", "✅ 1. Arrêter les modifications hasardeuses.\n"),
            ("liste", "✅ 2. Identifier le chantier sélectionné et le fichier concerné.\n"),
            ("liste", "✅ 3. Enregistrer Excel si possible.\n"),
            ("liste", "✅ 4. Créer une Sauvegarde PDF Historique si le document est lisible et important.\n"),
            ("liste", "✅ 5. Fermer les boîtes de dialogue Excel puis relancer une seule fois l’action.\n"),
            ("liste", "✅ 6. Si le problème persiste, noter le message affiché, le bouton utilisé et le nom du chantier.\n"),
            ("danger", "❌ Ne jamais modifier une formule, une validation, une feuille ou une structure pour dépanner dans l’urgence.\n"),

            ("texte", "\n---\n"),

            ("fin", "\nFin du dépannage.\n"),
        ]

        def afficher_aide(*_):
            texte.configure(state="normal")
            texte.delete("1.0", "end")

            if aide_var.get() == "🔧 Dépannage":
                contenu = contenu_depannage_complet
            else:
                contenu = contenu_mode_emploi

            for tag, bloc in contenu:
                texte.insert("end", bloc, tag)

            texte.configure(state="disabled")
            texte.yview_moveto(0)

        choix_aide.bind("<<ComboboxSelected>>", afficher_aide)
        afficher_aide()

        footer = ttk.Frame(win, padding=(14, 0, 14, 14))
        footer.pack(fill="x")
        ttk.Checkbutton(
            footer,
            text="Toujours visible",
            variable=topmost_var,
            command=appliquer_toujours_visible,
        ).pack(side="left")
        ttk.Button(footer, text="Fermer", command=fermer).pack(side="right")

    def ouvrir_aide_mode_emploi(self) -> None:
        self.ouvrir_aide()

    def calcul_etat_avancement(self):
        if not self._chemin_chantier_selectionne():
            messagebox.showwarning("Calcul état", "Sélectionne un chantier.")
            return False
        wb = valeurs = None
        try:
            chemin = _selectionner_etat(self._dossier_chantier_selectionne())
            signature = _signature_excel(chemin)
            wb = load_workbook(chemin, keep_vba=Path(chemin).suffix.lower() == ".xlsm")
            valeurs = load_workbook(chemin, data_only=True)
            _recalculer_etat(_feuille_etat(wb), _feuille_etat(valeurs))
            _protect_formula_cells(wb, chemin)
            _sauver_excel_atomique(wb, chemin, signature)
            print("Etat recalculé")
            messagebox.showinfo("Calcul état", "État enregistré. Pour actualiser les formules de synthèse, "
                                "ouvrez puis enregistrez l'état dans Excel avant le pilotage.")
            return True
        except Exception as e:
            messagebox.showerror("Calcul état", str(e))
            return False
        finally:
            if wb is not None:
                _fermer_classeur(wb)
            if valeurs is not None:
                _fermer_classeur(valeurs)

    def mise_a_zero_etat(self):
        if not self._confirmer_cloture_etat():
            return
        if not self._chemin_chantier_selectionne():
            messagebox.showwarning("Clôturer état", "Sélectionne un chantier dans la liste.")
            return
        wb = valeurs = None
        try:
            chemin = _selectionner_etat(self._dossier_chantier_selectionne())
            signature = _signature_excel(chemin)
            wb = load_workbook(chemin, keep_vba=Path(chemin).suffix.lower() == ".xlsm")
            valeurs = load_workbook(chemin, data_only=True)
            ws = _feuille_etat(wb)
            # Valider et recalculer avant le report, puis publier en une seule écriture.
            _recalculer_etat(ws, _feuille_etat(valeurs))
            _reporter_quantites_cloture(ws)
            _recalculer_etat(ws, _feuille_etat(valeurs))
            _protect_formula_cells(wb, chemin)
            sauvegarde = chemin.parent / "Sauvegarde"
            sauvegarde.mkdir(parents=True, exist_ok=True)
            copie = sauvegarde / f"{chemin.stem}_{datetime.now():%Y-%m-%d_%H-%M-%S_%f}{chemin.suffix}"
            shutil.copy2(chemin, copie)
            _sauver_excel_atomique(wb, chemin, signature)
            print("Clôture état effectuée")
            messagebox.showinfo("Clôturer état", "Clôture enregistrée et sauvegarde conservée. "
                                "Ouvrez puis enregistrez l'état dans Excel pour actualiser les synthèses avant le pilotage.")
        except Exception as e:
            messagebox.showerror("Clôturer état", f"La clôture n'a pas été enregistrée.\n{e}")
        finally:
            if wb is not None:
                _fermer_classeur(wb)
            if valeurs is not None:
                _fermer_classeur(valeurs)

    def verifier_pr(self):
        p = self._chemin_chantier_selectionne()
        if not p:
            messagebox.showwarning("PR", "Sélectionne un chantier dans la liste.")
            return

        dossier_chantier = self._dossier_chantier_selectionne()
        chemin = _normaliser_dossier_data(dossier_chantier) / "prix_de_revient.xlsx"

        import os
        from datetime import datetime

        if not chemin.exists():
            messagebox.showwarning(
                "Vérifier PR",
                "Le fichier prix_de_revient.xlsx est introuvable."
            )
            return

        timestamp = os.path.getmtime(chemin)
        date_modif = datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y %H:%M:%S")

        messagebox.showinfo(
            "Date du PR",
            f"Dernier enregistrement du PR :\n{date_modif}\n\n"
            "⚠️ Si tu viens de modifier le PR :\n"
            "➡️ enregistre dans Excel (Ctrl+S sur Windows, Cmd+S sur Mac)\n"
            "➡️ puis clique sur Vérifier PR"
        )
        
        wb = None
        wb_formules = None
        try:
            signature = _signature_excel(chemin)
            wb = _charger_pr_controle(chemin)
            wb_formules = load_workbook(chemin, read_only=True, data_only=False)
            nb_blocs, erreurs = _verifier_blocs_pr(wb["Chiffrage"], wb_formules["Chiffrage"])
            if _signature_excel(chemin) != signature:
                raise ValueError("Le PR a changé pendant le contrôle. Relancez la vérification.")
            if erreurs:
                messagebox.showwarning("PR", f"{len(erreurs)} écart(s) aux contrôles matière/MO sur {nb_blocs} bloc(s).\n"
                                       "Les règles particulières des postes restent à vérifier ; aucune formule n'est modifiée.\n\n"
                                       + "\n".join(erreurs[:40])
                                       + ("\n… Autres écarts non affichés." if len(erreurs) > 40 else ""))
            else:
                messagebox.showinfo("PR", f"Aucun écart dans les formules matière et main-d'œuvre reconnues "
                                    f"sur {nb_blocs} bloc(s). Les saisies et formules particulières non reconnues "
                                    "restent sous contrôle métier.")
        except Exception as e:
            messagebox.showerror("Vérifier PR", str(e))
        finally:
            if wb is not None:
                wb.close()
            if wb_formules is not None:
                wb_formules.close()

    def pilotage_chantier(self):
        import os
        import subprocess
        import tempfile
        from datetime import datetime
        import matplotlib.pyplot as plt
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.enums import TA_RIGHT, TA_CENTER

        nom_chantier = self._nom_chantier_selectionne()
        if not nom_chantier:
            return

        try:
            chemin = _selectionner_etat(self._dossier_chantier_selectionne()).as_posix()
        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("Pilotage", str(e))
            return

        try:
            # S contient les pourcentages des postes, inutilisés par le pilotage.
            # openpyxl efface leur cache lors du recalcul, même si les montants de synthèse restent lisibles.
            wb = _charger_valeurs_excel(chemin, "Bordereau", "État d'avancement", max_col=23,
                                        lecture_seule=False, colonnes_formules_facultatives=("S",))
        except Exception as e:
            messagebox.showerror("Pilotage", str(e))
            return
        ws = _feuille_etat(wb)
        

        def nombre(val):
            if val in (None, "", "-", "—"):
                return 0
            try:
                texte_nombre = str(val)
                texte_nombre = texte_nombre.replace("€", "")
                texte_nombre = texte_nombre.replace("−", "-")
                texte_nombre = texte_nombre.replace(" ", "")
                texte_nombre = texte_nombre.replace("\xa0", "")
                texte_nombre = texte_nombre.replace("\u202f", "")
                texte_nombre = texte_nombre.replace(",", ".")
                return float(texte_nombre)
            except:
                return 0

        def texte(val):
            return str(val or "").strip().lower()

        def trouver_ligne_libelle(mots):
            for row in range(ws.max_row, 1, -1):
                for col in range(1, ws.max_column + 1):
                    val = texte(ws.cell(row=row, column=col).value)
                    if not val:
                        continue
                    if all(mot in val for mot in mots):
                        return row, col
            return None, None

        def premiere_valeur_a_droite(row, col_depart, max_ecart=3):
            return _montant_a_droite(ws, row, col_depart, nombre, max_ecart)

        def valeur_soumission_publique(libelles):
            for libelle in libelles:
                row, col = trouver_ligne_libelle([libelle])
                if row and col:
                    return premiere_valeur_a_droite(row, col), row, col
            return 0, None, None

        montant_soumission, ligne_soumission, col_soumission = valeur_soumission_publique(
            ["total soumission hors tva"]
        )
        if ligne_soumission is None:
            montant_sauvegarde = _soumission_avant_cloture(chemin, _nombre_metier)
            if montant_sauvegarde is not None:
                montant_soumission = montant_sauvegarde
            else:
                wb.close()
                messagebox.showerror("Pilotage", "Montant de soumission introuvable dans l'état et ses sauvegardes.")
                return
        _total_etat_cumule, _ligne_etat_cumule, _col_etat_cumule = valeur_soumission_publique(
            ["total état cumulé hors tva", "total etat cumulé hors tva"]
        )
        total_mois, ligne_mois, col_mois = valeur_soumission_publique(
            ["total du mois hors tva", "total du mois hors tva"]
        )
        avenants_plus, _ligne_avenants, _col_avenants = valeur_soumission_publique(
            ["total des avenants cumulé", "total des avenants cumule"]
        )
        total_realise, ligne_execute, col_execute = valeur_soumission_publique(
            ["total exécuté", "total execute"]
        )
        if _ligne_etat_cumule is not None:
            total_realise = _total_etat_cumule
        revision, _ligne_revision, _col_revision = valeur_soumission_publique(
            ["montant de la révision", "montant de la revision"]
        )
        revision_globale, _, _ = valeur_soumission_publique(
            ["total des révision cumulé", "total des revision cumule"]
        )
        _montant_global, _ligne_global, _col_global = valeur_soumission_publique(
            ["montant global à facturer", "montant global a facturer"]
        )

        # Ces trois valeurs pilotent tous les indicateurs : une absence ne vaut pas zéro.
        try:
            if ligne_soumission is not None:
                montant_soumission = _montant_a_droite(ws, ligne_soumission, col_soumission,
                                                       nombre, requis=True)
            ligne_cumule = _ligne_etat_cumule if _ligne_etat_cumule is not None else ligne_execute
            col_cumule = _col_etat_cumule if _ligne_etat_cumule is not None else col_execute
            total_realise = _montant_a_droite(ws, ligne_cumule, col_cumule, nombre, requis=True)
            total_mois = _montant_a_droite(ws, ligne_mois, col_mois, nombre, requis=True)
        except ValueError as erreur:
            wb.close()
            messagebox.showerror("Pilotage", str(erreur))
            return

        if _est_metre_public(ws) and ws["P9"].value == "Numéro de l’état":
            avenants_plus, _, _ = valeur_soumission_publique(["avenants cumulés en plus"])
            try:
                avenants_moins = _lire_avenants_moins_publics(chemin)
            except (FileNotFoundError, ValueError) as erreur:
                wb.close()
                messagebox.showerror("Pilotage", str(erreur))
                return
        elif "Avenants" in wb.sheetnames:
            ws_avenants = wb["Avenants"]
            _, avenants_moins = _lire_synthese_avenants_pilotage(
                ws_avenants,
                nombre,
            )
        else:
            avenants_moins = 0
            chemin_avenants = Path(chemin).with_name("Avenants.xlsx")
            if chemin_avenants.exists():
                wb_avenants = _charger_valeurs_excel(chemin_avenants, None, "Avenants", max_col=7, lecture_seule=False)
                ws_avenants = wb_avenants["Avenants"] if "Avenants" in wb_avenants.sheetnames else wb_avenants.active
                _, avenants_moins = _lire_synthese_avenants_pilotage(
                    ws_avenants,
                    nombre,
                )
                wb_avenants.close()
        total_marche = montant_soumission + avenants_plus - avenants_moins

        print("avenants_plus =", avenants_plus)
        print("avenants_moins =", avenants_moins)

        print("ligne_mois =", ligne_mois, "col_mois =", col_mois)
        print("total_mois =", total_mois)
       

    
        date_du_jour = datetime.now().strftime("%d/%m/%Y")
        realise_mois = total_mois
        realise_cumule = total_realise
        production_mois = realise_mois + revision
        production_cumulee = realise_cumule + revision_globale
        reste_a_facturer = total_marche - production_cumulee
        avancement = 0 if total_marche == 0 else (production_cumulee / total_marche) * 100
        libelle_ecart_marche, montant_ecart_marche, ecart_marche_texte, ecart_initial_texte = _presentation_ecarts_pilotage(
            montant_soumission, total_marche, production_cumulee)
        montant_ecart_affiche = (f"{montant_ecart_marche:+,.2f} €" if reste_a_facturer < 0
                                else f"{montant_ecart_marche:,.2f} €")

        texte_popup = f"""PILOTAGE CHANTIER

       
    Chantier : {nom_chantier}
    Date     : {date_du_jour}

    Montant soumission              : {montant_soumission:,.2f} €
    Avenants cumulés en plus        : {avenants_plus:,.2f} €
    Avenants cumulés en moins       : {avenants_moins:,.2f} €
    Marché total                    : {total_marche:,.2f} €
    Réalisé du mois                 : {realise_mois:,.2f} €
    Révision renseignée             : {revision:,.2f} €
    Production du mois              : {production_mois:,.2f} €
    Réalisé cumulé                  : {realise_cumule:,.2f} €
    Production cumulée              : {production_cumulee:,.2f} €
    {libelle_ecart_marche} : {montant_ecart_affiche}
    Écart au marché total           : {ecart_marche_texte}
    Écart à la soumission initiale  : {ecart_initial_texte}
    Avancement                      : {avancement:.1f} %
    """

        pdf_path = str(Path(chemin).with_name("pilotage_chantier.pdf"))
        wb.close()

        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=landscape(A4),
            leftMargin=8 * mm,
            rightMargin=8 * mm,
            topMargin=5 * mm,
            bottomMargin=5 * mm,
        )

        elements = []
        graphique_tmp = None
        logo_path = str(dossier_base() / "logo_jt_bati.png")

        style_titre = ParagraphStyle(
            "PilotageTitre",
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=18,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#163A59"),
        )
        style_meta_label = ParagraphStyle(
            "PilotageMetaLabel",
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=9,
            textColor=colors.HexColor("#36556F"),
        )
        style_meta_valeur = ParagraphStyle(
            "PilotageMetaValeur",
            fontName="Helvetica",
            fontSize=8.8,
            leading=9.2,
            textColor=colors.black,
        )
        style_bloc_titre = ParagraphStyle(
            "PilotageBlocTitre",
            fontName="Helvetica-Bold",
            fontSize=9.6,
            leading=10,
            alignment=TA_CENTER,
            textColor=colors.white,
        )
        style_libelle = ParagraphStyle(
            "PilotageLibelle",
            fontName="Helvetica",
            fontSize=8.6,
            leading=9,
            textColor=colors.black,
        )
        style_valeur = ParagraphStyle(
            "PilotageValeur",
            fontName="Helvetica-Bold",
            fontSize=8.8,
            leading=9,
            alignment=TA_RIGHT,
            textColor=colors.black,
        )
        style_carte_label = ParagraphStyle(
            "PilotageCarteLabel",
            fontName="Helvetica-Bold",
            fontSize=7.6,
            leading=8,
            textColor=colors.HexColor("#5A7086"),
        )
        style_carte_valeur = ParagraphStyle(
            "PilotageCarteValeur",
            fontName="Helvetica-Bold",
            fontSize=11.2,
            leading=11.4,
            alignment=TA_RIGHT,
            textColor=colors.HexColor("#173C5A"),
        )
        style_carte_valeur_finale = ParagraphStyle(
            "PilotageCarteValeurFinale",
            fontName="Helvetica-Bold",
            fontSize=12.2,
            leading=12.2,
            alignment=TA_RIGHT,
            textColor=colors.HexColor("#C0504D" if reste_a_facturer < 0 else "#2E7D32"),
        )
        style_bandeau = ParagraphStyle(
            "PilotageBandeau",
            fontName="Helvetica-Bold",
            fontSize=9.3,
            leading=9.8,
            textColor=colors.HexColor("#173C5A"),
        )
        style_avancement_titre = ParagraphStyle(
            "PilotageAvancementTitre",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=8.2,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#5A7086"),
        )
        style_avancement_valeur = ParagraphStyle(
            "PilotageAvancementValeur",
            fontName="Helvetica-Bold",
            fontSize=10.5,
            leading=10.8,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#163A59"),
        )

        logo_cell = Spacer(1, 1)
        if os.path.exists(logo_path):
            try:
                logo_cell = Image(logo_path, width=27 * mm, height=10.8 * mm)
                logo_cell.hAlign = "LEFT"
            except:
                logo_cell = Spacer(1, 1)

        meta_table = Table([
            [
                Paragraph("CHANTIER", style_meta_label),
                Paragraph(nom_chantier, style_meta_valeur),
                Paragraph("DATE", style_meta_label),
                Paragraph(date_du_jour, style_meta_valeur),
            ]
        ], colWidths=[18 * mm, 47 * mm, 12 * mm, 22 * mm])
        meta_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))

        header_table = Table([
            [logo_cell, Paragraph("PILOTAGE CHANTIER", style_titre), meta_table]
        ], colWidths=[34 * mm, 158 * mm, 74 * mm])
        header_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F8FB")),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#D9E4EE")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ("ALIGN", (2, 0), (2, 0), "RIGHT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(header_table)
        elements.append(Spacer(1, 3))

        valeurs_graphique = [total_marche, production_cumulee, montant_ecart_marche]
        valeurs_plot = [max(total_marche, 0), max(production_cumulee, 0), montant_ecart_marche]
        etiquettes_graphique = ["Marché\ntotal", "Production\ncumulée",
                                 "Au-delà du\nmarché" if reste_a_facturer < 0 else "Reste à\nfacturer"]
        couleurs_graphique = ["#1F4E79", "#E69138", "#C0504D" if reste_a_facturer < 0 else "#70AD47"]

        fig, ax = plt.subplots(figsize=(3.1, 3.7), facecolor="white")
        barres = ax.bar(
            range(3),
            valeurs_plot,
            width=0.55,
            color=couleurs_graphique,
            edgecolor="#FFFFFF",
            linewidth=0.8,
            zorder=3,
        )

        max_valeur = max([abs(v) for v in valeurs_plot] + [1])
        marge = max_valeur * 0.18
        ax.set_ylim(0, max(max(valeurs_plot), 1) + marge)
        ax.set_xticks(range(3), etiquettes_graphique)
        ax.tick_params(axis="x", labelsize=7.8, colors="#173C5A", length=0, pad=6)
        ax.tick_params(axis="y", labelsize=7.5, colors="#6A7E90")
        ax.grid(axis="y", color="#D9E4EE", linewidth=0.8, alpha=0.8, zorder=0)
        ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#D9E4EE")
        ax.spines["bottom"].set_color("#D9E4EE")

        for index, (barre, valeur) in enumerate(zip(barres, valeurs_graphique)):
            ax.text(
                barre.get_x() + barre.get_width() / 2,
                barre.get_height() + max(marge * 0.08, 0.6),
                montant_ecart_affiche if index == 2 else f"{valeur:,.2f} €",
                ha="center",
                va="bottom",
                fontsize=7.8,
                fontweight="bold",
                color="#173C5A",
            )

        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            graphique_tmp = tmp.name

        plt.tight_layout()
        plt.savefig(graphique_tmp, dpi=220, bbox_inches="tight", facecolor="white")
        plt.close(fig)

        graphique = Image(graphique_tmp, width=66 * mm, height=66 * mm)
        graphique.hAlign = "CENTER"

        largeur_blocs_gauche = [95 * mm, 30 * mm]
        largeur_blocs_droite = [38 * mm, 24 * mm]

        bloc_1 = Table([
            ["CONSTRUCTION", ""],
            [Paragraph("Montant soumission", style_libelle), Paragraph(f"{montant_soumission:,.2f} €", style_valeur)],
            [Paragraph("+ Avenants cumulés en plus", style_libelle), Paragraph(f"{avenants_plus:,.2f} €", style_valeur)],
            [Paragraph("- Avenants cumulés en moins", style_libelle), Paragraph(f"{avenants_moins:,.2f} €", style_valeur)],
            [Paragraph("= Marché total", style_libelle), Paragraph(f"{total_marche:,.2f} €", style_valeur)],
        ], colWidths=largeur_blocs_gauche)
        bloc_1.hAlign = "CENTER"
        bloc_1.setStyle(TableStyle([
            ("SPAN", (0, 0), (1, 0)),
            ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#1F4E79")),
            ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
            ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (1, 0), 12),
            ("BACKGROUND", (0, 1), (1, -1), colors.white),
            ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#F8FBFF")]),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C5D3")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ("BACKGROUND", (0, 4), (1, 4), colors.HexColor("#FFF2CC")),
            ("FONTNAME", (0, 4), (1, 4), "Helvetica-Bold"),
        ]))

        bloc_2 = Table([
            ["RÉALISÉ", ""],
            [Paragraph("Réalisé du mois", style_libelle), Paragraph(f"{realise_mois:,.2f} €", style_valeur)],
            [Paragraph("+ Révision renseignée", style_libelle), Paragraph(f"{revision:,.2f} €", style_valeur)],
            [Paragraph("= Production du mois", style_libelle), Paragraph(f"{production_mois:,.2f} €", style_valeur)],
            [Paragraph("Réalisé cumulé", style_libelle), Paragraph(f"{realise_cumule:,.2f} €", style_valeur)],
            [Paragraph("= Production cumulée", style_libelle), Paragraph(f"{production_cumulee:,.2f} €", style_valeur)],
        ], colWidths=largeur_blocs_gauche)
        bloc_2.hAlign = "CENTER"
        bloc_2.setStyle(TableStyle([
            ("SPAN", (0, 0), (1, 0)),
            ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#4F81BD")),
            ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
            ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (1, 0), 12),
            ("BACKGROUND", (0, 1), (1, -1), colors.white),
            ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#F7FAFD")]),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C5D3")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ("BACKGROUND", (0, 3), (1, 3), colors.HexColor("#FFF2CC")),
            ("FONTNAME", (0, 3), (1, 3), "Helvetica-Bold"),
            ("BACKGROUND", (0, 5), (1, 5), colors.HexColor("#FFF2CC")),
            ("FONTNAME", (0, 5), (1, 5), "Helvetica-Bold"),
        ]))

        bloc_3 = Table([
            ["RÉSULTAT", ""],
            [Paragraph(libelle_ecart_marche, style_libelle), Paragraph(montant_ecart_affiche, style_valeur)],
            [Paragraph("Avancement", style_libelle), Paragraph(f"{avancement:.1f} %", style_valeur)],
            [Paragraph("Écart financier final sur le marché total", style_libelle), Paragraph(ecart_marche_texte, style_valeur)],
            [Paragraph("Écart à la soumission initiale", style_libelle), Paragraph(ecart_initial_texte, style_valeur)],
        ], colWidths=largeur_blocs_gauche)
        bloc_3.hAlign = "CENTER"
        bloc_3.setStyle(TableStyle([
            ("SPAN", (0, 0), (1, 0)),
            ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#C0504D")),
            ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
            ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (1, 0), 12),
            ("BACKGROUND", (0, 1), (1, -1), colors.white),
            ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#FCF7F7")]),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#C9B1B1")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ("BACKGROUND", (0, 1), (1, 1), colors.HexColor("#FCE8E6" if reste_a_facturer < 0 else "#D9EAD3")),
            ("FONTNAME", (0, 1), (1, 1), "Helvetica-Bold"),
            ("BACKGROUND", (0, 3), (1, 3), colors.HexColor("#FCE8E6" if reste_a_facturer < 0 else "#D9EAD3")),
            ("FONTNAME", (0, 3), (1, 3), "Helvetica-Bold"),
        ]))

        colonne_gauche = Table([
            [bloc_1],
            [bloc_2],
            [bloc_3],
        ], colWidths=[125 * mm])
        colonne_gauche.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))

        bloc_graphique = Table([
            [Paragraph("RÉSUMÉ VISUEL", style_bloc_titre)],
            [graphique],
        ], colWidths=[70 * mm])
        bloc_graphique.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#173C5A")),
            ("BACKGROUND", (0, 1), (0, 1), colors.white),
            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B7C5D3")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]))

        indicateurs_table = Table([
            [Paragraph("Marché total", style_carte_label), Paragraph(f"{total_marche:,.2f} €", style_carte_valeur)],
            [Paragraph("Production du mois", style_carte_label), Paragraph(f"{production_mois:,.2f} €", style_carte_valeur)],
            [Paragraph("Production cumulée", style_carte_label), Paragraph(f"{production_cumulee:,.2f} €", style_carte_valeur)],
            [Paragraph("Avancement %", style_carte_label), Paragraph(f"{avancement:.1f} %", style_carte_valeur)],
            [Paragraph(libelle_ecart_marche, style_carte_label), Paragraph(montant_ecart_affiche, style_carte_valeur_finale)],
            [Paragraph("Écart financier final sur le marché total", style_carte_label), Paragraph(ecart_marche_texte, style_carte_valeur_finale)],
            [Paragraph("Écart à la soumission initiale", style_carte_label), Paragraph(ecart_initial_texte, style_carte_valeur_finale)],
        ], colWidths=largeur_blocs_droite)
        indicateurs_table.setStyle(TableStyle([
            ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#F6F9FC")]),
            ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#D6E0EA")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ("BACKGROUND", (0, 4), (1, 4), colors.HexColor("#FCE8E6" if reste_a_facturer < 0 else "#D9EAD3")),
            ("BACKGROUND", (0, 5), (1, 5), colors.HexColor("#FCE8E6" if reste_a_facturer < 0 else "#D9EAD3")),
            ("BACKGROUND", (0, 6), (1, 6), colors.HexColor("#D9EAD3")),
        ]))

        indicateurs_cles = Table([
            [Paragraph("INDICATEURS CLÉS", style_bloc_titre)],
            [indicateurs_table],
        ], colWidths=[70 * mm])
        indicateurs_cles.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#365F91")),
            ("BACKGROUND", (0, 1), (0, 1), colors.white),
            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B7C5D3")),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]))

        colonne_droite = Table([
            [bloc_graphique],
            [indicateurs_cles],
        ], colWidths=[70 * mm])
        colonne_droite.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))

        corps = Table([
            [colonne_gauche, colonne_droite]
        ], colWidths=[128 * mm, 70 * mm])
        corps.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        elements.append(corps)
        elements.append(Spacer(1, 2))

        footer_row_height = 14.5 * mm
        footer_split_height = 5.2 * mm

        avancement_table = Table([
            [Paragraph("AVANCEMENT", style_avancement_titre)],
            [Paragraph(f"{avancement:.1f} %", style_avancement_valeur)],
        ], colWidths=[30 * mm], rowHeights=[footer_split_height, footer_row_height - footer_split_height])
        avancement_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EAF1F8")),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#B7C5D3")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))

        footer_explication = Table([
            [
                Paragraph(
                    f"{libelle_ecart_marche} = écart entre marché total et production cumulée<br/>"
                    f"{total_marche:,.2f} € et {production_cumulee:,.2f} € : {ecart_marche_texte}",
                    style_bandeau,
                )
            ]
        ], colWidths=[168 * mm], rowHeights=[footer_row_height])
        footer_explication.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#FFF2CC")),
            ("BOX", (0, 0), (0, 0), 0.7, colors.HexColor("#D6B656")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))

        footer_table = Table([
            [footer_explication, avancement_table]
        ], colWidths=[168 * mm, 30 * mm], rowHeights=[footer_row_height])
        footer_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        elements.append(footer_table)

        _publier_pdf_atomique(doc, elements, pdf_path)

        if graphique_tmp and os.path.exists(graphique_tmp):
            try:
                os.unlink(graphique_tmp)
            except:
                pass

        ouvrir_chemin(pdf_path)

        
        messagebox.showinfo("Pilotage chantier", texte_popup)
   
   
    def pilotage_delai(self):
        import os
        import subprocess
        import tempfile
        from datetime import datetime
        import matplotlib.pyplot as plt
        from openpyxl import load_workbook
        from tkinter import messagebox
        from reportlab.platypus import SimpleDocTemplate, Spacer, Table, TableStyle, Image, Paragraph
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.enums import TA_RIGHT, TA_CENTER

        def nombre(valeur):
            try:
                if valeur is None:
                    return 0
                if isinstance(valeur, str):
                    valeur = valeur.replace("%", "").replace("€", "").replace(" ", "").replace(",", ".").strip()
                    if valeur in ("", "-", "—"):
                        return 0
                return float(valeur)
            except:
                return 0

        try:
            nom_chantier = self._nom_chantier_selectionne()
            if not nom_chantier:
                return

            dossier_chantier = self._dossier_chantier_selectionne()
            fichier_delai = os.path.join(dossier_chantier, "Delai.xlsx")

            if not os.path.exists(fichier_delai):
                messagebox.showerror("Erreur", f"Fichier introuvable :\n{fichier_delai}")
                return

            wb = _charger_valeurs_excel(fichier_delai, None, Path(fichier_delai).stem, max_col=15, lecture_seule=False)
            ws = wb.active

            def lire_valeur_delai(libelles):
                if isinstance(libelles, str):
                    libelles = [libelles]
                for row in range(1, ws.max_row + 1):
                    libelle = str(ws[f"M{row}"].value or "").strip().lower()
                    for recherche in libelles:
                        if recherche in libelle:
                            return nombre(ws[f"O{row}"].value)
                return 0

            # Bloc haut : délai corrigé
            delai_soumission_jour = lire_valeur_delai("délai soumission jour")
            jours_avenant_plus = lire_valeur_delai("jour accordes avenant en plus")
            jours_attente_technique = lire_valeur_delai("jour accordes dû à une attente technique")
            jours_imprevu = lire_valeur_delai("jour accordes dû à imprévu, prévisible non visible")
            delai_corrige_jour = lire_valeur_delai("délai corrigé en jour")

            # Bloc bas : délai consommé
            jours_consommes = lire_valeur_delai(["jours actif consommé", "jours consommé"])
            jours_avenant_moins = lire_valeur_delai("jour accordes avenant en moins")
            jours_intemperies = lire_valeur_delai("jour intempéries")
            conges = lire_valeur_delai(["congé/ferrier/compensatoire jour", "congé/ferrier/compensatoire"])
            delai_consomme = lire_valeur_delai(["délai consommé jour", "délai consommé"])

            # Résultat
            jours_restants = lire_valeur_delai("jours restant pour exécution")
            depassement = lire_valeur_delai("ecart en jour en moins")
            ecart_pct = lire_valeur_delai("ecart en %")

            if ecart_pct <= 1 and ecart_pct not in (0,):
                ecart_pct = ecart_pct * 100

            texte = (
                f"Chantier : {nom_chantier}\n\n"

                f"DÉLAI CORRIGÉ\n"
                f"Délai soumission jour : {delai_soumission_jour:.2f} j\n"
                f"Jour accordés avenant en plus : {jours_avenant_plus:.2f} j\n"
                f"Jour accordés dû à une attente Technique : {jours_attente_technique:.2f} j\n"
                f"Jour accordés dû à imprévu, prévisible non visible : {jours_imprevu:.2f} j\n"
                f"Délai corrigé en jour : {delai_corrige_jour:.2f} j\n\n"

                f"DÉLAI CONSOMMÉ\n"
                f"Jours actif consommé : {jours_consommes:.2f} j\n"
                f"Jour accordés avenant en moins : {jours_avenant_moins:.2f} j\n"
                f"Jour intempéries : {jours_intemperies:.2f} j\n"
                f"Congé/ferrier/compensatoire : {conges:.2f} j\n"
                f"Délai consommé : {delai_consomme:.2f} j\n\n"

                f"RÉSULTAT\n"
                f"Jours restant pour exécution : {jours_restants:.2f} j\n"
                f"Ecart en jour en moins : {depassement:.2f} j\n"
                f"Ecart en % : {ecart_pct:.1f} %"
            )

            # Popup
            messagebox.showinfo("📊 Pilotage délai", texte)

            # PDF
            wb.close()
            pdf_path = os.path.join(dossier_chantier, "pilotage_delai.pdf")
            doc = SimpleDocTemplate(
                pdf_path,
                pagesize=landscape(A4),
                leftMargin=10 * mm,
                rightMargin=10 * mm,
                topMargin=8 * mm,
                bottomMargin=8 * mm,
            )

            elements = []
            graphique_tmp = None
            logo_path = str(dossier_base() / "logo_jt_bati.png")
            date_du_jour = datetime.now().strftime("%d/%m/%Y")

            style_titre = ParagraphStyle(
                "DelaiTitre",
                fontName="Helvetica-Bold",
                fontSize=18,
                leading=20,
                alignment=TA_CENTER,
                textColor=colors.HexColor("#163A59"),
            )
            style_meta_label = ParagraphStyle(
                "DelaiMetaLabel",
                fontName="Helvetica-Bold",
                fontSize=8.5,
                leading=10,
                textColor=colors.HexColor("#36556F"),
            )
            style_meta_valeur = ParagraphStyle(
                "DelaiMetaValeur",
                fontName="Helvetica",
                fontSize=8.8,
                leading=10.5,
                textColor=colors.black,
            )
            style_bloc_titre = ParagraphStyle(
                "DelaiBlocTitre",
                fontName="Helvetica-Bold",
                fontSize=9.4,
                leading=10.5,
                alignment=TA_CENTER,
                textColor=colors.white,
            )
            style_libelle = ParagraphStyle(
                "DelaiLibelle",
                fontName="Helvetica",
                fontSize=7.8,
                leading=8.8,
                textColor=colors.black,
            )
            style_valeur = ParagraphStyle(
                "DelaiValeur",
                fontName="Helvetica-Bold",
                fontSize=8.1,
                leading=9.2,
                alignment=TA_RIGHT,
                textColor=colors.black,
            )
            style_carte_label = ParagraphStyle(
                "DelaiCarteLabel",
                fontName="Helvetica-Bold",
                fontSize=7.4,
                leading=8.4,
                textColor=colors.HexColor("#5A7086"),
            )
            style_carte_valeur = ParagraphStyle(
                "DelaiCarteValeur",
                fontName="Helvetica-Bold",
                fontSize=10.8,
                leading=12,
                alignment=TA_RIGHT,
                textColor=colors.HexColor("#173C5A"),
            )
            style_carte_valeur_finale = ParagraphStyle(
                "DelaiCarteValeurFinale",
                fontName="Helvetica-Bold",
                fontSize=11.6,
                leading=12.8,
                alignment=TA_RIGHT,
                textColor=colors.HexColor("#9C3B32"),
            )
            style_bandeau = ParagraphStyle(
                "DelaiBandeau",
                fontName="Helvetica-Bold",
                fontSize=9.1,
                leading=10.5,
                textColor=colors.HexColor("#173C5A"),
            )
            style_ecart_titre = ParagraphStyle(
                "DelaiEcartTitre",
                fontName="Helvetica-Bold",
                fontSize=8,
                leading=9,
                alignment=TA_CENTER,
                textColor=colors.HexColor("#5A7086"),
            )
            style_ecart_valeur = ParagraphStyle(
                "DelaiEcartValeur",
                fontName="Helvetica-Bold",
                fontSize=10.4,
                leading=12,
                alignment=TA_CENTER,
                textColor=colors.HexColor("#163A59"),
            )

            logo_cell = Spacer(1, 1)
            if os.path.exists(logo_path):
                try:
                    logo_cell = Image(logo_path, width=30 * mm, height=12 * mm)
                    logo_cell.hAlign = "LEFT"
                except:
                    logo_cell = Spacer(1, 1)

            meta_table = Table([
                [
                    Paragraph("CHANTIER", style_meta_label),
                    Paragraph(nom_chantier, style_meta_valeur),
                    Paragraph("DATE", style_meta_label),
                    Paragraph(date_du_jour, style_meta_valeur),
                ]
            ], colWidths=[18 * mm, 47 * mm, 12 * mm, 22 * mm])
            meta_table.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))

            header_table = Table([
                [logo_cell, Paragraph("PILOTAGE DÉLAI", style_titre), meta_table]
            ], colWidths=[34 * mm, 158 * mm, 74 * mm])
            header_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F8FB")),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#D9E4EE")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
                ("ALIGN", (2, 0), (2, 0), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]))
            elements.append(header_table)
            elements.append(Spacer(1, 4))

            valeurs_graphique = [delai_corrige_jour, delai_consomme, jours_restants]
            valeurs_plot = [max(delai_corrige_jour, 0), max(delai_consomme, 0), max(jours_restants, 0)]
            etiquettes_graphique = ["Délai\ncorrigé", "Délai\nconsommé", "Jours\nrestants"]
            couleurs_graphique = ["#1F4E79", "#E69138", "#70AD47"]

            fig, ax = plt.subplots(figsize=(3.2, 4.0), facecolor="white")
            barres = ax.bar(
                range(3),
                valeurs_plot,
                width=0.55,
                color=couleurs_graphique,
                edgecolor="#FFFFFF",
                linewidth=0.8,
                zorder=3,
            )

            max_valeur = max([abs(v) for v in valeurs_plot] + [1])
            marge = max_valeur * 0.18
            ax.set_ylim(0, max(max(valeurs_plot), 1) + marge)
            ax.set_xticks(range(3), etiquettes_graphique)
            ax.tick_params(axis="x", labelsize=7.6, colors="#173C5A", length=0, pad=6)
            ax.tick_params(axis="y", labelsize=7.2, colors="#6A7E90")
            ax.grid(axis="y", color="#D9E4EE", linewidth=0.8, alpha=0.8, zorder=0)
            ax.set_axisbelow(True)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_color("#D9E4EE")
            ax.spines["bottom"].set_color("#D9E4EE")

            for barre, valeur in zip(barres, valeurs_graphique):
                ax.text(
                    barre.get_x() + barre.get_width() / 2,
                    barre.get_height() + max(marge * 0.08, 0.3),
                    f"{valeur:.2f} j",
                    ha="center",
                    va="bottom",
                    fontsize=7.6,
                    fontweight="bold",
                    color="#173C5A",
                )

            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                graphique_tmp = tmp.name

            plt.tight_layout()
            plt.savefig(graphique_tmp, dpi=220, bbox_inches="tight", facecolor="white")
            plt.close(fig)

            graphique = Image(graphique_tmp, width=72 * mm, height=72 * mm)
            graphique.hAlign = "CENTER"

            largeur_blocs_gauche = [96 * mm, 29 * mm]
            largeur_blocs_droite = [38 * mm, 24 * mm]

            bloc_1 = Table([
                ["CONSTRUCTION", ""],
                [Paragraph("Délai soumission jour", style_libelle), Paragraph(f"{delai_soumission_jour:.2f} j", style_valeur)],
                [Paragraph("+ Jour accordés avenant en plus", style_libelle), Paragraph(f"{jours_avenant_plus:.2f} j", style_valeur)],
                [Paragraph("+ Jour accordés dû à une attente Technique", style_libelle), Paragraph(f"{jours_attente_technique:.2f} j", style_valeur)],
                [Paragraph("+ Jour accordés dû à imprévu, prévisible non visible", style_libelle), Paragraph(f"{jours_imprevu:.2f} j", style_valeur)],
                [Paragraph("= Délai corrigé en jour", style_libelle), Paragraph(f"{delai_corrige_jour:.2f} j", style_valeur)],
            ], colWidths=largeur_blocs_gauche)
            bloc_1.hAlign = "CENTER"
            bloc_1.setStyle(TableStyle([
                ("SPAN", (0, 0), (1, 0)),
                ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
                ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (1, 0), 11),
                ("BACKGROUND", (0, 1), (1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#F8FBFF")]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C5D3")),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                ("BACKGROUND", (0, 5), (1, 5), colors.HexColor("#FFF2CC")),
                ("FONTNAME", (0, 5), (1, 5), "Helvetica-Bold"),
            ]))

            bloc_2 = Table([
                ["CONSOMMÉ", ""],
                [Paragraph("Jours actif consommé", style_libelle), Paragraph(f"{jours_consommes:.2f} j", style_valeur)],
                [Paragraph("- Jour accordés avenant en moins", style_libelle), Paragraph(f"{jours_avenant_moins:.2f} j", style_valeur)],
                [Paragraph("- Jour intempéries", style_libelle), Paragraph(f"{jours_intemperies:.2f} j", style_valeur)],
                [Paragraph("- Congé/ferrier/compensatoire", style_libelle), Paragraph(f"{conges:.2f} j", style_valeur)],
                [Paragraph("= Délai consommé", style_libelle), Paragraph(f"{delai_consomme:.2f} j", style_valeur)],
            ], colWidths=largeur_blocs_gauche)
            bloc_2.hAlign = "CENTER"
            bloc_2.setStyle(TableStyle([
                ("SPAN", (0, 0), (1, 0)),
                ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#4F81BD")),
                ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
                ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (1, 0), 11),
                ("BACKGROUND", (0, 1), (1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#F7FAFD")]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C5D3")),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                ("BACKGROUND", (0, 5), (1, 5), colors.HexColor("#FFF2CC")),
                ("FONTNAME", (0, 5), (1, 5), "Helvetica-Bold"),
            ]))

            bloc_3 = Table([
                ["RÉSULTAT", ""],
                [Paragraph("Jours restant pour exécution", style_libelle), Paragraph(f"{jours_restants:.2f} j", style_valeur)],
                [Paragraph("Ecart en jour en moins", style_libelle), Paragraph(f"{depassement:.2f} j", style_valeur)],
                [Paragraph("Ecart en %", style_libelle), Paragraph(f"{ecart_pct:.1f} %", style_valeur)],
            ], colWidths=largeur_blocs_gauche)
            bloc_3.hAlign = "CENTER"
            bloc_3.setStyle(TableStyle([
                ("SPAN", (0, 0), (1, 0)),
                ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#C0504D")),
                ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
                ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (1, 0), 11),
                ("BACKGROUND", (0, 1), (1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#FCF7F7")]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#C9B1B1")),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                ("BACKGROUND", (0, 1), (1, 1), colors.HexColor("#E7E7E7")),
                ("BACKGROUND", (0, 3), (1, 3), colors.HexColor("#D9EAD3")),
                ("FONTNAME", (0, 3), (1, 3), "Helvetica-Bold"),
            ]))

            colonne_gauche = Table([
                [bloc_1],
                [bloc_2],
                [bloc_3],
            ], colWidths=[125 * mm])
            colonne_gauche.setStyle(TableStyle([
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))

            bloc_graphique = Table([
                [Paragraph("RÉSUMÉ VISUEL", style_bloc_titre)],
                [graphique],
            ], colWidths=[70 * mm])
            bloc_graphique.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#173C5A")),
                ("BACKGROUND", (0, 1), (0, 1), colors.white),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B7C5D3")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]))

            indicateurs_table = Table([
                [Paragraph("Délai corrigé", style_carte_label), Paragraph(f"{delai_corrige_jour:.2f} j", style_carte_valeur)],
                [Paragraph("Jours actif consommé", style_carte_label), Paragraph(f"{jours_consommes:.2f} j", style_carte_valeur)],
                [Paragraph("Délai consommé", style_carte_label), Paragraph(f"{delai_consomme:.2f} j", style_carte_valeur)],
                [Paragraph("Jours restants", style_carte_label), Paragraph(f"{jours_restants:.2f} j", style_carte_valeur)],
                [Paragraph("Ecart en %", style_carte_label), Paragraph(f"{ecart_pct:.1f} %", style_carte_valeur_finale)],
            ], colWidths=largeur_blocs_droite)
            indicateurs_table.setStyle(TableStyle([
                ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#F6F9FC")]),
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#D6E0EA")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("BACKGROUND", (0, 4), (1, 4), colors.HexColor("#D9EAD3")),
            ]))

            indicateurs_cles = Table([
                [Paragraph("INDICATEURS CLÉS", style_bloc_titre)],
                [indicateurs_table],
            ], colWidths=[70 * mm])
            indicateurs_cles.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#365F91")),
                ("BACKGROUND", (0, 1), (0, 1), colors.white),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B7C5D3")),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]))

            colonne_droite = Table([
                [bloc_graphique],
                [indicateurs_cles],
            ], colWidths=[70 * mm])
            colonne_droite.setStyle(TableStyle([
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))

            corps = Table([
                [colonne_gauche, colonne_droite]
            ], colWidths=[128 * mm, 70 * mm])
            corps.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))
            elements.append(corps)
            elements.append(Spacer(1, 4))

            footer_row_height = 18 * mm
            footer_split_height = 6.5 * mm

            ecart_table = Table([
                [Paragraph("ECART %", style_ecart_titre)],
                [Paragraph(f"{ecart_pct:.1f} %", style_ecart_valeur)],
            ], colWidths=[30 * mm], rowHeights=[footer_split_height, footer_row_height - footer_split_height])
            ecart_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EAF1F8")),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#B7C5D3")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))

            footer_explication = Table([
                [
                    Paragraph(
                        f"Jours restant pour exécution = Délai corrigé - Délai consommé<br/>{delai_corrige_jour:.2f} j - {delai_consomme:.2f} j = {jours_restants:.2f} j",
                        style_bandeau,
                    )
                ]
            ], colWidths=[168 * mm], rowHeights=[footer_row_height])
            footer_explication.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#FFF2CC")),
                ("BOX", (0, 0), (0, 0), 0.7, colors.HexColor("#D6B656")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))

            footer_table = Table([
                [footer_explication, ecart_table]
            ], colWidths=[168 * mm, 30 * mm], rowHeights=[footer_row_height])
            footer_table.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))
            elements.append(footer_table)

            _publier_pdf_atomique(doc, elements, pdf_path)

            if graphique_tmp and os.path.exists(graphique_tmp):
                try:
                    os.unlink(graphique_tmp)
                except:
                    pass

            ouvrir_chemin(pdf_path)

        except Exception as e:
            if 'graphique_tmp' in locals() and graphique_tmp and os.path.exists(graphique_tmp):
                try:
                    os.unlink(graphique_tmp)
                except:
                    pass
            messagebox.showerror("Erreur pilotage délai", str(e))

    def rendements_chantier(self):
        import os
        import subprocess
        import tempfile
        from datetime import datetime
        import matplotlib.pyplot as plt
        from openpyxl import load_workbook
        from tkinter import messagebox
        from reportlab.platypus import SimpleDocTemplate, Spacer, Table, TableStyle, Image, Paragraph
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import mm
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.enums import TA_RIGHT, TA_CENTER
        from reportlab.lib.pagesizes import A4, landscape

        def nombre(valeur):
            try:
                if valeur is None:
                    return 0
                if isinstance(valeur, str):
                    valeur = valeur.replace("%", "").replace("€", "").replace(" ", "").replace(",", ".").strip()
                    if valeur in ("", "-", "—"):
                        return 0
                return float(valeur)
            except:
                return 0

        try:
            nom_chantier = self._nom_chantier_selectionne()
            if not nom_chantier:
                return

            dossier_chantier = self._dossier_chantier_selectionne()
            fichier_delai = os.path.join(dossier_chantier, "Rendement.xlsx")

            if not os.path.exists(fichier_delai):
                messagebox.showerror("Erreur", f"Fichier introuvable :\n{fichier_delai}")
                return

            wb = _charger_valeurs_excel(fichier_delai, None, Path(fichier_delai).stem, max_col=15, lecture_seule=False)
            ws = wb.active

            heures_soumission = nombre(ws["O65"].value)
            heures_avenant_plus = nombre(ws["O69"].value)
            heures_attente_technique = nombre(ws["O73"].value)
            heures_imprevu = nombre(ws["O75"].value)
            heures_consommees = nombre(ws["O67"].value)
            heures_avenant_moins = nombre(ws["O71"].value)
            heures_corrigees = heures_soumission + heures_avenant_plus + heures_attente_technique + heures_imprevu - heures_avenant_moins
            heures_consommees_finales = heures_consommees + heures_avenant_moins
            heures_delai_consomme = heures_consommees_finales
            heures_restantes = heures_corrigees - heures_delai_consomme
            ecart_heures = heures_consommees_finales - heures_corrigees
            ecart_pct = nombre(ws["O82"].value)

            if ecart_pct <= 1 and ws["O82"].value not in (None, "", 0):
                ecart_pct = ecart_pct * 100

            valorisation_finale = ecart_heures * 55
            wb.close()

            texte = (
                f"Chantier : {nom_chantier}\n\n"

                f"CONSTRUCTION DU DÉLAI\n"
                f"Heures soumission : {heures_soumission:.2f} h\n"
                f"Heures accordées avenant en plus : {heures_avenant_plus:.2f} h\n"
                f"Heures accordées dû à une attente Technique : {heures_attente_technique:.2f} h\n"
                f"Heures accordées dû à imprévu, prévisible non visible : {heures_imprevu:.2f} h\n"
                f"Délai corrigé en heures : {heures_corrigees:.2f} h\n\n"

                f"CONSOMMÉ\n"
                f"Heures consommées : {heures_consommees:.2f} h\n"
                f"Heures accordées avenant en moins : {heures_avenant_moins:.2f} h\n"
                f"Heures consommées finales / délai consommé en heures : {heures_delai_consomme:.2f} h\n\n"

                f"RÉSULTAT\n"
                f"Heures restantes pour exécution : {heures_restantes:.2f} h\n"
                f"Écart en heures : {ecart_heures:.2f} h\n"
                f"Pourcentage : {ecart_pct:.1f} %\n"
                f"Valorisation financière finale : {valorisation_finale:,.2f} €"
            )

            messagebox.showinfo("⏱ Rendement chantier", texte)

            pdf_path = os.path.join(dossier_chantier, "rendement_chantier.pdf")
            doc = SimpleDocTemplate(
                pdf_path,
                pagesize=landscape(A4),
                leftMargin=10 * mm,
                rightMargin=10 * mm,
                topMargin=8 * mm,
                bottomMargin=8 * mm,
            )

            elements = []
            graphique_tmp = None

            logo_path = str(dossier_base() / "logo_jt_bati.png")
            date_du_jour = datetime.now().strftime("%d/%m/%Y")
            style_titre = ParagraphStyle(
                "RendementTitre",
                fontName="Helvetica-Bold",
                fontSize=18,
                leading=20,
                alignment=TA_CENTER,
                textColor=colors.HexColor("#163A59"),
            )
            style_meta_label = ParagraphStyle(
                "RendementMetaLabel",
                fontName="Helvetica-Bold",
                fontSize=8.5,
                leading=10,
                textColor=colors.HexColor("#36556F"),
            )
            style_meta_valeur = ParagraphStyle(
                "RendementMetaValeur",
                fontName="Helvetica",
                fontSize=8.8,
                leading=10.5,
                textColor=colors.black,
            )
            style_bloc_titre = ParagraphStyle(
                "RendementBlocTitre",
                fontName="Helvetica-Bold",
                fontSize=9.6,
                leading=11,
                alignment=TA_CENTER,
                textColor=colors.white,
            )
            style_libelle = ParagraphStyle(
                "RendementLibelle",
                fontName="Helvetica",
                fontSize=8.6,
                leading=9.8,
                textColor=colors.black,
            )
            style_valeur = ParagraphStyle(
                "RendementValeur",
                fontName="Helvetica-Bold",
                fontSize=8.8,
                leading=10,
                alignment=TA_RIGHT,
                textColor=colors.black,
            )
            style_carte_label = ParagraphStyle(
                "RendementCarteLabel",
                fontName="Helvetica-Bold",
                fontSize=7.6,
                leading=8.6,
                textColor=colors.HexColor("#5A7086"),
            )
            style_carte_valeur = ParagraphStyle(
                "RendementCarteValeur",
                fontName="Helvetica-Bold",
                fontSize=11.2,
                leading=12.5,
                alignment=TA_RIGHT,
                textColor=colors.HexColor("#173C5A"),
            )
            style_carte_valeur_finale = ParagraphStyle(
                "RendementCarteValeurFinale",
                fontName="Helvetica-Bold",
                fontSize=12.2,
                leading=13.5,
                alignment=TA_RIGHT,
                textColor=colors.HexColor("#9C3B32"),
            )
            style_bandeau = ParagraphStyle(
                "RendementBandeau",
                fontName="Helvetica-Bold",
                fontSize=9.3,
                leading=11,
                textColor=colors.HexColor("#173C5A"),
            )
            style_bandeau_valeur = ParagraphStyle(
                "RendementBandeauValeur",
                fontName="Helvetica-Bold",
                fontSize=10.3,
                leading=11.5,
                alignment=TA_RIGHT,
                textColor=colors.HexColor("#9C3B32"),
            )
            style_taux_titre = ParagraphStyle(
                "RendementTauxTitre",
                fontName="Helvetica-Bold",
                fontSize=8,
                leading=9,
                alignment=TA_CENTER,
                textColor=colors.HexColor("#5A7086"),
            )
            style_taux_valeur = ParagraphStyle(
                "RendementTauxValeur",
                fontName="Helvetica-Bold",
                fontSize=10.5,
                leading=12,
                alignment=TA_CENTER,
                textColor=colors.HexColor("#163A59"),
            )

            logo_cell = Spacer(1, 1)
            if os.path.exists(logo_path):
                try:
                    logo_cell = Image(logo_path, width=30 * mm, height=12 * mm)
                    logo_cell.hAlign = "LEFT"
                except:
                    logo_cell = Spacer(1, 1)

            meta_table = Table([
                [
                    Paragraph("CHANTIER", style_meta_label),
                    Paragraph(nom_chantier, style_meta_valeur),
                    Paragraph("DATE", style_meta_label),
                    Paragraph(date_du_jour, style_meta_valeur),
                ]
            ], colWidths=[18 * mm, 47 * mm, 12 * mm, 22 * mm])
            meta_table.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))

            header_table = Table([
                [logo_cell, Paragraph("RENDEMENT CHANTIER", style_titre), meta_table]
            ], colWidths=[34 * mm, 158 * mm, 74 * mm])
            header_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F8FB")),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#D9E4EE")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
                ("ALIGN", (2, 0), (2, 0), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]))
            elements.append(header_table)
            elements.append(Spacer(1, 5))

            valeurs_graphique = [heures_corrigees, heures_delai_consomme, heures_restantes]
            etiquettes_graphique = ["Délai\ncorrigé", "Consommé\nfinal", "Heures\nrestantes"]
            couleurs_graphique = ["#1F4E79", "#E69138", "#70AD47"]

            fig, ax = plt.subplots(figsize=(3.25, 4.1), facecolor="white")
            barres = ax.bar(
                range(3),
                valeurs_graphique,
                width=0.55,
                color=couleurs_graphique,
                edgecolor="#FFFFFF",
                linewidth=0.8,
                zorder=3,
            )

            max_valeur = max([abs(v) for v in valeurs_graphique] + [1])
            marge = max_valeur * 0.18
            ax.set_ylim(0, max(max(valeurs_graphique), 1) + marge)
            ax.set_xticks(range(3), etiquettes_graphique)
            ax.tick_params(axis="x", labelsize=7.8, colors="#173C5A", length=0, pad=6)
            ax.tick_params(axis="y", labelsize=7.5, colors="#6A7E90")
            ax.grid(axis="y", color="#D9E4EE", linewidth=0.8, alpha=0.8, zorder=0)
            ax.set_axisbelow(True)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_color("#D9E4EE")
            ax.spines["bottom"].set_color("#D9E4EE")

            for barre, valeur in zip(barres, valeurs_graphique):
                ax.text(
                    barre.get_x() + barre.get_width() / 2,
                    barre.get_height() + max(marge * 0.08, 0.6),
                    f"{valeur:.2f} h",
                    ha="center",
                    va="bottom",
                    fontsize=7.8,
                    fontweight="bold",
                    color="#173C5A",
                )

            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                graphique_tmp = tmp.name

            plt.tight_layout()
            plt.savefig(graphique_tmp, dpi=220, bbox_inches="tight", facecolor="white")
            plt.close(fig)

            graphique = Image(graphique_tmp, width=74 * mm, height=74 * mm)
            graphique.hAlign = "CENTER"

            largeur_blocs_gauche = [95 * mm, 30 * mm]
            largeur_blocs_droite = [38 * mm, 24 * mm]

            bloc_1 = Table([
                ["CONSTRUCTION DU DÉLAI", ""],
                [Paragraph("Heures soumission", style_libelle), Paragraph(f"{heures_soumission:.2f} h", style_valeur)],
                [Paragraph("+ Heures accordées avenant en plus", style_libelle), Paragraph(f"{heures_avenant_plus:.2f} h", style_valeur)],
                [Paragraph("+ Heures accordées dû à une attente Technique", style_libelle), Paragraph(f"{heures_attente_technique:.2f} h", style_valeur)],
                [Paragraph("+ Heures accordées dû à imprévu, prévisible non visible", style_libelle), Paragraph(f"{heures_imprevu:.2f} h", style_valeur)],
                [Paragraph("= Délai corrigé en heures", style_libelle), Paragraph(f"{heures_corrigees:.2f} h", style_valeur)],
            ], colWidths=largeur_blocs_gauche)
            bloc_1.hAlign = "CENTER"
            bloc_1.setStyle(TableStyle([
                ("SPAN", (0, 0), (1, 0)),
                ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
                ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (1, 0), 12),
                ("BACKGROUND", (0, 1), (1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#F8FBFF")]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C5D3")),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("BACKGROUND", (0, 5), (1, 5), colors.HexColor("#FFF2CC")),
                ("FONTNAME", (0, 5), (1, 5), "Helvetica-Bold"),
            ]))

            bloc_2 = Table([
                ["CONSOMMÉ", ""],
                [Paragraph("- Heures consommées", style_libelle), Paragraph(f"{heures_consommees:.2f} h", style_valeur)],
                [Paragraph("- Heures accordées avenant en moins", style_libelle), Paragraph(f"{heures_avenant_moins:.2f} h", style_valeur)],
                [Paragraph("= Heures consommées finales / délai consommé en heures", style_libelle), Paragraph(f"{heures_delai_consomme:.2f} h", style_valeur)],
            ], colWidths=largeur_blocs_gauche)
            bloc_2.hAlign = "CENTER"
            bloc_2.setStyle(TableStyle([
                ("SPAN", (0, 0), (1, 0)),
                ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#4F81BD")),
                ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
                ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (1, 0), 12),
                ("BACKGROUND", (0, 1), (1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#F7FAFD")]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C5D3")),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("BACKGROUND", (0, 3), (1, 3), colors.HexColor("#FFF2CC")),
                ("FONTNAME", (0, 3), (1, 3), "Helvetica-Bold"),
            ]))

            bloc_3 = Table([
                ["RÉSULTAT", ""],
                [Paragraph("Heures restantes pour exécution", style_libelle), Paragraph(f"{heures_restantes:.2f} h", style_valeur)],
                [Paragraph("Écart en heures", style_libelle), Paragraph(f"{ecart_heures:.2f} h", style_valeur)],
                [Paragraph("Pourcentage d’avancement", style_libelle), Paragraph(f"{ecart_pct:.1f} %", style_valeur)],
                [Paragraph("Valorisation financière finale (Écart en heures × 55 €)", style_libelle), Paragraph(f"{valorisation_finale:,.2f} €", style_valeur)],
            ], colWidths=largeur_blocs_gauche)
            bloc_3.hAlign = "CENTER"
            bloc_3.setStyle(TableStyle([
                ("SPAN", (0, 0), (1, 0)),
                ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#C0504D")),
                ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
                ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (1, 0), 12),
                ("BACKGROUND", (0, 1), (1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (1, -1), [colors.white, colors.HexColor("#FCF7F7")]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#C9B1B1")),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("BACKGROUND", (0, 1), (1, 1), colors.HexColor("#E7E7E7")),
                ("BACKGROUND", (0, 4), (1, 4), colors.HexColor("#D9EAD3")),
                ("FONTNAME", (0, 4), (1, 4), "Helvetica-Bold"),
            ]))

            colonne_gauche = Table([
                [bloc_1],
                [bloc_2],
                [bloc_3],
            ], colWidths=[125 * mm])
            colonne_gauche.setStyle(TableStyle([
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))

            bloc_graphique = Table([
                [Paragraph("RÉSUMÉ VISUEL", style_bloc_titre)],
                [graphique],
            ], colWidths=[70 * mm])
            bloc_graphique.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#173C5A")),
                ("BACKGROUND", (0, 1), (0, 1), colors.white),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B7C5D3")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]))

            indicateurs_table = Table([
                [Paragraph("Délai corrigé en heures", style_carte_label), Paragraph(f"{heures_corrigees:.2f} h", style_carte_valeur)],
                [Paragraph("Heures consommées finales", style_carte_label), Paragraph(f"{heures_delai_consomme:.2f} h", style_carte_valeur)],
                [Paragraph("Heures restantes", style_carte_label), Paragraph(f"{heures_restantes:.2f} h", style_carte_valeur)],
                [Paragraph("Écart en heures", style_carte_label), Paragraph(f"{ecart_heures:.2f} h", style_carte_valeur)],
                [Paragraph("Pourcentage d’avancement", style_carte_label), Paragraph(f"{ecart_pct:.1f} %", style_carte_valeur)],
                [Paragraph("Valorisation financière finale", style_carte_label), Paragraph(f"{valorisation_finale:,.2f} €", style_carte_valeur_finale)],
            ], colWidths=largeur_blocs_droite)
            indicateurs_table.setStyle(TableStyle([
                ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#F6F9FC")]),
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#D6E0EA")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("BACKGROUND", (0, 5), (1, 5), colors.HexColor("#D9EAD3")),
            ]))

            indicateurs_cles = Table([
                [Paragraph("INDICATEURS CLÉS", style_bloc_titre)],
                [indicateurs_table],
            ], colWidths=[70 * mm])
            indicateurs_cles.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#365F91")),
                ("BACKGROUND", (0, 1), (0, 1), colors.white),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B7C5D3")),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]))

            colonne_droite = Table([
                [bloc_graphique],
                [indicateurs_cles],
            ], colWidths=[70 * mm])
            colonne_droite.setStyle(TableStyle([
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))

            corps = Table([
                [colonne_gauche, colonne_droite]
            ], colWidths=[128 * mm, 70 * mm])
            corps.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))
            elements.append(corps)
            elements.append(Spacer(1, 4))

            footer_row_height = 18 * mm
            footer_split_height = 6.5 * mm

            taux_table = Table([
                [Paragraph("TAUX HORAIRE", style_taux_titre)],
                [Paragraph("55.00 €/h", style_taux_valeur)],
            ], colWidths=[30 * mm], rowHeights=[footer_split_height, footer_row_height - footer_split_height])
            taux_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EAF1F8")),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#B7C5D3")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))

            footer_explication = Table([
                [
                    Paragraph(
                        f"Valorisation financière finale = Écart en heures × 55 €/h<br/>{ecart_heures:.2f} h × 55 €/h = {valorisation_finale:,.2f} €",
                        style_bandeau,
                    )
                ]
            ], colWidths=[168 * mm], rowHeights=[footer_row_height])
            footer_explication.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#FFF2CC")),
                ("BOX", (0, 0), (0, 0), 0.7, colors.HexColor("#D6B656")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))

            footer_table = Table([
                [footer_explication, taux_table]
            ], colWidths=[168 * mm, 30 * mm], rowHeights=[footer_row_height])
            footer_table.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))
            elements.append(footer_table)

            _publier_pdf_atomique(doc, elements, pdf_path)

            if graphique_tmp and os.path.exists(graphique_tmp):
                try:
                    os.unlink(graphique_tmp)
                except:
                    pass

            ouvrir_chemin(pdf_path)

        except Exception as e:
            if 'graphique_tmp' in locals() and graphique_tmp and os.path.exists(graphique_tmp):
                try:
                    os.unlink(graphique_tmp)
                except:
                    pass
            messagebox.showerror("Erreur rendement chantier", str(e))

    def _chemin_chantier_selectionne(self) -> Path | None:
        sel = self.tree.selection()
        if not sel:
            return None
        return Path(sel[0])

    def choisir_dossier_chantiers(self):
        global _DOSSIER_BASE_MEMOIRE
        choix = filedialog.askdirectory(title="Choisir Horizon_Chantier_Data ou le dossier Chantiers")
        if not choix:
            return
        try:
            memoriser_base(choix)
            _DOSSIER_BASE_MEMOIRE = Path(choix)
            self.refresh_liste()
        except OSError as e:
            messagebox.showerror("Emplacement des chantiers", str(e))

    def refresh_liste(self) -> None:
        try:
            emplacement = dossier_chantiers()
        except (OSError, ValueError) as e:
            self.lbl_path.config(text=str(e))
            for item in self.tree.get_children():
                self.tree.delete(item)
            messagebox.showerror("Emplacement des chantiers", str(e))
            return
        self.lbl_path.config(text=f"Emplacement: {emplacement}")
        for item in self.tree.get_children():
            self.tree.delete(item)

        d = emplacement
        fichiers = sorted(d.glob("*.json"))
        fichiers_illisibles = []

        for p in fichiers:
            if p.stat().st_size == 0:
                continue
            try:
                chantier = lire_json(p)
                info = infos_chantier(chantier)
                av = info["avancement"]
                try:
                    av_txt = f"{float(av):.0f}%"
                except Exception:
                    av_txt = str(av)

                self.tree.insert(
                    "",
                    "end",
                    iid=str(p),
                    values=(info["nom"], info["client"], info["etat"], av_txt, info["debut"], info["fin"]),
                )
            except Exception as e:
                fichiers_illisibles.append(f"{p.name} ({e})")
                continue

        if fichiers_illisibles:
            messagebox.showwarning(
                "Chantier",
                "Fichier chantier illisible :\n" + "\n".join(fichiers_illisibles)
            )

    def voir_clients(self) -> None:
        clients = {}
        for chemin in sorted(dossier_chantiers().glob("*.json")):
            try:
                donnees = lire_json(chemin)
                info = infos_chantier(donnees)
                nom_client = str(info.get("client", "") or "").strip()
                if not nom_client:
                    continue
                entree = clients.setdefault(
                    nom_client.casefold(),
                    {"nom": nom_client, "chantiers": [], "chemins": [], "donnees": []},
                )
                entree["chantiers"].append(str(info.get("nom", "") or chemin.stem))
                entree["chemins"].append(chemin)
                entree["donnees"].append(donnees)
            except Exception:
                continue

        win = tk.Toplevel(self)
        win.title("Clients")
        win.geometry("850x520")
        win.minsize(650, 350)
        win.transient(self)

        cadre = ttk.Frame(win, padding=12)
        cadre.pack(fill="both", expand=True)
        ttk.Label(
            cadre,
            text=f"Clients enregistrés : {len(clients)}",
            font=("Helvetica", 15, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        zone_liste = ttk.Frame(cadre)
        zone_liste.pack(fill="both", expand=True)
        colonnes = ("Client", "Nombre", "Chantiers")
        liste = ttk.Treeview(zone_liste, columns=colonnes, show="headings")
        for colonne in colonnes:
            liste.heading(colonne, text=colonne)
        liste.column("Client", width=220, anchor="w")
        liste.column("Nombre", width=75, anchor="center", stretch=False)
        liste.column("Chantiers", width=480, anchor="w")
        clients_par_chemin = {}

        def entree_selectionnee():
            selection = liste.selection()
            if not selection:
                messagebox.showinfo("Clients", "Sélectionnez un client.", parent=win)
                return None
            return clients_par_chemin.get(selection[0])

        def selectionner_chantier():
            entree = entree_selectionnee()
            if not entree:
                return
            chemin = entree["chemins"][0]
            if self.tree.exists(str(chemin)):
                self.tree.selection_set(str(chemin))
                self.tree.focus(str(chemin))
                self.tree.see(str(chemin))
            win.destroy()

        def afficher_fiche(_evenement=None):
            entree = entree_selectionnee()
            if not entree:
                return
            fiche = tk.Toplevel(win)
            fiche.title(f"Fiche client — {entree['nom']}")
            fiche.geometry("780x500")
            fiche.minsize(650, 400)
            fiche.transient(win)
            contenu = ttk.Frame(fiche, padding=14)
            contenu.pack(fill="both", expand=True)
            ttk.Label(contenu, text=entree["nom"], font=("Helvetica", 18, "bold")).pack(
                anchor="w", pady=(0, 12)
            )
            onglets = ttk.Notebook(contenu)
            onglets.pack(fill="both", expand=True)
            libelles = (
                ("client", "Client"),
                ("personne_contact", "Personne de contact"),
                ("telephone", "Téléphone"),
                ("email", "E-mail"),
                ("adresse", "Adresse"),
                ("chantier", "Chantier"),
                ("type_etat", "Type d'état"),
                ("etat", "État"),
                ("avancement", "Avancement"),
            )
            for index, donnees in enumerate(entree["donnees"]):
                nom_chantier = str(donnees.get("chantier", "") or entree["chantiers"][index])
                page = ttk.Frame(onglets, padding=12)
                onglets.add(page, text=nom_chantier)
                page.columnconfigure(1, weight=1)
                for ligne, (cle, libelle) in enumerate(libelles):
                    valeur = donnees.get(cle, "")
                    if cle == "avancement" and valeur != "":
                        valeur = f"{valeur} %"
                    ttk.Label(page, text=f"{libelle} :", font=("Helvetica", 11, "bold")).grid(
                        row=ligne, column=0, sticky="nw", padx=(0, 12), pady=4
                    )
                    ttk.Label(page, text=str(valeur or "—"), wraplength=510).grid(
                        row=ligne, column=1, sticky="nw", pady=4
                    )
            ttk.Button(contenu, text="Fermer", command=fiche.destroy).pack(anchor="e", pady=(10, 0))

        for entree in sorted(clients.values(), key=lambda item: item["nom"].casefold()):
            premier_chemin = entree["chemins"][0]
            clients_par_chemin[str(premier_chemin)] = entree
            liste.insert(
                "", "end", iid=str(premier_chemin),
                values=(entree["nom"], len(entree["chantiers"]), ", ".join(entree["chantiers"])),
            )

        defilement = ttk.Scrollbar(zone_liste, orient="vertical", command=liste.yview)
        liste.configure(yscrollcommand=defilement.set)
        liste.pack(side="left", fill="both", expand=True)
        defilement.pack(side="right", fill="y")
        liste.bind("<Double-1>", afficher_fiche)
        bas = ttk.Frame(cadre)
        bas.pack(fill="x", pady=(10, 0))
        ttk.Label(bas, text="Double-cliquez sur un client pour ouvrir sa fiche complète.").pack(side="left")
        ttk.Button(bas, text="Fermer", command=win.destroy).pack(side="right")
        ttk.Button(bas, text="Voir la fiche complète", command=afficher_fiche).pack(side="right", padx=(0, 8))
        ttk.Button(bas, text="Sélectionner le chantier", command=selectionner_chantier).pack(side="right", padx=(0, 8))

    def _pr_path_and_hint(self) -> tuple[Path, str]:
        p = self._chemin_chantier_selectionne()
        if not p:
            raise ValueError("Sélectionne un chantier dans la liste.")
        dossier_chantier = self._dossier_depuis_json(p)
        pr_path = _normaliser_dossier_data(dossier_chantier) / "prix_de_revient.xlsx"
        if not pr_path.exists():
            raise FileNotFoundError(f"PR introuvable : {pr_path}")
        hint = pr_path.name
        return pr_path, hint

    def prix_revient(self) -> None:
        p = self._chemin_chantier_selectionne()
        if not p:
            messagebox.showwarning("Prix de revient", "Sélectionne un chantier dans la liste.")
            return
        dossier_chantier = self._dossier_depuis_json(p)
        try:
            pr_path = _normaliser_dossier_data(dossier_chantier) / "prix_de_revient.xlsx"
            if not self._preparer_ouverture_document(pr_path):
                return
            listes_verifiees = open_pr(dossier_chantier)
            if not listes_verifiees:
                messagebox.showinfo("Prix de revient",
                                    "PR ouvert. Un verrou Excel empêche actuellement la réparation des listes. "
                                    "Après avoir enregistré et fermé le PR dans Excel, rouvrez-le depuis Horizon "
                                    "Chantier si les listes sont absentes.")
            self._planifier_rappel_pdf_historique_ouverture()
        except Exception as e:
            messagebox.showerror("Prix de revient", str(e))

    # ======================================================
    # ✅ OBJECTIF UNIQUE : AUTOMATISATION TEXTE -> COLONNE P
    # Matière  : P3:P12
    # MainOeuvre: P17:P26
    # ======================================================
    def recherche_matiere(self) -> None:
        try:
            pr_path, hint = self._pr_path_and_hint()
            if not self._preparer_ouverture_document(pr_path):
                return
            _prepare_excel_file_before_write(pr_path)
            ouvrir_chemin(pr_path)
            self._planifier_rappel_pdf_historique_ouverture()
            time.sleep(0.3)

            items = read_biblio_colA_openpyxl(pr_path, "Bibliothèque_Matière")

            def on_pick(txt: str):
                target = excel_find_first_empty_cell(hint, "Chiffrage", "P3:P12")
                if not target:
                    messagebox.showwarning("Matière", "Zone P3:P12 pleine ou Excel pas prêt.")
                    return
                excel_set_cell_value(hint, "Chiffrage", target, txt)

            popup_search_20(self, "Matière → Chiffrage (P3:P12)", items, on_pick)

        except Exception as e:
            messagebox.showerror("Matière", str(e))

    def recherche_mo(self) -> None:
        try:
            pr_path, hint = self._pr_path_and_hint()
            if not self._preparer_ouverture_document(pr_path):
                return
            _prepare_excel_file_before_write(pr_path)
            ouvrir_chemin(pr_path)
            self._planifier_rappel_pdf_historique_ouverture()
            time.sleep(0.3)

            items = read_biblio_colA_openpyxl(pr_path, "Bibliothèque_MainOeuvre")

            def on_pick(txt: str):
                target = excel_find_first_empty_cell(hint, "Chiffrage", "P17:P26")
                if not target:
                    messagebox.showwarning("Main-d'œuvre", "Zone P17:P26 pleine ou Excel pas prêt.")
                    return
                excel_set_cell_value(hint, "Chiffrage", target, txt)

            popup_search_20(self, "Main-d’œuvre → Chiffrage (P17:P26)", items, on_pick)

        except Exception as e:
            messagebox.showerror("Main-d'œuvre", str(e))

    # ---------------------------
    # Boutons non prioritaires (inchangés)
    # ---------------------------
    
    def calcul_reca(self) -> None:
        import os
        from datetime import datetime
        from tkinter import messagebox

        p = self._chemin_chantier_selectionne()
        if not p:
            messagebox.showwarning("Calcul / Reca", "Sélectionne un chantier.")
            return

        dossier = self._dossier_chantier_selectionne()
        chemin_pr = _normaliser_dossier_data(dossier) / "prix_de_revient.xlsx"

        if not chemin_pr.exists():
            messagebox.showwarning(
                "Date du PR",
                "Le fichier prix_de_revient.xlsx est introuvable."
            )
            return

        timestamp = os.path.getmtime(chemin_pr)
        date_modif = datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y %H:%M:%S")

        messagebox.showinfo(
            "Date du PR",
            f"Dernier enregistrement du PR :\n{date_modif}\n\n"
            "⚠️ Si tu viens de modifier le PR :\n"
            "➡️ enregistre dans Excel (Ctrl+S sur Windows, Cmd+S sur Mac)\n"
            "➡️ puis clique sur Calcul / Reca"
        )

        try:
            nb_importes = self.importer_pr_dans_etat()
            if nb_importes:
                doublons = getattr(self, "_dernier_import_pr_doublons", [])
                precision = ("\n\nCodes répétés avec des prix différents : " + ", ".join(doublons)
                             + ".\nComme auparavant, le dernier prix du PR est utilisé pour chaque code." if doublons else "")
                messagebox.showinfo("Calcul / Reca", f"{nb_importes} article(s) importé(s) dans l'état d'avancement.\n"
                                    "Ouvrez puis enregistrez l'état dans Excel pour actualiser ses formules avant la suite des calculs."
                                    + precision)
            else:
                messagebox.showwarning("Calcul / Reca", "Aucun article correspondant : aucun prix importé.")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _dossier_chantier_selectionne(self) -> Path:
        p = self._chemin_chantier_selectionne()
        if not p:
            raise ValueError("Sélectionne un chantier dans la liste.")
        return self._dossier_depuis_json(p)

    def importer_pr_dans_etat(self):
        self._dernier_import_pr_doublons = []
        dossier = self._dossier_chantier_selectionne()
        chemin_pr = dossier / "data" / "prix_de_revient.xlsx"
        chemin_etat = _selectionner_etat(dossier)
        if not chemin_pr.is_file():
            raise FileNotFoundError(f"PR introuvable : {chemin_pr}")
        signature = _signature_excel(chemin_etat)
        signature_pr = _signature_excel(chemin_pr, verifier_verrou=False)
        wb_pr = wb_etat = None
        try:
            wb_pr = _charger_pr_controle(chemin_pr)
            ws_pr = wb_pr["Chiffrage"]
            # Même règle métier qu'avant : prix HT en F, 27 lignes après l'article.
            lignes_pr = list(ws_pr.iter_rows(max_col=6, values_only=True))
            pr_map = {}
            doublons = set()
            for row, valeurs in enumerate(lignes_pr, 1):
                article = _article_pr(valeurs[1]) if row >= 3 else ""
                if not article:
                    continue
                valeur = lignes_pr[row + 26][5] if row + 26 < len(lignes_pr) else None
                if not isinstance(valeur, (int, float)) or isinstance(valeur, bool) or not math.isfinite(valeur):
                    raise ValueError(f"Prix HT indisponible pour l'article {article} (F{row + 27}). "
                                     "Vérifiez et enregistrez le PR dans Excel. Aucun import effectué.")
                prix = float(valeur)
                if article in pr_map and pr_map[article] != prix:
                    doublons.add(article)
                # Préserver la priorité historique : dernière occurrence du code dans le PR.
                pr_map[article] = prix
            if not pr_map:
                raise ValueError("Aucun article avec prix HT trouvé dans le PR.")
            wb_etat = load_workbook(chemin_etat, keep_vba=Path(chemin_etat).suffix.lower() == ".xlsm")
            ws_etat = _feuille_etat(wb_etat)
            def normaliser_libelle(val):
                txt = str(val or "").strip().lower()
                txt = txt.replace("é", "e").replace("è", "e").replace("ê", "e").replace("ë", "e")
                txt = txt.replace("à", "a").replace("â", "a").replace("ä", "a")
                txt = txt.replace("ù", "u").replace("û", "u").replace("ü", "u")
                txt = txt.replace("î", "i").replace("ï", "i")
                txt = txt.replace("ô", "o").replace("ö", "o")
                txt = txt.replace("ç", "c")
                txt = re.sub(r"\s+", " ", txt)
                return txt

            def est_modele_marche_public():
                for ligne in range(1, min(ws_etat.max_row, 25) + 1):
                    h = normaliser_libelle(ws_etat[f"H{ligne}"].value)
                    i = normaliser_libelle(ws_etat[f"I{ligne}"].value)
                    j = normaliser_libelle(ws_etat[f"J{ligne}"].value)
                    if "en chiffres" in h and "en lettres" in i and "somme" in j:
                        return True
                return False

            if _est_metre_public(ws_etat) and doublons:
                raise ValueError("Prix PR contradictoires pour : " + ", ".join(sorted(doublons)))
            marche_public = est_modele_marche_public()
            colonne = "H" if marche_public else "L"
            correspondances = []
            for row in (_postes_publics(ws_etat) if _est_metre_public(ws_etat) else range(24, ws_etat.max_row + 1)):
                article = _article_pr(ws_etat[f"B{row}"].value)
                if article in pr_map:
                    cellule = ws_etat[f"{colonne}{row}"]
                    if cellule.data_type == "f":
                        raise ValueError(f"Le prix {cellule.coordinate} contient une formule. "
                                         "Import annulé pour la préserver.")
                    correspondances.append((cellule, pr_map[article]))
            if not correspondances:
                return 0
            if _signature_excel(chemin_pr, verifier_verrou=False) != signature_pr:
                raise ValueError("Le PR a changé pendant la lecture. Relancez l'import.")
            for cellule, prix in correspondances:
                cellule.value = prix
            _protect_formula_cells(wb_etat, chemin_etat)
            _sauver_excel_atomique(wb_etat, chemin_etat, signature)
            self._dernier_import_pr_doublons = sorted(doublons)
            return len(correspondances)
        finally:
            if wb_pr is not None:
                _fermer_classeur(wb_pr)
            if wb_etat is not None:
                _fermer_classeur(wb_etat)

    def exporter_prix_bordereau_public(self):
        temporaire = None
        try:
            dossier = self._dossier_chantier_selectionne()
            config = _config_public(dossier)
            if not config:
                raise ValueError("Aucune copie de bordereau public configurée pour ce chantier.")
            original, copie = dossier / config["original"], dossier / config["copie"]
            signature = _signature_excel(original)
            signature_copie = _signature_excel(copie)
            fd, nom = tempfile.mkstemp(suffix=".xlsx", dir=dossier)
            os.close(fd)
            temporaire = Path(nom)
            nombre = _exporter_prix_public(original, copie, temporaire, config["sha256_original"])
            if not messagebox.askyesno("Transfert final des prix",
                    f"Transférer les {nombre} prix unitaires de {copie.name} vers {original.name} ?\n"
                    "Seules les cellules de prix seront modifiées. Une sauvegarde sera conservée."):
                return
            if _signature_excel(original) != signature or _signature_excel(copie) != signature_copie:
                raise ValueError("Un fichier a changé pendant la préparation. Relancez le transfert.")
            _backup_excel_before_write(original)
            if _signature_excel(original) != signature:
                raise ValueError("L'original a changé pendant la sauvegarde.")
            os.replace(temporaire, original)
            import hashlib
            config["sha256_original"] = hashlib.sha256(original.read_bytes()).hexdigest()
            ecrire_json(dossier / "bordereau_public.json", config)
            messagebox.showinfo("Prix transférés", f"{nombre} prix transférés. Ouvrez l’original dans Excel pour actualiser ses totaux.")
        except Exception as e:
            messagebox.showerror("Transfert des prix", str(e))
        finally:
            if temporaire is not None:
                temporaire.unlink(missing_ok=True)

    def ouvrir_chantier(self) -> None:
        p = self._chemin_chantier_selectionne()
        if not p:
            messagebox.showwarning("Chantier", "Sélectionne un chantier dans la liste.")
            return

        dossier = self._dossier_depuis_json(p)
        numero = re.match(r"^(\d+)", dossier.name)
        if not numero:
            messagebox.showerror(
                "Bordereau introuvable",
                "Le nom du chantier doit commencer par son numéro "
                "(exemple : 002_Saint-Symphorien).",
            )
            return

        nom_attendu = f"Bordereau_{numero.group(1)}.PDF"
        try:
            bordereau_pdf = next(
                fichier
                for fichier in dossier.iterdir()
                if fichier.is_file() and fichier.name.lower() == nom_attendu.lower()
            )
        except (FileNotFoundError, StopIteration):
            messagebox.showerror(
                "Bordereau introuvable",
                f"Fichier absent dans le dossier du chantier :\n{nom_attendu}",
            )
            return

        try:
            ouvrir_chemin(bordereau_pdf)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le bordereau PDF.\n{e}")

    def dossier_du_chantier(self) -> None:
        p = self._chemin_chantier_selectionne()
        if not p:
            messagebox.showwarning("Chantier", "Sélectionne un chantier dans la liste.")
            return
        dossier = self._dossier_depuis_json(p)
        if not self._preparer_ouverture_document(dossier):
            return
        ouvrir_dossier(dossier)
        if self._doit_rappeler_pdf_historique(dossier):
            self._planifier_rappel_pdf_historique_ouverture()

    def ouvrir_dossier_sauvegarde(self) -> None:
        p = self._chemin_chantier_selectionne()
        if not p:
            messagebox.showwarning("Sauvegarde", "Sélectionne un chantier dans la liste.")
            return

        dossier_sauvegarde = self._dossier_depuis_json(p) / "Sauvegarde"
        if not dossier_sauvegarde.exists():
            messagebox.showwarning(
                "Sauvegarde",
                f"Aucun dossier de sauvegarde pour ce chantier :\n{dossier_sauvegarde}",
            )
            return

        ouvrir_dossier(dossier_sauvegarde)

    def nouveau_chantier(self) -> None:
        try:
            _, _, etats_disponibles = _modeles_creation_chantier(dossier_chantiers())
        except Exception as e:
            messagebox.showerror("Nouveau chantier", str(e))
            return
        win = tk.Toplevel(self)
        win.title("Nouveau chantier")
        win.geometry("820x570")
        win.resizable(False, False)

        client_var = tk.StringVar(value="")
        chantier_var = tk.StringVar(value="")
        adresse_var = tk.StringVar(value="")
        type_etat_var = tk.StringVar(value="Privé")

        frame = ttk.Frame(win, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        win.columnconfigure(0, weight=1)
        win.rowconfigure(0, weight=1)

        ttk.Label(frame, text="Nom du client").grid(row=0, column=0, sticky="w")
        e_client = ttk.Entry(frame, textvariable=client_var, width=90)
        e_client.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 8))

        ttk.Label(frame, text="Nom du chantier").grid(row=2, column=0, sticky="w")
        e_chantier = ttk.Entry(frame, textvariable=chantier_var, width=90)
        e_chantier.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(2, 8))

        ttk.Label(frame, text="Adresse").grid(row=4, column=0, sticky="w")
        e_adresse = ttk.Entry(frame, textvariable=adresse_var, width=90)
        e_adresse.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(2, 12))

        entrees_contact = {}
        for ligne, (cle, libelle) in enumerate((
            ("personne_contact", "Personne de contact"),
            ("telephone", "Téléphone"),
            ("email", "E-mail"),
        )):
            ttk.Label(frame, text=libelle).grid(row=6 + ligne * 2, column=0, sticky="w")
            entree = ttk.Entry(frame, width=90)
            entree.grid(row=7 + ligne * 2, column=0, columnspan=2, sticky="ew", pady=(2, 8))
            entrees_contact[cle] = entree

        ttk.Label(frame, text="Type d'état d'avancement").grid(row=12, column=0, sticky="w")
        choix_etat = ttk.Frame(frame)
        choix_etat.grid(row=13, column=0, columnspan=2, sticky="w", pady=(2, 12))
        for type_etat in ("Privé", "Public"):
            ttk.Radiobutton(choix_etat, text=type_etat, variable=type_etat_var,
                            value=type_etat).pack(side="left", padx=(0, 18))

        if any(t not in etats_disponibles for t in ("Privé", "Public")):
            ttk.Label(frame, text=(
                "Les modèles d’état Public/Privé sont incomplets. Le chantier et le prix de revient\n"
                "peuvent être créés ; les états manquants devront être ajoutés."
            ), wraplength=760).grid(row=14, column=0, columnspan=2, sticky="w", pady=(0, 10))

        def valider():
            nom_client = e_client.get().strip()
            nom_chantier = e_chantier.get().strip()
            adresse_chantier = e_adresse.get().strip()

            self.nom_client = nom_client
            self.nom_chantier = nom_chantier
            self.adresse_chantier = adresse_chantier

            if not nom_chantier:
                messagebox.showwarning("Nouveau chantier", "Renseigne le nom du chantier.")
                return

            nom_fichier = nom_chantier.replace("/", "_").replace("\\", "_")
            base = dossier_chantiers()
            json_path = base / f"{nom_fichier}.json"
            dossier_path = base / nom_fichier

            if json_path.exists() or (dossier_path.exists() and (
                not dossier_path.is_dir() or dossier_path.is_symlink() or any(dossier_path.iterdir())
            )):
                messagebox.showwarning("Nouveau chantier", "Ce chantier existe déjà.")
                return

            chantier = {
                "chantier": nom_fichier,
                "client": nom_client,
                "adresse": adresse_chantier,
                **{cle: entree.get().strip() for cle, entree in entrees_contact.items()},
                "type_etat": type_etat_var.get(),
                "etat": "Devis",
                "avancement": 0,
                "bordereau": {"source": {}, "articles": {}},
            }

            try:
                _creer_documents_chantier(base, nom_fichier, chantier)
                self.refresh_liste()
                if self.tree.exists(str(json_path)):
                    self.tree.selection_set(str(json_path))
                    self.tree.focus(str(json_path))
                win.destroy()
                if lire_json(json_path).get("modele_etat_manquant"):
                    messagebox.showinfo("Chantier créé", (
                        "Le chantier est créé et son prix de revient est disponible.\n\n"
                        "Modèles d’état absents : " + ", ".join(lire_json(json_path).get("modeles_etat_manquants", [])) + ".\n"
                        "Aucun état d’un autre chantier ou logiciel n’a été utilisé."
                    ))
                return
            except Exception as e:
                messagebox.showerror("Nouveau chantier", f"Impossible de créer le chantier.\n{e}")
                return

            win.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=15, column=0, columnspan=2, sticky="e")
        ttk.Button(btns, text="Annuler", command=win.destroy).pack(side="right")
        ttk.Button(btns, text="Valider", command=valider).pack(side="right", padx=(0, 8))

        frame.columnconfigure(0, weight=1)
        e_client.focus_set()
        win.grab_set()



    def modifier_chantier(self) -> None:
        json_path = self._chemin_chantier_selectionne()
        if not json_path:
            messagebox.showwarning("Modifier chantier", "Sélectionne un chantier dans la liste.")
            return
        try:
            chantier = lire_json(json_path)
        except Exception as e:
            messagebox.showerror("Modifier chantier", f"Impossible de lire le chantier.\n{e}")
            return

        infos = infos_chantier(chantier)
        win = tk.Toplevel(self)
        win.title("Modifier chantier")
        win.geometry("820x410")
        win.resizable(False, False)
        win.transient(self)
        frame = ttk.Frame(win, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        win.columnconfigure(0, weight=1)
        win.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        champs = (
            ("Nom du client", infos["client"]),
            ("Nom du chantier", infos["nom"]),
            ("Adresse", chantier.get("adresse", "")),
            ("Personne de contact", chantier.get("personne_contact", "")),
            ("Téléphone", chantier.get("telephone", "")),
            ("E-mail", chantier.get("email", "")),
        )
        entrees = []
        for index, (libelle, valeur) in enumerate(champs):
            ligne = index * 2
            ttk.Label(frame, text=libelle).grid(row=ligne, column=0, sticky="w")
            entree = ttk.Entry(frame, width=90)
            entree.grid(row=ligne + 1, column=0, columnspan=2, sticky="ew", pady=(2, 8 if index < 5 else 12))
            entree.insert(0, str(valeur or ""))
            entrees.append(entree)

        def valider():
            nom_client, nom_chantier, adresse, personne_contact, telephone, email = (
                entree.get().strip() for entree in entrees
            )
            if not nom_chantier:
                messagebox.showwarning("Modifier chantier", "Renseigne le nom du chantier.", parent=win)
                return

            nouveau_nom_fichier = nom_chantier.replace("/", "_").replace("\\", "_")
            nouveau_json_path = json_path.parent / f"{nouveau_nom_fichier}.json"
            ancien_dossier_path = self._dossier_depuis_json(json_path)
            nouveau_dossier_path = json_path.parent / nouveau_nom_fichier

            try:
                if nouveau_json_path.exists() and os.path.samefile(json_path, nouveau_json_path):
                    nouveau_json_path = json_path
                if (ancien_dossier_path.exists() and nouveau_dossier_path.exists()
                        and os.path.samefile(ancien_dossier_path, nouveau_dossier_path)):
                    nouveau_dossier_path = ancien_dossier_path
            except FileNotFoundError:
                pass

            if nouveau_json_path != json_path and nouveau_json_path.exists():
                messagebox.showwarning("Modifier chantier", "Un chantier portant ce nom existe déjà.", parent=win)
                return
            if nouveau_dossier_path != ancien_dossier_path and nouveau_dossier_path.exists():
                messagebox.showwarning("Modifier chantier", "Un dossier chantier portant ce nom existe déjà.", parent=win)
                return

            donnees_modifiees = dict(chantier)
            if "nom" in donnees_modifiees:
                donnees_modifiees["nom"] = nouveau_nom_fichier
            if "chantier" in donnees_modifiees or "nom" not in donnees_modifiees:
                donnees_modifiees["chantier"] = nouveau_nom_fichier
            donnees_modifiees.update({
                "client": nom_client,
                "adresse": adresse,
                "personne_contact": personne_contact,
                "telephone": telephone,
                "email": email,
            })

            dossier_renomme = False
            json_renomme = False
            try:
                if ancien_dossier_path.exists() and nouveau_dossier_path != ancien_dossier_path:
                    ancien_dossier_path.rename(nouveau_dossier_path)
                    dossier_renomme = True
                if nouveau_json_path != json_path:
                    json_path.rename(nouveau_json_path)
                    json_renomme = True
                ecrire_json(nouveau_json_path, donnees_modifiees)
            except Exception as e:
                try:
                    if json_renomme and nouveau_json_path.exists():
                        nouveau_json_path.rename(json_path)
                    if dossier_renomme and nouveau_dossier_path.exists():
                        nouveau_dossier_path.rename(ancien_dossier_path)
                except Exception:
                    pass
                messagebox.showerror("Modifier chantier", f"Impossible de modifier le chantier.\n{e}", parent=win)
                return

            self.nom_client = nom_client
            self.nom_chantier = nom_chantier
            self.refresh_liste()
            if self.tree.exists(str(nouveau_json_path)):
                self.tree.selection_set(str(nouveau_json_path))
                self.tree.focus(str(nouveau_json_path))
            win.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=12, column=0, columnspan=2, sticky="e")
        ttk.Button(btns, text="Annuler", command=win.destroy).pack(side="right")
        ttk.Button(btns, text="Valider", command=valider).pack(side="right", padx=(0, 8))
        entrees[0].focus_set()
        win.grab_set()


    def supprimer_chantier(self) -> None:
        messagebox.showinfo("Info", "Supprimer chantier (à brancher)")


    def postes_metre(self) -> None:
        messagebox.showinfo("Info", "Postes / Métré (à faire)")

print("JE SUIS DANS LE BON FICHIER")

if __name__ == "__main__":
    app = HorizonChantierApp()
    app.mainloop()
