"""Contenu propre au chantier dans le planning de soumission, sans réécrire Excel."""

from html import escape
import os
from pathlib import Path
import re
import shutil
import tempfile
from copy import deepcopy
from xml.etree import ElementTree as ET
from zipfile import ZipFile, is_zipfile


PREMIERE_LIGNE = 23
DERNIERE_LIGNE = 62


def _mettre_texte(xml, reference, valeur):
    motif = rf'<c\b[^>]*\br="{reference}"[^>]*?(?:/>|>.*?</c>)'
    ancien = re.search(motif, xml, re.S)
    if ancien is None:
        raise ValueError(f'Cellule {reference} absente du modèle de planning.')
    style = re.search(r'\bs="\d+"', ancien.group().split('>', 1)[0])
    style = f' {style.group()}' if style else ''
    nouveau = (f'<c r="{reference}"{style} t="inlineStr"><is><t>{escape(str(valeur))}'
               '</t></is></c>') if valeur else f'<c r="{reference}"{style}/>'
    return xml[:ancien.start()] + nouveau + xml[ancien.end():]


def _mettre_formule(xml, reference, formule):
    motif = rf'<c\b[^>]*\br="{reference}"[^>]*?(?:/>|>.*?</c>)'
    ancien = re.search(motif, xml, re.S)
    if ancien is None:
        raise ValueError(f'Cellule {reference} absente du modèle de planning.')
    style = re.search(r'\bs="\d+"', ancien.group().split('>', 1)[0])
    style = f' {style.group()}' if style else ''
    nouveau = f'<c r="{reference}"{style}><f>{escape(formule)}</f></c>' if formule else f'<c r="{reference}"{style}/>'
    return xml[:ancien.start()] + nouveau + xml[ancien.end():]


def _mettre_nombre(xml, reference, valeur):
    motif = rf'<c\b[^>]*\br="{reference}"[^>]*?(?:/>|>.*?</c>)'
    ancien = re.search(motif, xml, re.S)
    if ancien is None:
        raise ValueError(f'Cellule {reference} absente du modèle de planning.')
    style = re.search(r'\bs="\d+"', ancien.group().split('>', 1)[0])
    style = f' {style.group()}' if style else ''
    nouveau = f'<c r="{reference}"{style}><v>{valeur:g}</v></c>'
    return xml[:ancien.start()] + nouveau + xml[ancien.end():]


def _dessin(donnees, client, adresse):
    xml = donnees.decode('utf-8')
    for ancre in re.findall(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S):
        textes = re.findall(r'<a:t>(.*?)</a:t>', ancre, re.S)
        contenu = ' '.join(textes).casefold()
        if any(mot in contenu for mot in ('week-ends', 'jours ouvrables', 'intempéries', 'congés', 'stater')):
            xml = xml.replace(ancre, '', 1)
        elif 'maître de l’ouvrage' in contenu:
            valeurs = ['MAÎTRE DE L’OUVRAGE', client or '', adresse or '', '']
            position = iter(valeurs)
            nouveau = re.sub(r'<a:t>.*?</a:t>',
                             lambda _m: f'<a:t>{escape(next(position, ""))}</a:t>', ancre, flags=re.S)
            xml = xml.replace(ancre, nouveau, 1)
    return xml.encode('utf-8')


def changements_soumission(archive: ZipFile, nom='', client='', adresse='', postes=None, debuts=None):
    """Prépare seulement les parties OOXML nécessaires à la soumission.

    Les autres membres ZIP, les liens et les règles de mise en forme restent intacts.
    """
    feuille = archive.read('xl/worksheets/sheet1.xml').decode('utf-8')
    if '<sheetData' not in feuille or 'PlanningProjet' not in archive.read('xl/workbook.xml').decode('utf-8'):
        raise ValueError('Modèle de planning non reconnu.')
    feuille = _mettre_texte(feuille, 'B15', f'Planning de soumission - {nom}' if nom else 'Planning de soumission')
    if postes is not None:
        debuts = debuts or {}
        if len(postes) > DERNIERE_LIGNE - PREMIERE_LIGNE + 1:
            raise ValueError('Le planning actuel contient 40 lignes de tâches ; le PR en contient davantage.')
        if any(jours is None for _, _, jours in postes):
            raise ValueError('Certains postes du PR n’ont pas de résultat de jours enregistré en G.')
        feuille = _mettre_texte(feuille, 'B22', '')
        for ligne in range(PREMIERE_LIGNE, DERNIERE_LIGNE + 1):
            position = ligne - PREMIERE_LIGNE
            if position < len(postes):
                article, enonce, jours = postes[position]
                feuille = _mettre_texte(feuille, f'B{ligne}', article)
                feuille = _mettre_texte(feuille, f'C{ligne}', enonce)
                # Même calcul FIN = DÉBUT + durée - 1 ; seul le nombre de jours change.
                feuille = _mettre_formule(feuille, f'F{ligne}', f'E{ligne}+{jours - 1:g}')
                debut = debuts.get(article)
                if isinstance(debut, (int, float)) and not isinstance(debut, bool):
                    feuille = _mettre_nombre(feuille, f'E{ligne}', debut)
                elif isinstance(debut, str) and debut.startswith('='):
                    feuille = _mettre_formule(feuille, f'E{ligne}', debut[1:])
                else:
                    feuille = _mettre_formule(feuille, f'E{ligne}', 'Début_Projet+1')
                cellule_jours = re.search(rf'<c\b[^>]*\br="G{ligne}"[^>]*?(?:/>|>.*?</c>)', feuille, re.S).group()
                if '<f>' not in cellule_jours:
                    feuille = _mettre_formule(feuille, f'G{ligne}',
                        'IF(OR(ISBLANK(début_tâche),ISBLANK(fin_tâche)),"",fin_tâche-début_tâche+1)')
            else:
                for col in ('B', 'C', 'E', 'F', 'G'):
                    feuille = _mettre_formule(feuille, f'{col}{ligne}', '')
    styles = archive.read('xl/styles.xml').decode('utf-8')
    bloc = re.search(r'<cellXfs\b[^>]*>.*?</cellXfs>', styles, re.S)
    if bloc is None:
        raise ValueError('Styles du planning non reconnus.')
    definitions = list(ET.fromstring(bloc.group().replace('<cellXfs',
        '<cellXfs xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', 1)))
    variantes = {}
    nouveaux = []

    def enlever_fond(match):
        cellule = match.group()
        ref = re.search(r'\br="([A-Z]+)(\d+)"', cellule)
        style = re.search(r'\bs="(\d+)"', cellule.split('>', 1)[0])
        if not ref or not style or not 21 <= int(ref.group(2)) <= 62:
            return cellule
        colonne = ref.group(1)
        if len(colonne) == 1 and colonne < 'H':
            return cellule
        numero = int(style.group(1))
        if definitions[numero].get('fillId') != '34':
            return cellule
        if numero not in variantes:
            variante = deepcopy(definitions[numero])
            variante.set('fillId', '0')
            variantes[numero] = len(definitions) + len(nouveaux)
            nouveaux.append(ET.tostring(variante, encoding='unicode'))
        return cellule.replace(style.group(), f's="{variantes[numero]}"', 1)

    feuille = re.sub(r'<c\b[^>]*?/>|<c\b[^>]*>.*?</c>', enlever_fond, feuille, flags=re.S)
    changements = {'xl/worksheets/sheet1.xml': feuille.encode('utf-8')}
    if nouveaux:
        bloc_nouveau = bloc.group().replace('</cellXfs>', ''.join(nouveaux) + '</cellXfs>')
        bloc_nouveau = re.sub(r'count="\d+"', f'count="{len(definitions) + len(nouveaux)}"',
                              bloc_nouveau, count=1)
        styles = styles[:bloc.start()] + bloc_nouveau + styles[bloc.end():]
        changements['xl/styles.xml'] = styles.encode('utf-8')
    if 'xl/drawings/drawing1.xml' in archive.namelist():
        changements['xl/drawings/drawing1.xml'] = _dessin(
            archive.read('xl/drawings/drawing1.xml'), client, adresse)
    return {nom: contenu for nom, contenu in changements.items() if archive.read(nom) != contenu}


def preparer_copie_soumission(chemin, nom='', client='', adresse=''):
    """Personnalise une copie neuve du modèle sans toucher au classeur d'exécution."""
    chemin = Path(chemin)
    if not is_zipfile(chemin):
        return False
    with ZipFile(chemin) as archive:
        changements = changements_soumission(archive, nom, client, adresse)
        if not changements:
            return False
        fd, temporaire = tempfile.mkstemp(prefix='.planning_soumission_', suffix='.xlsx', dir=chemin.parent)
        os.close(fd)
        try:
            with ZipFile(temporaire, 'w') as sortie:
                for info in archive.infolist():
                    sortie.writestr(info, changements.get(info.filename, archive.read(info.filename)))
            shutil.copymode(chemin, temporaire)
        except Exception:
            Path(temporaire).unlink(missing_ok=True)
            raise
    try:
        os.replace(temporaire, chemin)
    finally:
        Path(temporaire).unlink(missing_ok=True)
    return True
