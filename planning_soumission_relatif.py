"""Remplace le calendrier daté par des jours relatifs dans la copie de soumission."""

import os
from pathlib import Path
import re
import shutil
import tempfile
from copy import deepcopy
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.utils import column_index_from_string, get_column_letter


def adapter_planning_soumission(chemin, sauvegarder=None):
    chemin = Path(chemin)
    if chemin.with_name('~$' + chemin.name).exists():
        raise PermissionError('Enregistrez et fermez le planning de soumission dans Excel.')
    with ZipFile(chemin) as archive:
        original = archive.read('xl/worksheets/sheet1.xml').decode('utf-8')
        if '1+7*(Semaine_Affichage-1)' in original and '<v>0</v>' in original:
            return False
        if 'WEEKDAY(' not in original or 'Semaine_Affichage-1' not in original:
            raise ValueError('Calendrier du planning de soumission non reconnu.')
        styles = archive.read('xl/styles.xml').decode('utf-8')
        bloc = re.search(r'<cellXfs\b[^>]*>.*?</cellXfs>', styles, re.S)
        if bloc is None:
            raise ValueError('Styles du planning non reconnus.')
        definitions = list(ET.fromstring(bloc.group().replace('<cellXfs',
                           '<cellXfs xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', 1)))
        nouveaux = []
        variantes = {}

        def style_pour(numero, nombre=False, fond=False):
            if numero is None or (not nombre and not fond):
                return numero
            cle = (numero, nombre, fond)
            if cle not in variantes:
                definition = deepcopy(definitions[numero])
                if nombre:
                    definition.set('numFmtId', '0')
                if fond:
                    definition.set('fillId', '0')
                variantes[cle] = len(definitions) + len(nouveaux)
                nouveaux.append(ET.tostring(definition, encoding='unicode'))
            return variantes[cle]

        def cellule(match):
            xml = match.group()
            ref = re.search(r'\br="([A-Z]+)(\d+)"', xml)
            if ref is None:
                return xml
            col, ligne = ref.group(1), int(ref.group(2))
            index = column_index_from_string(col)
            style = re.search(r'\bs="(\d+)"', xml.split('>', 1)[0])
            style = int(style.group(1)) if style else None

            def style_xml(valeur):
                return f' s="{valeur}"' if valeur is not None else ''

            if ref.group() == 'r="C17"':
                return '<c r="C17" t="inlineStr"><is><t>Commencement (jour 0)</t></is></c>'
            if ref.group() == 'r="E17"':
                return f'<c r="E17"{style_xml(style_pour(style, nombre=True))}><v>0</v></c>'
            if ligne == 18 and index >= 8 and '<f>' in xml:
                fin = get_column_letter(index + 6)
                return f'<c r="{col}18"{style_xml(style)} t="str"><f>"Jours "&amp;{col}19&amp;" à "&amp;{fin}19</f></c>'
            if ligne == 19 and index >= 8 and '<f>' in xml:
                if col == 'H':
                    formule = '1+7*(Semaine_Affichage-1)'
                else:
                    formule = re.search(r'<f[^>]*>(.*?)</f>', xml, re.S).group(1)
                nouveau_style = style_pour(style, nombre=True)
                return (f'<c r="{col}19"{style_xml(nouveau_style)}><f>{formule}</f>'
                        f'<v>{index - 7}</v></c>')
            if ligne == 20 and index >= 8:
                return f'<c r="{col}20"{style_xml(style)}/>'
            if 21 <= ligne <= 35 and col in ('E', 'F'):
                if style is not None:
                    nouveau_style = style_pour(style, nombre=True)
                    xml = re.sub(r'\bs="\d+"', f's="{nouveau_style}"', xml, count=1)
                if '<f' in xml:
                    if col == 'E':
                        xml = xml.replace('Début_Projet+0', 'Début_Projet+1')
                        xml = xml.replace('D&#233;but_Projet+0', 'D&#233;but_Projet+1')
                    xml = re.sub(r'<v>.*?</v>', '', xml, count=1, flags=re.S)
                return xml
            if 21 <= ligne <= 35 and index >= 8 and style is not None:
                definition = definitions[style]
                if definition.get('fillId') == '34':
                    nouveau_style = style_pour(style, fond=True)
                    return re.sub(r'\bs="\d+"', f's="{nouveau_style}"', xml, count=1)
            return xml

        contenu = re.sub(r'<c\b[^>]*?/>|<c\b[^>]*>.*?</c>', cellule, original, flags=re.S)
        if nouveaux:
            bloc_nouveau = bloc.group().replace('</cellXfs>', ''.join(nouveaux) + '</cellXfs>')
            bloc_nouveau = re.sub(r'count="\d+"', f'count="{len(definitions) + len(nouveaux)}"',
                                  bloc_nouveau, count=1)
            styles = styles[:bloc.start()] + bloc_nouveau + styles[bloc.end():]
        classeur = archive.read('xl/workbook.xml').decode('utf-8')
        if '<calcPr' in classeur:
            classeur = re.sub(r'<calcPr\b([^>]*)/>',
                              lambda m: '<calcPr' + re.sub(r'\sfullCalcOnLoad="[^"]*"', '', m.group(1))
                              + ' fullCalcOnLoad="1"/>', classeur, count=1)
        calculs = archive.read('xl/calcChain.xml').decode('utf-8') if 'xl/calcChain.xml' in archive.namelist() else None
        changements = {'xl/worksheets/sheet1.xml': contenu.encode('utf-8'),
                       'xl/styles.xml': styles.encode('utf-8'),
                       'xl/workbook.xml': classeur.encode('utf-8')}
        if calculs is not None:
            calculs = re.sub(r'<c\b[^>]*\br="[A-Z]+20"[^>]*/>', '', calculs)
            changements['xl/calcChain.xml'] = calculs.encode('utf-8')
        if sauvegarder:
            sauvegarder(chemin)
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
