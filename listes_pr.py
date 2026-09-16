"""Listes PR au format Excel standard, sans réécrire les cellules du classeur."""
import html
import os
from pathlib import Path
import re
import shutil
import tempfile
import unicodedata
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
BIBLIOS = ('Bibliothèque_Matière', 'Bibliothèque_MainOeuvre')


def _cle(value):
    return ''.join(c for c in unicodedata.normalize('NFKD', str(value or '').lower().replace('œ', 'oe'))
                   if c.isalnum())


def _listes(wb):
    if 'Chiffrage' not in wb.sheetnames or any(n not in wb.sheetnames for n in BIBLIOS):
        return []
    rows = list(wb['Chiffrage'].iter_rows(max_col=5, values_only=True))
    plages = {n: [] for n in BIBLIOS}
    for i, row in enumerate(rows[:-11]):
        bloc = _cle(row[1])
        if bloc not in ('matiere', 'maindoeuvre'):
            continue
        header = rows[i + 1]
        # Les anciens PR utilisent C ; les PR avec désignation du bordereau utilisent D.
        cols = [j for j, v in enumerate(header) if _cle(v) == 'designationmatiere']
        if not cols:
            continue
        col = get_column_letter(cols[0] + 1)
        nom = BIBLIOS[0 if bloc == 'matiere' else 1]
        plages[nom].append(f'{col}{i + 3}:{col}{i + 12}')
    resultat = []
    for nom, ranges in plages.items():
        if not ranges:
            continue
        fin = max((i for i, row in enumerate(wb[nom].iter_rows(max_col=1, values_only=True), 1)
                   if row[0] not in (None, '')), default=0)
        if not fin:
            raise ValueError(f'Bibliothèque vide : {nom}')
        dv = ET.Element('dataValidation', type='list', allowBlank='1', showDropDown='0',
                        showErrorMessage='1', sqref=' '.join(ranges))
        ET.SubElement(dv, 'formula1').text = f'INDIRECT("\'{nom}\'!$A$1:$A${fin}")'
        resultat.append(dv)
    return resultat


def reparer_listes_pr(path, sauvegarder=None):
    """Répare les listes reconnues ; conserve intégralement les autres entrées ZIP."""
    path = Path(path)
    wb = load_workbook(path, read_only=True)
    try:
        listes = _listes(wb)
    finally:
        wb.close()
    if not listes:
        return False
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        sheet = next(s for s in workbook.find(f'{{{NS}}}sheets') if s.get('name') == 'Chiffrage')
        rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        target = next(r.get('Target') for r in rels if r.get('Id') == rid)
        member = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        original = archive.read(member).decode('utf-8')
        # Retire seulement les anciennes validations des bibliothèques, y compris x14.
        def filtrer(match):
            return '' if any(n in html.unescape(match.group()) for n in BIBLIOS) else match.group()
        contenu = re.sub(r'<(?:\w+:)?dataValidation\b[^>]*>.*?</(?:\w+:)?dataValidation>',
                         filtrer, original, flags=re.S)
        bloc = re.search(r'<dataValidations\b[^>]*>(.*?)</dataValidations>', contenu, re.S)
        anciens = bloc.group(1) if bloc else ''
        nouveaux = anciens + ''.join(ET.tostring(d, encoding='unicode') for d in listes)
        count = len(re.findall(r'<dataValidation\b', nouveaux))
        remplacement = f'<dataValidations count="{count}">{nouveaux}</dataValidations>'
        if bloc:
            contenu = contenu[:bloc.start()] + remplacement + contenu[bloc.end():]
        else:
            # Position OOXML : après les fusions, avant les hyperliens et l'impression.
            pos = re.search(r'<(?:hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|drawing|legacyDrawing|extLst)\b|</worksheet>', contenu).start()
            contenu = contenu[:pos] + remplacement + contenu[pos:]
        # Les conteneurs x14 restants doivent annoncer leur nombre réel.
        def compter(match):
            texte = match.group()
            count = len(re.findall(r'<x14:dataValidation\b', texte))
            return re.sub(r'count="\d+"', f'count="{count}"', texte, count=1) if count else ''
        contenu = re.sub(r'<x14:dataValidations\b.*?</x14:dataValidations>', compter, contenu, flags=re.S)
        if contenu == original:
            return False
        if path.with_name('~$' + path.name).exists():
            raise PermissionError('Enregistrez et fermez le PR dans Excel, puis rouvrez-le depuis Horizon Chantier pour rétablir les listes.')
        if sauvegarder:
            sauvegarder(path)
        fd, tmp = tempfile.mkstemp(prefix='.listes_pr_', suffix='.xlsx', dir=path.parent)
        os.close(fd)
        try:
            with ZipFile(tmp, 'w') as sortie:
                for info in archive.infolist():
                    sortie.writestr(info, contenu.encode('utf-8') if info.filename == member else archive.read(info.filename))
            shutil.copymode(path, tmp)
        except Exception:
            Path(tmp).unlink(missing_ok=True)
            raise
    try:
        os.replace(tmp, path)
    finally:
        Path(tmp).unlink(missing_ok=True)
    return True
