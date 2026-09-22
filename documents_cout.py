"""Documents par poste : liens relatifs, modification ciblée du ZIP Excel."""
import os
import posixpath
import re
import shutil
import tempfile
import unicodedata
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, unquote
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from xml.sax.saxutils import quoteattr

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _feuilles(z):
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    cibles = {r.get('Id'): r.get('Target') for r in rels}
    for feuille in ET.fromstring(z.read('xl/workbook.xml')).find(f'{{{S}}}sheets'):
        cible = cibles[feuille.get(f'{{{R}}}id')]
        membre = cible.lstrip('/') if cible.startswith('/') else posixpath.normpath('xl/' + cible)
        yield feuille.get('name'), membre


def _texte(cellule, chaines):
    if cellule is None:
        return ''
    valeur = cellule.find(f'{{{S}}}v')
    if cellule.get('t') == 's' and valeur is not None:
        return chaines[int(valeur.text)]
    if cellule.get('t') == 'inlineStr':
        return ''.join(t.text or '' for t in cellule.iter(f'{{{S}}}t'))
    return valeur.text if valeur is not None else ''


def _lire(z):
    chaines = []
    if 'xl/sharedStrings.xml' in z.namelist():
        chaines = [''.join(t.text or '' for t in si.iter(f'{{{S}}}t'))
                   for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
    for nom, membre in _feuilles(z):
        xml = ET.fromstring(z.read(membre))
        rp = posixpath.dirname(membre) + '/_rels/' + posixpath.basename(membre) + '.rels'
        rels = ET.fromstring(z.read(rp)) if rp in z.namelist() else ET.Element(f'{{{P}}}Relationships')
        cibles = {r.get('Id'): r.get('Target') for r in rels}
        liens = {h.get('ref'): cibles.get(h.get(f'{{{R}}}id'), '')
                 for h in xml.findall(f'{{{S}}}hyperlinks/{{{S}}}hyperlink')}
        postes = []
        for ligne in xml.findall(f'{{{S}}}sheetData/{{{S}}}row'):
            cellules = {c.get('r'): c for c in ligne}
            numero = ligne.get('r')
            libelle = _texte(cellules.get('A' + numero), chaines)
            if re.match(r'^\d{3,}\s*[·.\-]', libelle or ''):
                ref = 'B' + numero
                postes.append(dict(feuille=nom, membre=membre, cellule=ref,
                                   libelle=libelle, cible=liens.get(ref, ''),
                                   texte=_texte(cellules.get(ref), chaines)))
        yield membre, rp, xml, rels, postes


def chemin_document(classeur, cible):
    """N'accepte que des références relatives contenues dans ce chantier."""
    if not cible or re.match(r'^[a-zA-Z]+:', cible) or cible.startswith(('/', '\\')):
        return None
    racine = Path(classeur).resolve().parent
    chemin = (racine / unquote(cible).replace('\\', '/')).resolve()
    try:
        chemin.relative_to(racine)
    except ValueError:
        return None
    return chemin if chemin.is_file() else None


def lire_postes(classeur):
    with ZipFile(classeur) as z:
        return [p for *_, postes in _lire(z) for p in postes
                if p['texte'] != 'Entreprise']


def dossier_poste(classeur, poste):
    """Réutilise le dossier numéroté existant, sinon crée numéro + libellé."""
    racine = Path(classeur).resolve().parent
    base = racine / 'Documents_installation'
    base.resolve().relative_to(racine)
    base.mkdir(exist_ok=True)
    numero, libelle = re.match(r'^(\d{3,})\s*[·.\-]\s*(.*)', poste['libelle']).groups()
    candidats = sorted(p for p in base.iterdir()
                        if p.is_dir() and re.match(r'^' + numero + r'(?:_|$)', p.name))
    for chemin in candidats:
        chemin.resolve().relative_to(base.resolve())
    if len(candidats) > 1:
        raise ValueError(f'Plusieurs dossiers portent le numéro {numero} : ' + ', '.join(p.name for p in candidats))
    if candidats:
        return candidats[0]
    nom = unicodedata.normalize('NFKD', libelle).encode('ascii', 'ignore').decode()
    nom = re.sub(r'[^A-Za-z0-9]+', '_', nom).strip('_')[:100]
    dossier = base / (numero + ('_' + nom if nom else ''))
    dossier.mkdir(exist_ok=True)
    return dossier


def renommer_dossier_poste(classeur, poste):
    """Aligne le nom d'un dossier de poste sans toucher à ses documents."""
    actuel = dossier_poste(classeur, poste)
    if actuel.is_symlink():
        raise ValueError(f'Le dossier du poste est un lien symbolique : {actuel.name}')
    numero, libelle = re.match(r'^(\d{3,})\s*[·.\-]\s*(.*)', poste['libelle']).groups()
    nom = unicodedata.normalize('NFC', libelle)
    nom = re.sub(r'[<>:"/\\|?*\x00-\x1f\s]+', '_', nom).strip('._')[:100].rstrip('._')
    attendu = actuel.parent / (numero + ('_' + nom if nom else ''))
    if actuel == attendu:
        return actuel
    if attendu.exists():
        raise FileExistsError(f'Le dossier {attendu.name} existe déjà.')
    actuel.rename(attendu)
    return attendu


def renommer_libelle_poste(classeur, numero, libelle):
    """Change uniquement le libellé du poste demandé dans le classeur du chantier."""
    from module_engins_cout import _remplacer_cellule, _texte
    numero = str(numero).strip()
    libelle = libelle.strip()
    if not re.fullmatch(r'\d{3,}', numero) or not libelle:
        raise ValueError('Numéro ou nom de poste invalide.')
    classeur = Path(classeur)
    if classeur.with_name('~$' + classeur.name).exists():
        raise PermissionError('Fermez le coût de la sécurité dans Excel pour renommer ce poste.')
    attendu = f'{numero} · {libelle}'
    modifications = {}
    trouve = False
    with ZipFile(classeur) as archive:
        for membre, _, arbre, _, postes in _lire(archive):
            cellules = {c.get('r'): c for c in arbre.iter(f'{{{S}}}c')}
            contenu = archive.read(membre).decode('utf-8')
            for poste in postes:
                if not re.match(rf'^{re.escape(numero)}\s*[·.\-]', poste['libelle']):
                    continue
                trouve = True
                if poste['libelle'] == attendu:
                    continue
                ref = 'A' + poste['cellule'][1:]
                cellule = ET.fromstring(ET.tostring(cellules[ref]))
                _texte(cellule, attendu)
                contenu = _remplacer_cellule(contenu, cellule)
            if contenu.encode('utf-8') != archive.read(membre):
                modifications[membre] = contenu.encode('utf-8')
        if not trouve:
            raise ValueError(f'Le poste {numero} est absent du classeur du chantier.')
        if not modifications:
            return False
        historique = classeur.parent / 'Historique_Sauvegardes'
        historique.mkdir(exist_ok=True)
        shutil.copy2(classeur, historique / (datetime.now().strftime('%Y%m%d_%H%M%S_%f_') + 'Cout_avant_renommage_poste.xlsx'))
        fd, temporaire = tempfile.mkstemp(prefix='.renommer_poste_', suffix='.xlsx', dir=classeur.parent)
        os.close(fd)
        try:
            with ZipFile(temporaire, 'w') as sortie:
                for info in archive.infolist():
                    sortie.writestr(info, modifications.get(info.filename, archive.read(info.filename)))
            shutil.copymode(classeur, temporaire)
        except Exception:
            Path(temporaire).unlink(missing_ok=True)
            raise
    try:
        os.replace(temporaire, classeur)
    finally:
        Path(temporaire).unlink(missing_ok=True)
    return True


def actualiser_liens(classeur, bibliotheque=None):
    """Chaque poste ouvre son dossier numéroté, même vide.

    Les calculs, styles, dessins et caches Excel sont conservés.
    """
    from module_engins_cout import (ajouter_module_engins, ajouter_postes_supplementaires,
                                   numeroter_articles_engins,
                                   postes_supplementaires_bibliotheque)
    supplements = postes_supplementaires_bibliotheque(bibliotheque) if bibliotheque is not None else ()
    ajouter_module_engins(classeur)
    numeroter_articles_engins(classeur)
    ajouter_postes_supplementaires(classeur, supplements)
    if supplements:
        libelles = {int(re.match(r'^\d+', p['libelle']).group()): p['libelle']
                    for p in lire_postes(classeur)}
        for numero, libelle in enumerate(supplements, 193):
            attendu = f'{numero:03d} · {libelle}'
            if libelles.get(numero) != attendu:
                raise ValueError(f'Le poste {numero:03d} de la bibliothèque n’apparaît pas correctement dans Excel.')
    classeur = Path(classeur)
    modifications = {}
    with ZipFile(classeur) as z:
        for membre, rp, xml, rels, postes in _lire(z):
            contenu = z.read(membre).decode('utf-8')
            liens = xml.find(f'{{{S}}}hyperlinks')
            if liens is None:
                liens = ET.Element(f'{{{S}}}hyperlinks')
            # Excel attend les relations OPC dans leur espace de noms par défaut.
            # Répare aussi les fichiers déjà produits avec des préfixes ns0.
            change = (rp in z.namelist() and b"<ns0:Relationships" in z.read(rp))
            for poste in postes:
                document = None
                cible = ''
                if poste['texte'] != 'Entreprise':
                    document = dossier_poste(classeur, poste)
                    from liens_horizon import lien_poste
                    cible = lien_poste(classeur.parent, re.match(r'^\d+', poste['libelle']).group())
                if cible == poste['cible'] and (cible or poste['texte'] != 'Docs'):
                    continue
                change = True
                for lien in list(liens):
                    if lien.get('ref') == poste['cellule']:
                        liens.remove(lien)
                if cible:
                    ids = {r.get('Id') for r in rels}
                    i = 1
                    while f'rIdDoc{i}' in ids:
                        i += 1
                    rid = f'rIdDoc{i}'
                    ET.SubElement(rels, f'{{{P}}}Relationship', Id=rid, Type=R + '/hyperlink', Target=cible, TargetMode='External')
                    ET.SubElement(liens, f'{{{S}}}hyperlink', {'ref': poste['cellule'], f'{{{R}}}id': rid, 'tooltip': 'Ouvrir le dossier ' + document.name})
                # Ne remplace jamais les en-têtes ou une saisie métier en colonne B.
                if poste['texte'] in ('', 'Docs'):
                    ref = poste['cellule']
                    motif = rf'<c\b[^>]*\br="{ref}"[^>]*?(?:/>|>.*?</c>)'
                    ancien = re.search(motif, contenu, re.S)
                    if ancien:
                        style = re.search(r'\bs="[^"]*"', ancien.group())
                        attribut = ' ' + style.group() if style else ''
                        nouveau = f'<c r="{ref}"{attribut} t="inlineStr"><is><t>Docs</t></is></c>' if cible else f'<c r="{ref}"{attribut}/>'
                        contenu = contenu[:ancien.start()] + nouveau + contenu[ancien.end():]
                    elif cible:
                        numero = ref[1:]
                        # B suit A dans la ligne du poste.
                        motif_a = rf'(<c\b[^>]*\br="A{numero}"[^>]*?(?:/>|>.*?</c>))'
                        contenu, n = re.subn(motif_a, lambda m: m.group() + f'<c r="{ref}" t="inlineStr"><is><t>Docs</t></is></c>', contenu, count=1, flags=re.S)
                        if n != 1:
                            raise ValueError('Cellule du poste introuvable.')
            if not change:
                continue
            def attributs(element):
                return ' '.join(('r:id' if k == f'{{{R}}}id' else k) + '=' + quoteattr(v)
                                for k, v in element.attrib.items())
            bloc = ('<hyperlinks xmlns="' + S + '" xmlns:r="' + R + '">' +
                    ''.join('<hyperlink ' + attributs(h) + '/>' for h in liens) +
                    '</hyperlinks>') if len(liens) else '' 
            motif = r'<(?:\w+:)?hyperlinks\b[^>]*(?:/>|>.*?</(?:\w+:)?hyperlinks>)'
            if re.search(motif, contenu, re.S):
                contenu = re.sub(motif, lambda _: bloc, contenu, count=1, flags=re.S)
            elif bloc:
                pos = re.search(r'<(?:printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|</worksheet>', contenu).start()
                contenu = contenu[:pos] + bloc + contenu[pos:]
            # Une cellule vide ne doit jamais englober les cellules suivantes.
            protegees = {p['cellule'] for p in postes}
            def cellules_metier(arbre):
                return {c.get('r'): ET.tostring(c) for c in arbre.iter(f'{{{S}}}c')
                        if c.get('r') not in protegees}
            if cellules_metier(xml) != cellules_metier(ET.fromstring(contenu)):
                raise ValueError('Modification des cellules métier détectée : classeur laissé intact.')
            modifications[membre] = contenu.encode('utf-8')
            # Les relations peuvent porter des attributs XML qualifiés.
            # Une concaténation de leurs noms bruts produit alors un XML illisible.
            relations_sortie = deepcopy(rels)
            for relation in relations_sortie.iter():
                if relation.tag.startswith('{' + P + '}'):
                    relation.tag = relation.tag.split('}', 1)[1]
            relations_sortie.set('xmlns', P)
            relations_xml = ET.tostring(relations_sortie, encoding='utf-8', xml_declaration=True)
            ET.fromstring(relations_xml)
            modifications[rp] = relations_xml
        if not modifications:
            return False
        if classeur.with_name('~$' + classeur.name).exists():
            raise PermissionError('Enregistrez et fermez le coût de la sécurité dans Excel, puis recommencez.')
        historique = classeur.parent / 'Historique_Sauvegardes'
        historique.mkdir(exist_ok=True)
        shutil.copy2(classeur, historique / (datetime.now().strftime('%Y%m%d_%H%M%S_%f_') + classeur.stem + '_avant_liens.xlsx'))
        fd, tmp = tempfile.mkstemp(prefix='.documents_', suffix='.xlsx', dir=classeur.parent)
        os.close(fd)
        noms_modifies = set(modifications)
        try:
            with ZipFile(tmp, 'w') as sortie:
                for info in z.infolist():
                    sortie.writestr(info, modifications.pop(info.filename, z.read(info.filename)))
                for nom, data in modifications.items():
                    sortie.writestr(nom, data)
            with ZipFile(tmp) as verification:
                for nom in noms_modifies:
                    ET.fromstring(verification.read(nom))
            shutil.copymode(classeur, tmp)
        except Exception:
            Path(tmp).unlink(missing_ok=True)
            raise
    try:
        os.replace(tmp, classeur)
    finally:
        Path(tmp).unlink(missing_ok=True)
    return True
