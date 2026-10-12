"""Zones de texte Excel dans les barres, dans la même ligne de poste."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
import os
import posixpath

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

X = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'


def texte_legende(label):
    if not isinstance(label, str):
        return ''
    return label


def _forme(parent, identifiant, nom, gauche, haut, largeur, hauteur, texte, fond, font):
    def sous(noeud, espace, balise, **attributs):
        return ET.SubElement(noeud, f'{{{espace}}}{balise}', {k:str(v) for k,v in attributs.items()})
    anchor = sous(parent, X, 'absoluteAnchor')
    sous(anchor, X, 'pos', x=round(gauche*12700), y=round(haut*12700))
    sous(anchor, X, 'ext', cx=max(0,round(largeur*12700)), cy=max(0,round(hauteur*12700)))
    forme = sous(anchor, X, 'sp')
    nv = sous(forme, X, 'nvSpPr')
    sous(nv, X, 'cNvPr', id=identifiant, name=nom)
    sous(nv, X, 'cNvSpPr', txBox=1)
    props = sous(forme, X, 'spPr')
    trans = sous(props, A, 'xfrm')
    sous(trans, A, 'off', x=round(gauche*12700), y=round(haut*12700))
    sous(trans, A, 'ext', cx=max(0,round(largeur*12700)), cy=max(0,round(hauteur*12700)))
    geo = sous(props, A, 'prstGeom', prst='rect')
    sous(geo, A, 'avLst')
    if fond:
        sous(sous(props, A, 'solidFill'), A, 'srgbClr', val=fond)
    else:
        sous(props, A, 'noFill')
    sous(sous(props, A, 'ln'), A, 'noFill')
    body = sous(forme, X, 'txBody')
    sous(body, A, 'bodyPr', wrap='none', lIns=0, rIns=0, tIns=0, bIns=0, anchor='ctr',
         vertOverflow='clip', horzOverflow='clip')
    sous(body, A, 'lstStyle')
    p = sous(body, A, 'p')
    sous(p, A, 'pPr', algn='l')
    run = sous(p, A, 'r')
    rp = sous(run, A, 'rPr', lang='fr-BE', sz=round((font.sz or 10)*100))
    sous(sous(rp, A, 'solidFill'), A, 'srgbClr', val='163A59')
    sous(rp, A, 'latin', typeface=font.name or 'Helvetica')
    sous(run, A, 't').text = texte
    sous(anchor, X, 'clientData', fPrintsWithSheet=1)


def ecrire_legendes(chemin):
    """Conserve le dessin et le logo existants, puis ajoute les seules zones Horizon."""
    chemin = Path(chemin)
    with ZipFile(chemin) as z:
        fichiers = {n:z.read(n) for n in z.namelist()}
    w = load_workbook(chemin)
    try:
        s = w['Planning soumission']
        nombre, fin = int(s.cell(1,159).value), int(s.cell(2,159).value)
        feuille = ET.fromstring(fichiers['xl/worksheets/sheet1.xml'])
        liens_nom = 'xl/worksheets/_rels/sheet1.xml.rels'
        liens = ET.fromstring(fichiers[liens_nom]) if liens_nom in fichiers else ET.Element(f'{{{P}}}Relationships')
        drawing = feuille.find(f'{{{S}}}drawing')
        nouveau_dessin = drawing is None
        if drawing is None:
            identifiant = 'rIdLegendesHorizon'
            drawing = ET.SubElement(feuille, f'{{{S}}}drawing', {f'{{{R}}}id':identifiant})
            chemin_dessin = 'xl/drawings/legendesHorizon.xml'
            ET.SubElement(liens, f'{{{P}}}Relationship', {'Id':identifiant, 'Type':R+'/drawing', 'Target':'/xl/drawings/legendesHorizon.xml'})
            root = ET.Element(f'{{{X}}}wsDr')
        else:
            rel = next(v for v in liens if v.get('Id') == drawing.get(f'{{{R}}}id'))
            cible = rel.get('Target')
            chemin_dessin = cible.lstrip('/') if cible.startswith('/') else posixpath.normpath('xl/worksheets/'+cible)
            root = ET.fromstring(fichiers[chemin_dessin])
        for anchor in list(root):
            nv = anchor.find(f'.//{{{X}}}cNvPr')
            if nv is not None and nv.get('name','').startswith(
                    ('HorizonLegende_', 'HorizonFond_', 'HorizonGuide_', 'HorizonCadre_')):
                root.remove(anchor)
        identifiant = max((int(v.get('id','0')) for v in root.iter(f'{{{X}}}cNvPr')),default=0)+1
        # Les positions sont recalées par Excel à l'ouverture avec ses mesures exactes.
        largeurs = {c: (s.column_dimensions[get_column_letter(c)].width*7+5)*.75 for c in range(1,147)}
        gauches = {c:sum(largeurs[i] for i in range(1,c)) for c in range(1,148)}
        hauts = {r:sum(s.row_dimensions[i].height or s.sheet_format.defaultRowHeight or 15 for i in range(1,r)) for r in range(9,fin+1)}
        jour = s['G4'].value if isinstance(s['G4'].value,(int,float)) else 1
        for r in range(9,fin+1):
            hauteur = s.row_dimensions[r].height or 32
            font = s.cell(r,2).font
            infos = []
            for n in range(nombre):
                texte = texte_legende(s.cell(r,160+2*nombre+n).value)
                debut, dernier = s.cell(r,160+2*n).value, s.cell(r,161+2*n).value
                indices = [c for c in range(7,147) if isinstance(debut,(int,float)) and isinstance(dernier,(int,float))
                           and debut >= 1 and dernier > debut-1 and debut <= jour+c-7 < dernier+1]
                infos.append((texte, indices))
            for n,(texte,indices) in enumerate(infos,1):
                gauche, largeur, haut = gauches[7], .1, .1
                if texte and indices:
                    gauche = gauches[indices[0]]+1
                    limite = min([gauches[i[0]]-1 for t,i in infos if t and i and i[0]>indices[0]]+[gauches[147]])
                    largeur = max(.1,min(limite-gauche,max(gauches[indices[-1]+1]-gauche,len(texte)*(font.sz or 10)*.55+4)))
                    haut = hauteur-2
                _forme(root,identifiant,f'HorizonLegende_{r}_{n}',gauche,hauts[r]+1,largeur,haut,texte,None,font)
                identifiant += 1
        ET.register_namespace('xdr', X)
        ET.register_namespace('a', A)
        ET.register_namespace('r', R)
        fichiers[chemin_dessin] = ET.tostring(root,encoding='utf-8',xml_declaration=True)
        if nouveau_dessin:
            ET.register_namespace('', S)
            fichiers['xl/worksheets/sheet1.xml'] = ET.tostring(feuille,encoding='utf-8',xml_declaration=True)
            fichiers[liens_nom] = ET.tostring(liens,encoding='utf-8',xml_declaration=True)
        types = ET.fromstring(fichiers['[Content_Types].xml'])
        if not any(v.get('PartName') == '/'+chemin_dessin for v in types):
            ET.SubElement(types,f'{{{C}}}Override',{'PartName':'/'+chemin_dessin,'ContentType':'application/vnd.openxmlformats-officedocument.drawing+xml'})
            fichiers['[Content_Types].xml'] = ET.tostring(types,encoding='utf-8',xml_declaration=True)
    finally:
        w.close()
    with TemporaryDirectory(prefix='.legendes_',dir=chemin.parent) as temporaire:
        nouveau = Path(temporaire)/chemin.name
        with ZipFile(nouveau,'w',ZIP_DEFLATED) as z:
            for nom,data in fichiers.items():
                z.writestr(nom,data)
        os.replace(nouveau,chemin)
