"""Ajout ciblé de la rubrique engins, sans réécriture globale du classeur."""
from copy import deepcopy
from datetime import datetime
import os
from pathlib import Path
import re
import shutil
import tempfile
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from documents_cout import S, _feuilles

TITRE = '11 — CONTAINERS, ENGINS DE LEVAGE ET VÉGÉTATION'
POSTES = (
    'Containers', 'Engins de levage', 'Mini-pelle',
    'Chariot télescopique', 'Petit tracteur', 'Broyeur de végétaux',
    'Dumper', 'Pelle mécanique', 'Chargeuse', 'Mini-chargeuse',
    'Tractopelle', 'Grue mobile', 'Camion-grue',
    'Grutier — intervention toiture', 'Chariot élévateur',
    'Rouleau compacteur', 'Dessoucheuse', 'Débroussailleuse',
    'Transport et reprise des engins', 'Grue à tour',
)
FIN = 175 + len(POSTES)
TOTAL_INITIAL = 'SUM(K19,K37,K60,K84,K94,K116,K126,K136,K146,K163)'


def postes_supplementaires_bibliotheque(bibliotheque):
    """Lit les nouveaux postes consécutifs à partir de 193 dans la bibliothèque."""
    bibliotheque = Path(bibliotheque)
    if not bibliotheque.is_dir():
        return ()
    numeros = {}
    for dossier in bibliotheque.iterdir():
        if dossier.is_symlink() or not dossier.is_dir():
            continue
        correspondance = re.fullmatch(r'(\d{3,})_(.+)', dossier.name)
        if correspondance is None:
            continue
        numero = int(correspondance.group(1))
        if numero < 193:
            continue
        if numero in numeros:
            raise ValueError(f'Plusieurs dossiers de bibliothèque portent le numéro {numero:03d}.')
        numeros[numero] = correspondance.group(2).replace('_', ' ')
    if numeros and sorted(numeros) != list(range(193, max(numeros) + 1)):
        raise ValueError('Les nouveaux numéros de sécurité doivent se suivre à partir de 193.')
    return tuple(numeros[n] for n in sorted(numeros))


def _serie(element):
    element = deepcopy(element)
    for e in element.iter():
        if e.tag.startswith('{' + S + '}'):
            e.tag = e.tag.split('}', 1)[1]
    return ET.tostring(element, encoding='unicode')


def _texte(cellule, valeur):
    for enfant in list(cellule):
        cellule.remove(enfant)
    cellule.set('t', 'inlineStr')
    ET.SubElement(ET.SubElement(cellule, 'is'), 't').text = valeur


def _formule(cellule, formule, cache=None):
    cellule.attrib.pop('t', None)
    for enfant in list(cellule):
        cellule.remove(enfant)
    ET.SubElement(cellule, 'f').text = formule
    if cache is not None:
        ET.SubElement(cellule, 'v').text = cache


def _remplacer_cellule(xml, cellule):
    ref = cellule.get('r')
    motif = rf'<c\b[^>]*\br="{ref}"[^>]*?(?:/>|>.*?</c>)'
    xml, n = re.subn(motif, lambda _: _serie(cellule), xml, count=1, flags=re.S)
    if n != 1:
        raise ValueError(f'Cellule {ref} introuvable : ajout du module annulé.')
    return xml


def ajouter_module_engins(classeur):
    """Étend uniquement les feuilles de calcul reconnues, pas les copies figées."""
    classeur = Path(classeur)
    modifications = {}
    with ZipFile(classeur) as archive:
        for _, membre in _feuilles(archive):
            original = archive.read(membre).decode('utf-8')
            racine = ET.fromstring(original)
            cellules = {c.get('r'): c for c in racine.iter(f'{{{S}}}c')}
            total = cellules.get('K17')
            if total is None or total.find(f'{{{S}}}f') is None:
                continue
            # Étendre un module installé sans toucher aux postes déjà saisis.
            if total.find(f'{{{S}}}f').text == TOTAL_INITIAL[:-1] + ',K175)':
                somme = cellules.get('K175')
                formule = somme.find(f'{{{S}}}f') if somme is not None else None
                plage = re.fullmatch(r'SUM\(\$K176:\$K(\d+)(?:,K\d+)*\)', formule.text or '') if formule is not None else None
                if plage is None:
                    raise ValueError('Sous-total engins non reconnu.')
                fin_avant = int(plage.group(1))
                if fin_avant >= FIN:
                    continue
                lignes = {int(r.get('r')): r for r in racine.findall(f'{{{S}}}sheetData/{{{S}}}row')}
                if any(n in lignes for n in range(fin_avant + 1, FIN + 1)):
                    raise ValueError('Les lignes suivant le module sont occupées.')
                cloture = cellules['K163']
                if cloture.find(f'{{{S}}}f').text != f'SUM($K164:$K173,$K{fin_avant + 1}:$K524)':
                    raise ValueError('Somme de clôture différente : extension annulée.')
                formule_source = cellules['K164'].find(f'{{{S}}}f').text
                nouveaux = []
                for numero in range(fin_avant + 1, FIN + 1):
                    ligne = deepcopy(lignes[164])
                    ligne.set('r', str(numero))
                    for c in ligne:
                        col = re.sub(r'\d+$', '', c.get('r'))
                        c.set('r', col + str(numero))
                        for enfant in list(c):
                            c.remove(enfant)
                        c.attrib.pop('t', None)
                        if col == 'A':
                            _texte(c, f'{numero - 3:03d} · {POSTES[numero - 176]}')
                        elif col == 'B':
                            _texte(c, 'Docs')
                        elif col == 'K':
                            _formule(c, re.sub(r'(?<=[A-Z])164\b', str(numero), formule_source))
                            c.set('t', 'str')
                            ET.SubElement(c, 'v')
                    nouveaux.append(ligne)
                suivant = min((n for n in lignes if n > FIN), default=None)
                marqueur = re.search(rf'<row\b[^>]*\br="{suivant}"', original) if suivant else re.search(r'</sheetData>', original)
                xml = original[:marqueur.start()] + ''.join(_serie(r) for r in nouveaux) + original[marqueur.start():]
                somme = deepcopy(somme)
                somme.find(f'{{{S}}}f').text = f'SUM($K176:$K{FIN})'
                cloture = deepcopy(cloture)
                cloture.find(f'{{{S}}}f').text = f'SUM($K164:$K173,$K{FIN + 1}:$K524)'
                xml = _remplacer_cellule(xml, somme)
                xml = _remplacer_cellule(xml, cloture)
                apres = {c.get('r'): c for c in ET.fromstring(xml).iter(f'{{{S}}}c')}
                assert all(ET.tostring(c) == ET.tostring(apres[ref]) for ref, c in cellules.items() if ref not in ('K175', 'K163'))
                modifications[membre] = xml.encode('utf-8')
                continue
            if total.find(f'{{{S}}}f').text != TOTAL_INITIAL:
                continue  # Autre organisation : aucune règle de calcul déduite.
            cloture = cellules.get('K163')
            if cloture is None or cloture.find(f'{{{S}}}f') is None or cloture.find(f'{{{S}}}f').text != 'SUM($K164:$K524)':
                raise ValueError('Somme de clôture différente : ajout du module annulé.')
            lignes = {int(r.get('r')): r for r in racine.findall(f'{{{S}}}sheetData/{{{S}}}row')}
            for numero in range(174, FIN + 1):
                if numero in lignes:
                    raise ValueError(f'La ligne {numero} est déjà occupée : ajout du module annulé.')
            if any(n not in lignes for n in (162, 163, 164)):
                raise ValueError('Présentation de référence introuvable.')

            def copier_ligne(source, destination):
                ligne = deepcopy(lignes[source])
                ligne.set('r', str(destination))
                for c in ligne:
                    c.set('r', re.sub(r'\d+$', str(destination), c.get('r')))
                return ligne

            entete = copier_ligne(162, 174)
            _texte(entete[0], 'Définition')
            rubrique = copier_ligne(163, 175)
            _texte(rubrique[0], TITRE)
            _formule(next(c for c in rubrique if c.get('r') == 'K175'), f'SUM($K176:$K{FIN})', '0')
            nouveaux = [entete, rubrique]
            formule_source = cellules['K164'].find(f'{{{S}}}f').text
            if not formule_source:
                raise ValueError('Formule de poste introuvable.')
            for numero, libelle in enumerate(POSTES, 176):
                ligne = copier_ligne(164, numero)
                for c in ligne:
                    col = re.sub(r'\d+$', '', c.get('r'))
                    for enfant in list(c):
                        c.remove(enfant)
                    c.attrib.pop('t', None)
                    if col == 'A':
                        _texte(c, f'{numero - 3:03d} · {libelle}')
                    elif col == 'B':
                        _texte(c, 'Docs')
                    elif col == 'K':
                        # Même multiplication des quantités que les postes existants.
                        _formule(c, re.sub(r'(?<=[A-Z])164\b', str(numero), formule_source))
                        c.set('t', 'str')
                        ET.SubElement(c, 'v')
                nouveaux.append(ligne)
            suivant = min((n for n in lignes if n > FIN), default=None)
            marqueur = re.search(rf'<row\b[^>]*\br="{suivant}"', original) if suivant else re.search(r'</sheetData>', original)
            if marqueur is None:
                raise ValueError('Fin de tableau introuvable.')
            xml = original[:marqueur.start()] + ''.join(_serie(r) for r in nouveaux) + original[marqueur.start():]
            total = deepcopy(total)
            total.find(f'{{{S}}}f').text = TOTAL_INITIAL[:-1] + ',K175)'
            cloture = deepcopy(cloture)
            cloture.find(f'{{{S}}}f').text = f'SUM($K164:$K173,$K{FIN + 1}:$K524)'
            xml = _remplacer_cellule(xml, total)
            xml = _remplacer_cellule(xml, cloture)
            if 'A161' in cellules:
                note = deepcopy(cellules['A161'])
                _texte(note, 'Containers et engins : voir rubrique 11.')
                xml = _remplacer_cellule(xml, note)
            apres = {c.get('r'): c for c in ET.fromstring(xml).iter(f'{{{S}}}c')}
            for ref, avant in cellules.items():
                if ref not in ('K17', 'K163', 'A161') and ET.tostring(avant) != ET.tostring(apres.get(ref)):
                    raise ValueError(f'Cellule existante modifiée ({ref}) : ajout annulé.')
            modifications[membre] = xml.encode('utf-8')
        if not modifications:
            return False
        # Excel reconstruira les dépendances des nouvelles formules à l'ouverture.
        supprimes = set()
        if 'xl/calcChain.xml' in archive.namelist():
            supprimes.add('xl/calcChain.xml')
            for nom, motif in (
                ('xl/_rels/workbook.xml.rels', r'<(?:\w+:)?Relationship\b[^>]*Type="[^"]*/calcChain"[^>]*/>'),
                ('[Content_Types].xml', r'<(?:\w+:)?Override\b[^>]*PartName="/xl/calcChain.xml"[^>]*/>'),
            ):
                contenu = archive.read(nom).decode('utf-8')
                contenu, n = re.subn(motif, '', contenu)
                if n != 1:
                    raise ValueError('Référence de calcul non reconnue : ajout annulé.')
                modifications[nom] = contenu.encode('utf-8')
        workbook = archive.read('xl/workbook.xml').decode('utf-8')
        calc = re.search(r'<calcPr\b[^>]*/>', workbook)
        if calc:
            element = ET.fromstring(calc.group())
            element.set('fullCalcOnLoad', '1')
            element.set('forceFullCalc', '1')
            workbook = workbook[:calc.start()] + ET.tostring(element, encoding='unicode') + workbook[calc.end():]
        else:
            workbook = workbook.replace('</workbook>', '<calcPr fullCalcOnLoad="1" forceFullCalc="1"/></workbook>')
        modifications['xl/workbook.xml'] = workbook.encode('utf-8')
        if classeur.with_name('~$' + classeur.name).exists():
            raise PermissionError('Enregistrez et fermez le coût de la sécurité dans Excel pour ajouter le module engins.')
        historique = classeur.parent / 'Historique_Sauvegardes'
        historique.mkdir(exist_ok=True)
        shutil.copy2(classeur, historique / (datetime.now().strftime('%Y%m%d_%H%M%S_%f_') + 'Cout_avant_module_engins.xlsx'))
        fd, temporaire = tempfile.mkstemp(prefix='.engins_', suffix='.xlsx', dir=classeur.parent)
        os.close(fd)
        try:
            with ZipFile(temporaire, 'w') as sortie:
                for info in archive.infolist():
                    if info.filename in supprimes:
                        continue
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


def ajouter_postes_supplementaires(classeur, libelles):
    """Ajoute les postes après les lignes occupées, sans déplacer la clôture."""
    from documents_cout import _lire
    libelles = tuple(libelles)
    if not libelles:
        return False
    classeur = Path(classeur)
    modifications = {}
    with ZipFile(classeur) as archive:
        for membre, _, arbre, _, postes in _lire(archive):
            cellules = {c.get('r'): c for c in arbre.iter(f'{{{S}}}c')}
            total = cellules.get('K17')
            formule_total = total.find(f'{{{S}}}f') if total is not None else None
            # Le total peut avoir été réenregistré avec des espaces ou des
            # références absolues. Seule la présence du sous-total K175 compte ici.
            if (formule_total is None or not re.search(
                    r'(?<![A-Z0-9])\$?K\$?175(?!\d)', formule_total.text or '', re.I)):
                continue
            somme = cellules.get('K175')
            formule_somme = somme.find(f'{{{S}}}f') if somme is not None else None
            formule = re.sub(r'\s+', '', formule_somme.text or '') if formule_somme is not None else ''
            match = re.fullmatch(r'SUM\(\$K176:\$K(\d+)((?:,K\d+)*)\)', formule)
            if match is None or int(match.group(1)) < FIN:
                raise ValueError('Sous-total de sécurité non reconnu : ajout des postes annulé.')
            lignes = {int(r.get('r')): r for r in arbre.findall(f'{{{S}}}sheetData/{{{S}}}row')}
            source = lignes.get(195)
            formule_source = cellules.get('K164')
            formule_source = formule_source.find(f'{{{S}}}f') if formule_source is not None else None
            if source is None or formule_source is None or not formule_source.text:
                raise ValueError('Ligne de poste de référence introuvable.')
            existants = {}
            for poste in postes:
                numero = re.match(r'^(\d{3,})\s*[·.\-]', poste['libelle'])
                if numero is not None and int(numero.group(1)) >= 193:
                    n = int(numero.group(1))
                    if n in existants:
                        raise ValueError(f'Plusieurs postes Excel portent le numéro {n:03d}.')
                    existants[n] = (poste['libelle'], int(poste['cellule'][1:]))
            attendus = {n: f'{n:03d} · {libelle}' for n, libelle in enumerate(libelles, 193)}
            renommages = {}
            sommes = set(range(176, int(match.group(1)) + 1))
            sommes.update(int(ref[1:]) for ref in match.group(2).split(',') if ref)
            refs_manquantes = []
            for n, (libelle, ligne) in existants.items():
                if n not in attendus:
                    raise ValueError(f'Le poste Excel {n:03d} est absent de la bibliothèque.')
                if libelle != attendus[n]:
                    renommages[f'A{ligne}'] = attendus[n]
                if ligne not in sommes:
                    refs_manquantes.append(f'K{ligne}')
            nouveaux = [(n, libelle) for n, libelle in attendus.items() if n not in existants]
            if not nouveaux and not renommages and not refs_manquantes:
                continue
            if [n for n, _ in nouveaux] != list(range(193 + len(existants), 193 + len(libelles))):
                raise ValueError('Les postes Excel existants ne forment pas une suite continue.')
            debut = max(524, *lignes) + 1
            ajouts = []
            refs = list(refs_manquantes)
            for decalage, (numero, libelle) in enumerate(nouveaux):
                ligne_numero = debut + decalage
                ligne = deepcopy(source)
                ligne.set('r', str(ligne_numero))
                for cellule in ligne:
                    col = re.sub(r'\d+$', '', cellule.get('r'))
                    cellule.set('r', col + str(ligne_numero))
                    for enfant in list(cellule):
                        cellule.remove(enfant)
                    cellule.attrib.pop('t', None)
                    if col == 'A':
                        _texte(cellule, libelle)
                    elif col == 'B':
                        _texte(cellule, 'Docs')
                    elif col == 'K':
                        _formule(cellule, re.sub(r'(?<=[A-Z])164\b', str(ligne_numero),
                                                   formule_source.text))
                        cellule.set('t', 'str')
                        ET.SubElement(cellule, 'v')
                ajouts.append(_serie(ligne))
                refs.append(f'K{ligne_numero}')
            xml = archive.read(membre).decode('utf-8')
            if '</sheetData>' not in xml:
                raise ValueError('Fin du tableau Excel introuvable : ajout annulé.')
            xml = xml.replace('</sheetData>', ''.join(ajouts) + '</sheetData>', 1)
            if refs:
                nouveau_total = deepcopy(somme)
                _formule(nouveau_total, formule[:-1] + ',' + ','.join(refs) + ')')
                xml = _remplacer_cellule(xml, nouveau_total)
            for ref, libelle in renommages.items():
                cellule = deepcopy(cellules[ref])
                _texte(cellule, libelle)
                xml = _remplacer_cellule(xml, cellule)
            dimension = re.search(r'<dimension\b[^>]*\bref="([A-Z]+\d+:?[A-Z]*)(\d+)"[^>]*/>', xml)
            if nouveaux and dimension and int(dimension.group(2)) < debut + len(nouveaux) - 1:
                xml = xml[:dimension.start(2)] + str(debut + len(nouveaux) - 1) + xml[dimension.end(2):]
            apres = {c.get('r'): c for c in ET.fromstring(xml).iter(f'{{{S}}}c')}
            modifiees = set(renommages) | ({'K175'} if refs else set())
            if any(ET.tostring(c) != ET.tostring(apres.get(ref)) for ref, c in cellules.items()
                   if ref not in modifiees):
                raise ValueError('Une cellule existante serait modifiée : ajout annulé.')
            if any(f'A{debut + index}' not in apres for index in range(len(nouveaux))):
                raise ValueError('Nouvelles lignes Excel introuvables : ajout annulé.')
            modifications[membre] = xml.encode('utf-8')
        if not modifications:
            return False
        supprimes = set()
        if 'xl/calcChain.xml' in archive.namelist():
            supprimes.add('xl/calcChain.xml')
            for nom, motif in (
                ('xl/_rels/workbook.xml.rels', r'<(?:\w+:)?Relationship\b[^>]*Type="[^"]*/calcChain"[^>]*/>'),
                ('[Content_Types].xml', r'<(?:\w+:)?Override\b[^>]*PartName="/xl/calcChain.xml"[^>]*/>'),
            ):
                contenu, n = re.subn(motif, '', archive.read(nom).decode('utf-8'))
                if n != 1:
                    raise ValueError('Référence de calcul non reconnue : ajout annulé.')
                modifications[nom] = contenu.encode('utf-8')
        workbook = archive.read('xl/workbook.xml').decode('utf-8')
        calc = re.search(r'<calcPr\b[^>]*/>', workbook)
        if calc:
            element = ET.fromstring(calc.group())
            element.set('fullCalcOnLoad', '1')
            element.set('forceFullCalc', '1')
            workbook = workbook[:calc.start()] + ET.tostring(element, encoding='unicode') + workbook[calc.end():]
        else:
            workbook = workbook.replace('</workbook>', '<calcPr fullCalcOnLoad="1" forceFullCalc="1"/></workbook>')
        modifications['xl/workbook.xml'] = workbook.encode('utf-8')
        if classeur.with_name('~$' + classeur.name).exists():
            raise PermissionError('Fermez le coût de la sécurité dans Excel pour ajouter les postes.')
        historique = classeur.parent / 'Historique_Sauvegardes'
        historique.mkdir(exist_ok=True)
        shutil.copy2(classeur, historique / (datetime.now().strftime('%Y%m%d_%H%M%S_%f_') + 'Cout_avant_postes_supplementaires.xlsx'))
        fd, temporaire = tempfile.mkstemp(prefix='.postes_securite_', suffix='.xlsx', dir=classeur.parent)
        os.close(fd)
        try:
            with ZipFile(temporaire, 'w') as sortie:
                for info in archive.infolist():
                    if info.filename not in supprimes:
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


def numeroter_articles_engins(classeur):
    """Migration 176–195 vers 173–192 ; les lignes Excel ne bougent pas."""
    from urllib.parse import quote, unquote
    from documents_cout import _lire
    classeur = Path(classeur)
    modifications = {}
    deplacements = {}
    with ZipFile(classeur) as archive:
        for membre, rp, xml, rels, postes in _lire(archive):
            cellules = {c.get('r'): c for c in xml.iter(f'{{{S}}}c')}
            formule = cellules.get('K175')
            if formule is None or formule.find(f'{{{S}}}f') is None:
                continue
            contenu = archive.read(membre).decode('utf-8')
            change = False
            for poste in postes:
                ligne = int(poste['cellule'][1:])
                if not 176 <= ligne <= FIN:
                    continue
                libelle = poste['libelle']
                numero = re.match(r'^(\d{3,})\s*([·.\-])\s*(.*)$', libelle)
                if numero is None:
                    raise ValueError(f'Numéro du poste à la ligne {ligne} non reconnu : renumérotation annulée.')
                ancien, separateur, description = numero.groups()
                if int(ancien) == ligne - 3:
                    continue
                if int(ancien) != ligne:
                    raise ValueError(f'Numéro du poste à la ligne {ligne} non reconnu : renumérotation annulée.')
                cellule = deepcopy(cellules[f'A{ligne}'])
                _texte(cellule, f'{ligne - 3:03d} {separateur} {description}')
                contenu = _remplacer_cellule(contenu, cellule)
                change = True
                base = classeur.parent / 'Documents_installation'
                sources = list(base.glob(f'{ligne:03d}_*')) if base.exists() else []
                if len(sources) > 1:
                    raise ValueError(f'Plusieurs dossiers pour le poste {ligne}.')
                for source in sources:
                    if source.is_symlink() or not source.is_dir():
                        raise ValueError('Dossier de documents non standard.')
                    destination = source.with_name(f'{ligne - 3:03d}' + source.name[3:])
                    if destination.exists():
                        raise FileExistsError(f'Dossier déjà existant : {destination.name}')
                    deplacements[source] = destination
            if change:
                modifications[membre] = contenu.encode('utf-8')
        if not modifications:
            return False
        # Mettre à jour les cibles dans la même opération que les libellés.
        from xml.sax.saxutils import quoteattr
        for nom in archive.namelist():
            if not nom.startswith('xl/worksheets/_rels/') or not nom.endswith('.rels'):
                continue
            texte = archive.read(nom).decode('utf-8')
            for source, destination in deplacements.items():
                ancien = source.relative_to(classeur.parent).as_posix()
                nouveau = destination.relative_to(classeur.parent).as_posix()
                def remplacer(m):
                    from html import unescape
                    cible = unquote(unescape(m.group(1))).replace('\\', '/')
                    if cible.rstrip('/') == ancien:
                        return 'Target=' + quoteattr(quote(nouveau, safe='/') + '/')
                    return m.group()
                texte = re.sub(r'Target="([^"]*)"', remplacer, texte)
            if texte.encode('utf-8') != archive.read(nom):
                modifications[nom] = texte.encode('utf-8')
        if classeur.with_name('~$' + classeur.name).exists():
            raise PermissionError('Fermez le coût de la sécurité dans Excel pour renuméroter les articles.')
        historique = classeur.parent / 'Historique_Sauvegardes'
        historique.mkdir(exist_ok=True)
        shutil.copy2(classeur, historique / (datetime.now().strftime('%Y%m%d_%H%M%S_%f_') + 'Cout_avant_numerotation_articles.xlsx'))
        fd, temporaire = tempfile.mkstemp(prefix='.articles_', suffix='.xlsx', dir=classeur.parent)
        os.close(fd)
        try:
            with ZipFile(temporaire, 'w') as sortie:
                for info in archive.infolist():
                    sortie.writestr(info, modifications.get(info.filename, archive.read(info.filename)))
            shutil.copymode(classeur, temporaire)
        except Exception:
            Path(temporaire).unlink(missing_ok=True)
            raise
    faits = []
    try:
        for source, destination in deplacements.items():
            source.rename(destination)
            faits.append((source, destination))
        os.replace(temporaire, classeur)
    except Exception:
        for source, destination in reversed(faits):
            destination.rename(source)
        raise
    finally:
        Path(temporaire).unlink(missing_ok=True)
    return True
