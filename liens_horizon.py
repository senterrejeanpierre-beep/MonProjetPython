"""Liens natifs vers les listes de documents d'Horizon, sans navigateur."""
import json
import os
from pathlib import Path
import plistlib
import re
import secrets
import shlex
import subprocess
import sys
import tempfile
import time
from urllib.parse import urlsplit

MARQUEUR = '.horizon_documents.json'
SCHEMA = 'horizonchantier'


def dossier_local():
    if sys.platform == 'win32':
        base = Path(os.environ['LOCALAPPDATA'])
    else:
        base = Path.home() / 'Library' / 'Application Support'
    return base / 'HorizonChantier' / 'LiensDocuments'


def identifier(chantier):
    p = Path(chantier) / MARQUEUR
    if p.exists():
        identifiant = json.loads(p.read_text(encoding='utf-8'))['identifiant']
        if not re.fullmatch(r'[A-Za-z0-9_-]{24,64}', identifiant):
            raise ValueError('Identifiant du chantier invalide.')
        return identifiant
    identifiant = secrets.token_urlsafe(24)
    try:
        with p.open('x', encoding='utf-8') as f:
            json.dump({'identifiant': identifiant}, f)
    except FileExistsError:
        return identifier(chantier)
    return identifiant


def lien_poste(chantier, numero):
    return f'{SCHEMA}://documents/{identifier(chantier)}/{int(numero):03d}'


def analyser(lien):
    url = urlsplit(lien)
    correspondance = re.fullmatch(r'/([A-Za-z0-9_-]{24,64})/(\d{3,6})', url.path)
    if url.scheme != SCHEMA or url.netloc != 'documents' or not correspondance or url.query or url.fragment:
        raise ValueError('Lien de document Horizon invalide.')
    return correspondance.group(1), correspondance.group(2)


def resoudre(base, lien):
    identifiant, numero = analyser(lien)
    candidats = []
    for p in Path(base).glob('*/' + MARQUEUR):
        try:
            if json.loads(p.read_text(encoding='utf-8')).get('identifiant') == identifiant:
                candidats.append(p.parent)
        except (ValueError, OSError):
            continue
    if len(candidats) != 1:
        raise ValueError('Le chantier du lien est introuvable ou dupliqué sur ce partage.')
    return candidats[0], numero


def deposer(lien):
    analyser(lien)
    dossier = dossier_local()
    dossier.mkdir(parents=True, exist_ok=True)
    nom = secrets.token_hex(12)
    p = dossier / (nom + '.tmp')
    p.write_text(json.dumps({'lien': lien}), encoding='utf-8')
    p.rename(p.with_suffix('.json'))
    heartbeat = dossier / 'actif'
    return heartbeat.exists() and time.time() - heartbeat.stat().st_mtime < 5


def commandes():
    dossier = dossier_local()
    dossier.mkdir(parents=True, exist_ok=True)
    (dossier / 'actif').touch()
    for p in sorted(dossier.glob('*.json')):
        try:
            yield json.loads(p.read_text(encoding='utf-8'))['lien']
        finally:
            p.unlink(missing_ok=True)


def commande_application(main):
    if getattr(sys, 'frozen', False):
        return [sys.executable, '--document-url']
    executable = sys.executable
    if sys.platform == 'win32':
        pythonw = Path(executable).with_name('pythonw.exe')
        if pythonw.exists():
            executable = str(pythonw)
    return [executable, str(Path(main).resolve()), '--document-url']


def installer(main):
    """Association propre au poste, sans chemin absolu dans les classeurs."""
    commande = commande_application(main)
    local = dossier_local()
    local.mkdir(parents=True, exist_ok=True)
    empreinte = local / 'installation.txt'
    valeur = '3:' + json.dumps(commande)
    if sys.platform == 'win32':
        import winreg
        cle = 'Software\\Classes\\' + SCHEMA
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, cle) as k:
            winreg.SetValueEx(k, '', 0, winreg.REG_SZ, 'URL:Horizon Chantier')
            winreg.SetValueEx(k, 'URL Protocol', 0, winreg.REG_SZ, '')
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, cle + r'\shell\open\command') as k:
            winreg.SetValueEx(k, '', 0, winreg.REG_SZ, subprocess.list2cmdline(commande) + ' "%1"')
    elif sys.platform == 'darwin':
        # La réception d'un lien ne relit pas le projet sur le Bureau.
        # Le relais local transmet uniquement l'identifiant et le poste.
        if not getattr(sys, 'frozen', False):
            relais = local / 'dispatch.py'
            relais.write_text(Path(__file__).read_text(encoding='utf-8'), encoding='utf-8')
            (local / 'application.txt').write_text(json.dumps(commande[:-1]), encoding='utf-8')
            commande = [sys.executable, str(relais)]
        application = Path.home() / 'Applications' / 'Horizon Documents.app'
        if empreinte.exists() and empreinte.read_text() == valeur and application.exists():
            return
        application.parent.mkdir(parents=True, exist_ok=True)
        source = 'on open location adresse\n do shell script ' + json.dumps(shlex.join(commande) + ' ', ensure_ascii=False) + ' & quoted form of adresse & ' + json.dumps(' >' + shlex.quote(str(local / 'lancement.log')) + ' 2>&1 < /dev/null &') + '\nend open location\n'
        with tempfile.TemporaryDirectory() as d:
            script = Path(d) / 'ouvrir.applescript'
            script.write_text(source, encoding='utf-8')
            subprocess.run(['osacompile', '-o', str(application), str(script)], check=True, capture_output=True, timeout=30)
        info = application / 'Contents' / 'Info.plist'
        with info.open('rb') as f:
            plist = plistlib.load(f)
        plist.update(CFBundleIdentifier='be.horizonchantier.documents', LSUIElement=True,
                     CFBundleURLTypes=[{'CFBundleURLName': 'Documents Horizon', 'CFBundleURLSchemes': [SCHEMA]}])
        with info.open('wb') as f:
            plistlib.dump(plist, f)
        subprocess.run(['codesign', '--force', '--sign', '-', str(application)], check=True, capture_output=True, timeout=20)
        subprocess.run(['/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister', '-f', str(application)], check=True, capture_output=True, timeout=20)
    else:
        raise RuntimeError('Ouverture native des documents prévue pour Windows et macOS.')
    empreinte.write_text(valeur, encoding='utf-8')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        raise SystemExit('Un lien Horizon est nécessaire.')
    if not deposer(sys.argv[1]):
        commande = json.loads((dossier_local() / 'application.txt').read_text(encoding='utf-8'))
        subprocess.Popen(commande, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
