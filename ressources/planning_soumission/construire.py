"""Reconstruit vbaProject.bin ; outil de développement pyOpenVBA==6.3.0."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from pyopenvba import ExcelFile

base = Path(__file__).resolve().parent
with TemporaryDirectory(prefix='.construction_', dir=base) as temporaire:
    with ExcelFile.create_new(Path(temporaire)/'planning.xlsm') as classeur:
        for module, fichier in [('ThisWorkbook', 'ThisWorkbook.cls'), ('Module1', 'Periodes.bas')]:
            attributs = '\r\n'.join(l for l in classeur.get_module(module).splitlines()
                                    if l.startswith('Attribute ')) + '\r\n'
            classeur.set_module(module, attributs + base.joinpath(fichier).read_text().replace('\n', '\r\n'))
        classeur.save()
    with ZipFile(Path(temporaire)/'planning.xlsm') as archive:
        base.joinpath('vbaProject.bin').write_bytes(archive.read('xl/vbaProject.bin'))
