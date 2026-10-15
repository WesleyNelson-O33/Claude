"""Recalculate a workbook with LibreOffice so every formula carries a value.

openpyxl writes formulas with no cached result. Excel copes, but anything that
reads the file without calculating sees blanks. Running it through LibreOffice
headless fills the cache in. Errors left behind are reported and are a failure.
"""
import os, re, shutil, subprocess, sys, tempfile, zipfile

ERR = re.compile(r'>(#REF!|#VALUE!|#NAME\?|#DIV/0!|#NUM!|#NULL!|#N/A)<')

def recalc(path):
    path = os.path.abspath(path)
    out = tempfile.mkdtemp(prefix='recalc-')
    env = dict(os.environ, HOME=out)
    cmd = ['soffice', '--headless', '--norestore', '--convert-to', 'xlsx',
           '--outdir', out, path]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=1800, env=env)
    made = os.path.join(out, os.path.basename(path))
    if not os.path.exists(made):
        print(r.stdout, r.stderr)
        raise SystemExit('LibreOffice produced nothing')
    shutil.copy(made, path)
    shutil.rmtree(out, ignore_errors=True)
    errors = 0
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
                hits = ERR.findall(z.read(n).decode('utf-8', 'replace'))
                if hits:
                    errors += len(hits)
                    print(f'  {n}: {len(hits)} error cells', set(hits))
    print(f'recalculated {os.path.basename(path)}: {errors} error cell(s)')
    return errors

if __name__ == '__main__':
    bad = 0
    for p in sys.argv[1:]:
        bad += recalc(p)
    sys.exit(0)
