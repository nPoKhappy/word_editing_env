"""Idempotent Windows-only DOCX extraction compatibility patch for pinned citecheck."""
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4] / '.tools/endnote'
FILE = ROOT / 'citecheck/node_modules/@jhlee0619/citecheck/src/index.ts'
OLD = '  const { stdout } = await execFileAsync("unzip", ["-p", filePath, "word/document.xml"], {'
NEW = '''  const windowsPython = process.env.CITECHECK_PYTHON;
  const zipCommand = windowsPython || "unzip";
  const zipArgs = windowsPython
    ? ["-c", "import sys,zipfile;sys.stdout.buffer.write(zipfile.ZipFile(sys.argv[1]).read('word/document.xml'))", filePath]
    : ["-p", filePath, "word/document.xml"];
  const { stdout } = await execFileAsync(zipCommand, zipArgs, {'''


def main():
    package = json.loads((FILE.parent.parent / 'package.json').read_text(encoding='utf-8'))
    if package['version'] != '0.1.0':
        raise ValueError('Review the adapter before patching a different citecheck release.')
    source = FILE.read_text(encoding='utf-8')
    if NEW not in source:
        if source.count(OLD) != 1:
            raise ValueError('Pinned DOCX extraction fragment was not found uniquely.')
        FILE.with_suffix('.ts.windows-original').write_text(source, encoding='utf-8')
        FILE.write_text(source.replace(OLD, NEW, 1), encoding='utf-8', newline='\n')
    digest = hashlib.sha256(FILE.read_bytes()).hexdigest()
    print(json.dumps({'ok': True, 'file': str(FILE), 'sha256': digest}))


if __name__ == '__main__':
    main()
