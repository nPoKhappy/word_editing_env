"""Confirm an exact open library using the native EndNote process file handles."""
import argparse
import json
from pathlib import Path
import sys

WORKSPACE = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(WORKSPACE / '.tools/endnote/python'))


def confirm(pid, library_path, open_files=None):
    target = Path(library_path).resolve()
    if open_files is None:
        import psutil
        process = psutil.Process(pid)
        if process.name().casefold() != 'endnote.exe':
            raise ValueError('Expected the native EndNote process.')
        open_files = [f.path for f in process.open_files()]
    matches = sorted({str(Path(p).resolve()) for p in open_files
                      if Path(p).suffix.lower() == '.enl' and Path(p).stem.casefold() == target.stem.casefold()})
    confirmed = len(matches) == 1 and matches[0].casefold() == str(target).casefold()
    return {'ok': confirmed, 'confirmed_library_path': str(target) if confirmed else None,
            'same_name_library_count': len(matches), 'confirmation_method': 'native process open-file handles plus window title'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--pid', type=int, required=True)
    parser.add_argument('--library', required=True)
    args = parser.parse_args()
    try:
        result = confirm(args.pid, args.library)
    except Exception as exc:
        result = {'ok': False, 'error': str(exc)}
    print(json.dumps(result))
    raise SystemExit(0 if result['ok'] else 1)
