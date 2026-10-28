"""Launch the installed search MCP using the workspace Python and isolated packages."""
from pathlib import Path
import os
import sys

WORKSPACE = Path(__file__).resolve().parents[4]
PACKAGES = WORKSPACE / '.tools/endnote/python'
os.environ['PYTHONPATH'] = str(PACKAGES)
os.environ['ENDNOTE_MCP_CONFIG'] = str(WORKSPACE / '.tools/endnote/state/library.yaml')
os.environ.setdefault('HF_HOME', str(WORKSPACE / '.tools/endnote/models'))
os.environ['HF_HUB_DISABLE_TELEMETRY'] = '1'
os.environ['DO_NOT_TRACK'] = '1'
os.environ['TOKENIZERS_PARALLELISM'] = 'false'
os.environ['PYTHONIOENCODING'] = 'utf-8'
os.environ['TEMP'] = os.environ['TMP'] = str(WORKSPACE / '.tmp')


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else 'serve'
    if mode == 'refresh':
        # Load local adapters before importing the third-party package with the same name.
        import argparse
        import endnote_integrations as adapter
        parser = argparse.ArgumentParser()
        parser.add_argument('--library-path', required=True)
        parser.add_argument('--semantic', action='store_true')
        args = parser.parse_args(sys.argv[2:])
        info = adapter.create_search_snapshot(args.library_path)
        print('Search-only snapshot:', info['record_count'], 'references', flush=True)
        sys.modules.pop('endnote_mcp', None)
        sys.path.insert(0, str(PACKAGES))
        # Keep per-record subfolders: upstream otherwise discards them, making
        # different attachments named "paper.pdf" indistinguishable.
        import endnote_mcp.endnote_parser as parser_module
        def pdf_reference(record):
            for element in record.findall('.//urls/pdf-urls/url'):
                url = ''.join(element.itertext()).strip()
                if url.startswith('internal-pdf://'):
                    relative = url[len('internal-pdf://'):]
                    parts = Path(relative.replace('\\', '/'))
                    if parts.is_absolute() or '..' in parts.parts:
                        raise ValueError('Unsafe internal attachment path.')
                    return relative
            return ''
        parser_module._extract_pdf_filename = pdf_reference
        from endnote_mcp.cli import _run_index, _run_embed
        _run_index(os.environ['ENDNOTE_MCP_CONFIG'], full=True, skip_pdfs=False)
        if args.semantic:
            _run_embed(os.environ['ENDNOTE_MCP_CONFIG'], full=False)
        import json
        info['indexing_complete'] = True
        info['semantic_requested'] = args.semantic
        (adapter.STATE / 'snapshot.json').write_text(json.dumps(info, indent=2), encoding='utf-8')
        return
    sys.path.insert(0, str(PACKAGES))
    if mode == 'serve':
        # Downloads happen during explicit setup/refresh; ordinary library search is local.
        os.environ['HF_HUB_OFFLINE'] = '1'
        from endnote_mcp.server import mcp
        mcp.run()
    elif mode == 'status':
        from endnote_mcp.cli import cli
        cli(['status'])
    elif mode == 'embed':
        from endnote_mcp.cli import _run_embed
        _run_embed(os.environ['ENDNOTE_MCP_CONFIG'])
    else:
        raise ValueError('Unknown operation: ' + mode)


if __name__ == '__main__':
    main()
