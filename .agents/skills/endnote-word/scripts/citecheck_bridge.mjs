// Thin one-shot adapter: all bibliographic matching is performed by citecheck.
import {readFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
const workspace = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../../..');
const api = await import(pathToFileURL(path.join(workspace,
  '.tools/endnote/citecheck/node_modules/@jhlee0619/citecheck/src/index.ts')).href);
try {
  for (const key of ['REFFORGE_HTTP_MAX_RETRIES', 'REFFORGE_PUBMED_RETRIES',
    'REFFORGE_CROSSREF_RETRIES', 'REFFORGE_ARXIV_RETRIES', 'REFFORGE_SEMANTIC_SCHOLAR_RETRIES'])
    process.env[key] ??= '0';
  process.env.REFFORGE_HTTP_TIMEOUT_MS ??= '7000';
  process.env.REFFORGE_ENABLE_SEMANTIC_SCHOLAR ??= 'true';
  const args = JSON.parse(await readFile(process.argv[2], 'utf8'));
  const report = await api.analyzeReferences(args.input_path, {
    useLiveConnectors: true, mode: 'review', policy: 'strict',
  });
  process.stdout.write(JSON.stringify({ok: true, provider: '@jhlee0619/citecheck@0.1.0',
    library_modified: false, native_writeback_available: false, report}));
} catch (error) {
  process.stdout.write(JSON.stringify({ok: false, error: String(error), library_modified: false}));
  process.exitCode = 1;
}
