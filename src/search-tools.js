import { execFile } from 'node:child_process';
import { stat } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import { promisify } from 'node:util';

const exec = promisify(execFile);
// Pi's normal grep/find need rg/fd and can try downloading them. The target has only Git for Windows.
// Delegate ignore matching to Git, including nested rules, negation, info/exclude and global excludes.
// These tools require a Git working tree; never fall back to an unrestricted recursive shell search.
async function git(cwd, args, signal, input) {
  signal?.throwIfAborted();
  try {
    const pending = exec('git', ['-c', 'core.quotepath=false', ...args], {
      cwd, signal, windowsHide: true, encoding: 'utf8', timeout: 15000, maxBuffer: 512 * 1024
    });
    pending.child.stdin.on('error', () => {}); // A failed Git process can close stdin before consuming the file list.
    pending.child.stdin.end(input);
    return { output: (await pending).stdout, truncated: false };
  } catch (error) {
    signal?.throwIfAborted();
    if (error.code === 1 && ['grep', 'check-ignore'].includes(args[0])) return { output: '', truncated: false };
    if (error.code === 'ERR_CHILD_PROCESS_STDIO_MAXBUFFER' && error.message.startsWith('stdout')) {
      return { output: error.stdout, truncated: true };
    }
    // Tool errors also reach the web model. Keep execution details in the local cause, not the tool reply.
    if (error.code === 'ENOENT') throw new Error('Search is unavailable. Check the local installation.', { cause: error });
    if (error.killed) throw new Error('Search timed out. Narrow the search path or pattern.', { cause: error });
    throw new Error('Search failed. Check the pattern and select a directory inside a version-controlled project.', { cause: error });
  }
}

function records(result, separator) {
  // Never present a partial filename/match from the process buffer as a complete result.
  const end = result.output.lastIndexOf(separator);
  const text = result.truncated ? result.output.slice(0, end + 1) : result.output;
  return text.split(separator).filter(Boolean).map(line => separator === '\n' ? line.replace(/\r$/, '') : line);
}

function formatResult(lines, limit, truncated, kind) {
  const selected = [];
  let size = 0, linesTruncated = false;
  for (const line of lines) {
    // Web chat has a per-message budget, unlike Pi's usual API transport. Count JSON-escaped characters
    // because tool results are serialized again for the composer; keep room for the protocol envelope.
    const clipped = kind === 'grep' && line.length > 500;
    const text = clipped ? Array.from(line).slice(0, 500).join('') + ' [line truncated; use read]' : line;
    const cost = JSON.stringify(text).length;
    if (selected.length >= limit || size + cost > 8000) { truncated = true; break; }
    linesTruncated ||= clipped;
    selected.push(text);
    size += cost;
  }
  let text = selected.join('\n') || (truncated ? 'No complete results within the output limit.' : 'No matches.');
  if (truncated || linesTruncated) text += '\n\n[Partial results. Narrow the pattern/path, then use read with offset/limit for file contents.]';
  return {
    content: [{ type: 'text', text }],
    // Preserve the details used by Pi's built-in grep/find renderers; the notice also reaches the model.
    details: kind === 'grep'
      ? { ...(truncated ? { matchLimitReached: selected.length } : {}), linesTruncated }
      : { ...(truncated ? { resultLimitReached: selected.length } : {}) }
  };
}

const pathParameter = { type: 'string', description: 'Search directory, relative to the project or absolute (default: .)' };
const limitParameter = { type: 'integer', minimum: 1, maximum: 200, description: 'Maximum result lines (default: 50, maximum: 200)' };

export function createSearchTools(cwd) {
  // Keep the model-facing contract independent of the backend. Model guidance describes search semantics;
  // Git commands, rg/fd availability and the web encoding budget belong to this host implementation.
  return [
    {
      name: 'grep', label: 'grep',
      description: 'Search file contents. Returns path:line:content relative to path. Narrow the scope if results are partial; use read for full context.',
      promptSnippet: 'Search file contents',
      promptGuidelines: ['Use the grep tool for content searches. Use read with offset/limit around matching line numbers.'],
      parameters: { type: 'object', required: ['pattern'], additionalProperties: false, properties: {
        pattern: { type: 'string', minLength: 1, description: 'Text to find; extended regular expression when literal=false' },
        path: { ...pathParameter, description: 'Directory or file to search, relative to the project or absolute (default: .)' },
        glob: { type: 'string', description: 'Filename wildcard, e.g. *.js (directory searches only)' },
        literal: { type: 'boolean', description: 'Literal matching (default: true); false uses extended regex' },
        ignoreCase: { type: 'boolean', description: 'Ignore letter case (default: false)' }, limit: limitParameter
      } },
      async execute(_id, { pattern, path = '.', glob, literal = true, ignoreCase = false, limit = 50 }, signal, _onUpdate, ctx) {
        signal?.throwIfAborted();
        const target = resolve(ctx?.cwd || cwd, path), directory = (await stat(target)).isDirectory();
        if (!directory && glob) throw new Error('Use glob with a directory path, or search the file without glob.');
        const args = ['grep', '--untracked', '--exclude-standard', '-n', '-I', '--no-color', '--no-heading', '--no-break', '--no-column', literal ? '-F' : '-E'];
        if (ignoreCase) args.push('-i');
        // Use argv, not shell interpolation: patterns and filenames can contain quotes, spaces or option-like text.
        args.push('-e', pattern, '--', directory ? (glob || '.') : `:(literal)${basename(target)}`);
        const result = await git(directory ? target : dirname(target), args, signal);
        return formatResult(records(result, '\n'), limit, result.truncated, 'grep');
      }
    },
    {
      name: 'find', label: 'find',
      description: 'Find files by filename pattern. Returns paths relative to path. Narrow the scope if results are partial.',
      promptSnippet: 'Find files by filename pattern',
      promptGuidelines: ['Use the find tool to discover files. Narrow path or pattern when results are partial.'],
      parameters: { type: 'object', required: ['pattern'], additionalProperties: false, properties: {
        pattern: { type: 'string', minLength: 1, description: 'Filename wildcard, e.g. *.ts, src/**, or * for all files' },
        path: pathParameter, limit: limitParameter
      } },
      async execute(_id, { pattern, path = '.', limit = 50 }, signal, _onUpdate, ctx) {
        const target = resolve(ctx?.cwd || cwd, path);
        const result = await git(target, ['ls-files', '--cached', '--others', '--exclude-standard', '-z', '--', pattern], signal);
        const files = [...new Set(records(result, '\0'))];
        // ls-files still includes tracked files that now match ignore rules. --no-index checks those too,
        // so ignored build artifacts/IDE patches cannot leak into find results just because they were committed.
        const ignored = files.length ? await git(target, ['check-ignore', '--no-index', '-z', '--stdin'], signal, files.join('\0') + '\0') : { output: '' };
        if (ignored.truncated) throw new Error('Search is too broad. Narrow the path or pattern.');
        const excluded = new Set(records(ignored, '\0'));
        return formatResult(files.filter(file => !excluded.has(file)), limit, result.truncated, 'find');
      }
    }
  ];
}
