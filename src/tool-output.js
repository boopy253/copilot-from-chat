import { randomUUID } from 'node:crypto';
import { writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// This is a soft budget for NEW tool-result text, not a browser input or conversation limit.
// Pi's usual 50 KiB output cap can exceed it even once, especially after JSON escapes double in size.
const TOOL_TEXT_BUDGET = 18000;
const size = content => JSON.stringify(content).length;

async function saveOutput(text, signal) {
  const path = join(tmpdir(), `copilot-output-${randomUUID()}.txt`);
  await writeFile(path, text, { encoding: 'utf8', flag: 'wx', mode: 0o600, signal });
  // Ordinary temporary tool output, like Pi's saved command logs; no history index or session recovery.
  return path.replaceAll('\\', '/');
}

function preview(text, path, budget) {
  const characters = Array.from(text); // Do not split surrogate pairs when shortening Unicode output.
  const note = `\n\n[Partial output]\nSaved tool result: ${path}\nUse read with offset/limit to inspect this file in small sections. For very long lines, use bash to inspect a small byte range.`;
  const content = count => {
    const head = Math.ceil(count * 2 / 3), tail = count - head;
    return [{ type: 'text', text: characters.slice(0, head).join('') + '\n\n[... middle omitted ...]\n\n'
      + (tail ? characters.slice(-tail).join('') : '') + note }];
  };
  let low = 0, high = characters.length - 1, best = content(0);
  while (low <= high) {
    const middle = Math.floor((low + high) / 2), candidate = content(middle);
    if (size(candidate) <= budget) { best = candidate; low = middle + 1; }
    else high = middle - 1;
  }
  // Both ends preserve useful context and command failure/status lines. The original returned text is saved
  // unchanged; if Pi had already shortened it, its original continuation/full-log notice is saved as well.
  return best;
}

export async function limitToolResults(messages, { signal, save = saveOutput } = {}) {
  signal?.throwIfAborted();
  const results = messages.map((message, index) => ({ message, index }))
    .filter(({ message }) => message.role === 'toolResult');
  if (results.reduce((sum, { message }) => sum + size(message.content), 0) <= TOOL_TEXT_BUDGET) return messages;

  const limited = [...messages];
  let remaining = TOOL_TEXT_BUDGET, count = results.length;
  // Small results stay whole; large results share the remaining budget across this batch, not per tool.
  for (const { message, index } of results.sort((a, b) => size(a.message.content) - size(b.message.content))) {
    signal?.throwIfAborted();
    // Every result must retain a readable continuation note. Very large batches can exceed this soft
    // target through their notes/metadata; that still must not become a fabricated browser input limit.
    const share = Math.max(512, Math.floor(remaining / count--));
    let content = message.content;
    if (size(content) > share) {
      const text = typeof content === 'string' ? content : content.map(block => block.text).join('\n\n');
      const path = await save(text, signal);
      signal?.throwIfAborted();
      content = preview(text, path, share);
      // Change only the transport copy. Keep tool IDs, errors and the local transcript intact, so subsequent
      // delta comparisons still use the real Pi history. Never repeat the command to recover omitted output.
      limited[index] = { ...message, content };
    }
    remaining -= size(content);
  }
  return limited;
}
