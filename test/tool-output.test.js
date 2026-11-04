import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, rm } from 'node:fs/promises';
import { limitToolResults } from '../src/tool-output.js';

test('saved results preserve the complete returned text and partial/continuation notices', async t => {
  const text = 'first line\r\n' + '"\\🙂'.repeat(20000) + '\n[Earlier truncation: original log is elsewhere.]';
  const messages = [{ role: 'toolResult', toolCallId: 'call-0', toolName: 'read', isError: false,
    content: [{ type: 'text', text }] }];
  const limited = await limitToolResults(messages);
  const preview = limited[0].content[0].text;
  const path = preview.match(/\nSaved tool result: (.+)\n/)[1];
  t.after(() => rm(path, { force: true }));
  assert.equal(await readFile(path, 'utf8'), text);
  assert.match(preview, /Earlier truncation: original log is elsewhere/);
  assert.ok(preview.isWellFormed());
  assert.ok(JSON.stringify(limited[0].content).length <= 18000);
  assert.equal(messages[0].content[0].text, text);
});

test('small results and arbitrarily large user messages are neither shortened nor saved', async () => {
  const messages = [{ role: 'user', content: 'x'.repeat(100000) },
    { role: 'toolResult', content: [{ type: 'text', text: 'Small.' }] }];
  const limited = await limitToolResults(messages, { save: () => { throw new Error('Must not save'); } });
  assert.equal(limited, messages);
});

test('an aborted exchange does not save tool output', async () => {
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(limitToolResults([{ role: 'toolResult', content: [{ type: 'text', text: 'x'.repeat(50000) }] }], {
    signal: controller.signal, save: () => { throw new Error('Must not save'); }
  }), { name: 'AbortError' });
});
