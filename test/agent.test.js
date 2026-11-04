import test from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, readFile, writeFile, mkdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createAgentSessionRuntime, SessionManager } from '@earendil-works/pi-coding-agent';
import { createRuntime } from '../src/runtime.js';

process.env.PI_OFFLINE = '1';
const packet = (request, value) => {
  const text = JSON.stringify({ requestId: request.requestId, ...value });
  return { completed: true, text, codeBlocks: [text] };
};
const final = (request, text = 'Done.') => packet(request, { type: 'final', text });
const call = (request, name, args) => packet(request, { type: 'tool_calls', calls: [{ name, arguments: args }] });

async function fixture(t, respond) {
  const root = await mkdtemp(join(tmpdir(), 'copilot-cli-test-'));
  const cwd = join(root, 'project'), agentDir = join(root, 'agent');
  await mkdir(cwd);
  const browsers = [];
  const factory = options => {
    const browser = {
      calls: [], closes: 0, history: [], context: null,
      async ask(prompt, options) {
        const request = JSON.parse(prompt.split('\nTASK_CONTEXT_JSON:\n')[1]);
        this.calls.push({ request, fresh: options.fresh, length: prompt.length });
        // Emulate the webpage retaining earlier messages, including its own assistant replies.
        if (request.mode === 'bootstrap') {
          this.context = request;
          this.history = [...request.messages];
        } else if (request.mode === 'continue') {
          assert.equal(options.fresh, false);
          assert.equal(options.requireContinuation, true);
          assert.equal(request.sessionId, this.context.sessionId);
          assert.equal(request.systemPrompt, undefined);
          assert.equal(request.tools, undefined);
          this.history.push(...request.messages);
        }
        const reply = await respond({ ...this.context, ...request, messages: this.history }, options);
        const value = JSON.parse(reply.codeBlocks[0]);
        this.history.push({ role: 'assistant', content: value.type === 'final'
          ? [{ type: 'text', text: value.text }]
          : value.calls.map((call, index) => ({ type: 'toolCall', id: `${value.requestId}_${index}`, ...call })) });
        return reply;
      },
      async close() { this.closes++; }
    };
    browsers.push(browser);
    return createRuntime({ ...options, browser });
  };
  const runtime = await createAgentSessionRuntime(factory, {
    cwd, agentDir, sessionManager: SessionManager.inMemory(cwd)
  });
  runtime.setRebindSession(session => session.bindExtensions({}));
  await runtime.session.bindExtensions({});
  t.after(async () => { await runtime.dispose(); await rm(root, { recursive: true, force: true }); });
  return { root, cwd, runtime, browsers };
}

test('Pi runs write, edit, Git Bash and read in the selected project, then continues the same chat', async t => {
  const operations = [
    ['write', { path: 'note.txt', content: 'before\n' }],
    ['edit', { path: 'note.txt', edits: [{ oldText: 'before', newText: 'after' }] }],
    ['bash', { command: 'cat note.txt' }],
    ['read', { path: 'note.txt', limit: 5 }]
  ];
  const { cwd, runtime, browsers } = await fixture(t, request => {
    const results = request.messages.filter(message => message.role === 'toolResult');
    assert.ok(results.every(result => !result.isError));
    const operation = operations[results.length];
    return operation ? call(request, ...operation) : final(request);
  });
  await runtime.session.prompt('Create and verify note.txt.');
  assert.equal(await readFile(join(cwd, 'note.txt'), 'utf8'), 'after\n');
  assert.equal(runtime.session.getLastAssistantText(), 'Done.');
  const replies = browsers[0].history.filter(m => m.role === 'toolResult');
  assert.match(replies[2].content[0].text, /after/);
  assert.match(replies[3].content[0].text, /after/);
  await runtime.session.prompt('Continue from those results.');
  assert.deepEqual(browsers[0].calls.map(call => call.fresh), [true, false, false, false, false, false]);
  assert.equal(browsers[0].history.filter(m => m.role === 'user').length, 2);
  assert.equal(browsers[0].calls.at(-1).request.messages.length, 1);
  assert.deepEqual(browsers[0].calls.slice(1, 5).map(call => call.request.messages.map(m => m.role)),
    [['toolResult'], ['toolResult'], ['toolResult'], ['toolResult']]);
});

test('large read results use a saved preview; the next user turn sends only new text', async t => {
  const { cwd, runtime, browsers } = await fixture(t, request => request.messages.some(m => m.role === 'toolResult')
    ? final(request, 'File inspected.') : call(request, 'read', { path: 'large.txt', limit: 200 }));
  await writeFile(join(cwd, 'large.txt'), ('x'.repeat(199) + '\n').repeat(100));
  await runtime.session.prompt('Read large.txt.');
  assert.equal(runtime.session.getLastAssistantText(), 'File inspected.');
  const preview = browsers[0].history.find(message => message.role === 'toolResult').content[0].text;
  const path = preview.match(/\nSaved tool result: (.+)\n/)[1];
  t.after(() => rm(path, { force: true }));
  const original = runtime.session.messages.find(message => message.role === 'toolResult').content[0].text;
  assert.equal(await readFile(path, 'utf8'), original);
  assert.ok(preview.length < original.length);
  await runtime.session.prompt('Continue using what you just read.');
  assert.equal(runtime.session.getLastAssistantText(), 'File inspected.');
  const browser = browsers[0];
  assert.ok(JSON.stringify({ ...browser.context, messages: browser.history }).length > 24000);
  assert.deepEqual(browser.calls.map(call => call.request.mode), ['bootstrap', 'continue', 'continue']);
  assert.ok(browser.calls.at(-1).length < 4000);
  assert.equal(browser.calls.at(-1).request.messages.length, 1);
  assert.equal(browser.calls.at(-1).request.messages[0].role, 'user');
});

test('Pi can read omitted command output from its saved result without executing the command again', async t => {
  let savedPath;
  const { cwd, runtime, browsers } = await fixture(t, request => {
    const results = request.messages.filter(message => message.role === 'toolResult');
    if (!results.length) return call(request, 'bash', { command: 'node output.cjs' });
    if (results.length === 1) {
      const result = results[0];
      savedPath = result.content[0].text.match(/\nSaved tool result: (.+)\n/)[1];
      t.after(() => rm(savedPath, { force: true }));
      assert.equal(result.isError, true);
      assert.match(result.content[0].text, /Command exited with code 1/);
      return call(request, 'read', { path: savedPath, offset: 100, limit: 2 });
    }
    assert.equal(results[1].isError, false);
    assert.match(results[1].content[0].text, /ROW_100[\s\S]*ROW_101/);
    return final(request, 'Output inspected.');
  });
  // Below Pi's native 50 KiB cap, but quotes/backslashes almost double the JSON sent to the webpage.
  const output = Array.from({ length: 200 }, (_, i) => `ROW_${String(i + 1).padStart(3, '0')} ` + '"\\'.repeat(118)).join('\n') + '\n';
  await writeFile(join(cwd, 'output.cjs'),
    `require('node:fs').appendFileSync('runs.txt', 'run\\n');\nprocess.stdout.write(${JSON.stringify(output)});\nprocess.exitCode = 1;\n`);
  await runtime.session.prompt('Run output.cjs once and inspect the result near line 100.');
  assert.equal(runtime.session.getLastAssistantText(), 'Output inspected.');
  const original = runtime.session.messages.find(message => message.role === 'toolResult');
  assert.ok(JSON.stringify(original.content).length > 96000);
  assert.equal(await readFile(savedPath, 'utf8'), original.content[0].text);
  assert.ok(browsers[0].calls[1].length < 24000);
  await runtime.session.prompt('Continue from the inspected result.');
  assert.equal(runtime.session.getLastAssistantText(), 'Output inspected.');
  assert.equal(await readFile(join(cwd, 'runs.txt'), 'utf8'), 'run\n');
  assert.deepEqual(browsers[0].calls.map(call => call.request.mode), ['bootstrap', 'continue', 'continue', 'continue']);
  assert.deepEqual(browsers[0].calls.at(-1).request.messages.map(message => message.role), ['user']);
});

test('Pi searches ignored-aware project files through generic tools, then reads the matching source', async t => {
  const { cwd, runtime, browsers } = await fixture(t, request => {
    const results = request.messages.filter(message => message.role === 'toolResult');
    const operations = [
      ['find', { pattern: '*.js' }],
      ['grep', { pattern: 'PROJECT_NEEDLE', path: '.', glob: '*.js' }],
      ['read', { path: 'src/main.js', offset: 1, limit: 5 }]
    ];
    assert.ok(results.every(result => !result.isError), JSON.stringify(results));
    const advertised = request.tools.filter(tool => ['find', 'grep'].includes(tool.name));
    assert.equal(advertised.length, 2);
    assert.doesNotMatch(JSON.stringify(advertised), /\b(?:git|rg|fd)\b|8000|JSON budget/i);
    assert.doesNotMatch(request.systemPrompt, /Use bash for file operations like ls, rg, find/);
    return operations[results.length] ? call(request, ...operations[results.length]) : final(request, 'Found the source.');
  });
  await promisify(execFile)('git', ['init', '-q'], { cwd });
  for (const dir of ['src', 'build', '.idea/patches']) await mkdir(join(cwd, dir), { recursive: true });
  await writeFile(join(cwd, '.gitignore'), 'build/\n.idea/\n');
  for (const path of ['src/main.js', 'build/generated.js', '.idea/patches/old.js']) {
    await writeFile(join(cwd, path), 'const name = "PROJECT_NEEDLE";\n');
  }
  await runtime.session.prompt('Find PROJECT_NEEDLE in source files and inspect the surrounding code.');
  assert.equal(runtime.session.getLastAssistantText(), 'Found the source.');
  const results = browsers[0].history.filter(message => message.role === 'toolResult');
  assert.equal(results[0].content[0].text, 'src/main.js');
  assert.match(results[1].content[0].text, /src\/main\.js:1:/);
  assert.match(results[2].content[0].text, /PROJECT_NEEDLE/);
  assert.deepEqual(browsers[0].calls.map(call => call.request.mode), ['bootstrap', 'continue', 'continue', 'continue']);
});

test('search failures reach the model as tool-contract errors without leaking the execution backend', async t => {
  const { runtime, browsers } = await fixture(t, request => request.messages.some(message => message.role === 'toolResult')
    ? final(request, 'Select a project directory.') : call(request, 'find', { pattern: '*.js' }));
  await runtime.session.prompt('Find JavaScript files.');
  const result = browsers[0].history.find(message => message.role === 'toolResult');
  assert.equal(result.isError, true);
  assert.match(result.content[0].text, /version-controlled project/);
  assert.doesNotMatch(result.content[0].text, /\bgit\b|ls-files|execFile|stderr/i);
  assert.equal(runtime.session.getLastAssistantText(), 'Select a project directory.');
});

test('/new closes the old browser and starts an independent session without persisting history', async t => {
  const { cwd, runtime, browsers } = await fixture(t, request => final(request));
  await runtime.session.prompt('Remember project one.');
  assert.equal(runtime.session.sessionManager.isPersisted(), false);
  assert.equal(runtime.session.sessionFile, undefined);
  await runtime.newSession();
  assert.equal(browsers[0].closes, 1);
  assert.equal(runtime.session.messages.filter(m => m.role === 'user').length, 0);
  await runtime.session.prompt('A different task.');
  assert.equal(runtime.cwd, cwd);
  const request = browsers[1].calls[0];
  assert.equal(request.fresh, true);
  assert.deepEqual(request.request.messages.filter(m => m.role === 'user').map(m => m.content[0].text),
    ['A different task.']);
  assert.equal(runtime.session.sessionManager.isPersisted(), false);
  assert.equal(runtime.session.sessionFile, undefined);
});

test('abort never replays the interrupted conversation; /new starts a clean task', async t => {
  let entered;
  const started = new Promise(resolve => { entered = resolve; });
  let count = 0;
  const { runtime, browsers } = await fixture(t, async (request, { signal }) => {
    if (++count === 2) {
      entered();
      await new Promise((_, reject) => signal.addEventListener('abort', () => reject(new Error('Aborted')), { once: true }));
    }
    return final(request);
  });
  await runtime.session.prompt('First task.');
  const pending = runtime.session.prompt('Slow task.');
  await started;
  await runtime.session.abort();
  await pending;
  assert.equal(count, 2);
  await runtime.session.prompt('A new instruction after stopping.');
  assert.match(runtime.session.messages.at(-1).errorMessage, /Use \/new/);
  assert.deepEqual(browsers[0].calls.map(call => call.fresh), [true, false]);
  await runtime.newSession();
  await runtime.session.prompt('A clean task.');
  assert.equal(runtime.session.getLastAssistantText(), 'Done.');
  assert.equal(browsers[1].calls.length, 1);
  assert.deepEqual(browsers[1].calls[0].request.messages.map(m => m.content[0].text), ['A clean task.']);
});
