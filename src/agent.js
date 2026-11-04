#!/usr/bin/env node
import { parseArgs } from 'node:util';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { stat } from 'node:fs/promises';

const HELP = `Copilot Chat

Usage: copilot [options] [prompt]

  No prompt             Open the interactive editor
  [prompt]              Start a task and stay in the interactive session
  -p, --print           Run one task and exit (also used without a terminal)
  -C, --cwd <dir>       Start in a project directory
  -h, --help            Show this help

Interactive: /new, /web, /quit
Enter submits; Shift+Enter adds a line; Esc stops the current task.
Sessions are not saved after exit. Use /new to start a new task.
`;

try {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      help: { type: 'boolean', short: 'h' }, print: { type: 'boolean', short: 'p' },
      cwd: { type: 'string', short: 'C' }
    }
  });
  if (values.help) {
    console.log(HELP);
  } else {
    const prompt = positionals.join(' ').trim();
    const interactive = !values.print && Boolean(process.stdin.isTTY && process.stdout.isTTY);
    if (!interactive && !prompt) throw new Error('A prompt is required in print mode. Use a terminal for interactive mode.');
    if (process.platform !== 'win32') throw new Error('This CLI requires Windows.');
    // Only npm's own launcher changes cwd. A linked `copilot` always uses the actual current directory.
    const installDir = fileURLToPath(new URL('..', import.meta.url));
    const npmLauncher = process.env.npm_lifecycle_event === 'start'
      && resolve(process.cwd()) === resolve(installDir);
    const cwd = resolve(values.cwd ?? (npmLauncher ? process.env.INIT_CWD : undefined) ?? process.cwd());
    if (!(await stat(cwd)).isDirectory()) throw new Error(`Not a directory: ${cwd}`);
    const agentDir = fileURLToPath(new URL('../.local/pi', import.meta.url));
    // Set before importing Pi so terminal configuration stays local to this installation.
    process.env.PI_CODING_AGENT_DIR = agentDir;
    process.env.PI_OFFLINE = '1';
    const { runAgent } = await import('./runtime.js');
    await runAgent({ cwd, agentDir, prompt, interactive });
  }
} catch (error) {
  console.error(`Error: ${error.message}`);
  process.exitCode ||= 1;
}
