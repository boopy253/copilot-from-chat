import { fileURLToPath } from 'node:url';
import { InMemoryCredentialStore, InMemoryModelsStore } from '@earendil-works/pi-ai';
import { createAgentSession, ModelRuntime, DefaultResourceLoader, SessionManager, SettingsManager } from '@earendil-works/pi-coding-agent';
import { CopilotBrowser } from './browser.js';
import { createCopilotProvider } from './pi-provider.js';

const task = process.argv.slice(2).join(' ').trim();
const browser = new CopilotBrowser();
let session;
const interrupt = () => { process.exitCode = 130; void Promise.allSettled([session?.abort(), browser.close()]); };

try {
  console.error(`Pi Agent 入口：${fileURLToPath(import.meta.url)}`);
  if (!task) throw new Error('用法：npm run agent -- "读取 package.json，告诉我项目名称"');
  console.error(`任务：${task}`);
  if (process.platform !== 'win32') throw new Error('此入口面向 Windows + Edge，请在目标电脑运行。');
  const cwd = process.cwd();
  const agentDir = fileURLToPath(new URL('../.local/pi', import.meta.url));
  const settingsManager = SettingsManager.inMemory({
    compaction: { enabled: false }, retry: { enabled: false }, shellPath: process.env.COPILOT_BASH_PATH || undefined
  });
  const resourceLoader = new DefaultResourceLoader({
    cwd, agentDir, settingsManager, noExtensions: true, disabledBuiltinExtensions: ['mcp'],
    noSkills: true, noPromptTemplates: true, noThemes: true, noContextFiles: true,
    appendSystemPrompt: ['The bash tool runs Git Bash on Windows. Use Bash/POSIX commands, Bash quoting and forward-slash paths.']
  });
  await resourceLoader.reload();
  const modelRuntime = await ModelRuntime.create({
    credentials: new InMemoryCredentialStore(), modelsStore: new InMemoryModelsStore(),
    modelsPath: null, refreshOnCreate: false
  });
  modelRuntime.registerNativeProvider(createCopilotProvider(browser, { onStatus: text => console.error(`[网页] ${text}`) }));
  ({ session } = await createAgentSession({
    cwd, agentDir, settingsManager, resourceLoader, modelRuntime,
    model: modelRuntime.getModel('copilot-web', 'm365-web'), thinkingLevel: 'off',
    tools: ['read', 'bash', 'edit', 'write'], sessionManager: SessionManager.inMemory(cwd)
  }));
  session.subscribe(event => {
    if (event.type === 'message_update' && event.assistantMessageEvent.type === 'text_delta') process.stdout.write(event.assistantMessageEvent.delta);
    if (event.type === 'tool_execution_start') console.error(`\n[Pi] 执行 ${event.toolName} ${JSON.stringify(event.args)}`);
    if (event.type === 'tool_execution_end') console.error(`[Pi] ${event.toolName} ${event.isError ? '失败，结果将交回模型' : '完成'}`);
  });
  process.on('SIGINT', interrupt);
  console.error(`工作目录：${cwd}\n本地工具：${session.getActiveToolNames().join(', ')}\nCtrl+C 中止；网页无法提供真实 token 用量，本次关闭自动压缩和重试。`);
  await session.prompt(task);
  const last = session.state.messages.findLast(message => message.role === 'assistant');
  if (!last || last.stopReason !== 'stop') throw new Error(last?.errorMessage ?? '任务没有正常结束。');
  console.log();
} catch (error) {
  console.error(`\n任务中断：${error.message}`);
  process.exitCode ||= 1;
  if (browser.context && process.exitCode !== 130) {
    console.error('请检查网页；关闭 Edge 窗口结束本次验证。');
    await browser.closed;
  }
} finally {
  process.off('SIGINT', interrupt);
  session?.dispose();
  await browser.close();
}
