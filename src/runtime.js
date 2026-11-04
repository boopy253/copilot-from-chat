import { InMemoryCredentialStore, InMemoryModelsStore } from '@earendil-works/pi-ai';
import {
  createAgentSessionRuntime, createAgentSessionServices, createAgentSessionFromServices,
  InteractiveMode, ModelRuntime, SessionManager, SettingsManager
} from '@earendil-works/pi-coding-agent';
import { CopilotBrowser } from './browser.js';
import { createCopilotProvider } from './pi-provider.js';
import { createWebUI } from './web-ui.js';
import { createSearchTools } from './search-tools.js';

// Pi 1.5 history commands cannot rewind the webpage's conversation.
class CopilotInteractiveMode extends InteractiveMode {
  async showSessionSelector() { this.showWarning('Saved-session restore is disabled. Use /new to start a new chat.'); }
  async handleImportCommand() { this.showWarning('Session import is disabled. Use /new to start a new chat.'); }
  showUserMessageSelector() { this.showWarning('Session branching is unavailable. Use /new to start a new chat.'); }
  showTreeSelector() { this.showWarning('History rewind is unavailable. Use /new to start a new chat.'); }
  async handleCloneCommand() { this.showWarning('Session cloning is unavailable. Use /new to start a new chat.'); }
}

export async function createRuntime({ cwd, agentDir, sessionManager, sessionStartEvent, browser = new CopilotBrowser() }) {
  const webUI = createWebUI(browser);
  const settingsManager = SettingsManager.inMemory({
    // Let Git Bash's terminal own mouse scrolling and scrollback.
    tuiMode: 'regular',
    compaction: { enabled: false }, retry: { enabled: false }, cacheWarming: 'off',
    branchSummary: { skipPrompt: true }, quietStartup: true,
    enableInstallTelemetry: false, enableAnalytics: false, showCacheMissNotices: false,
    images: { blockImages: true }, shellPath: process.env.COPILOT_BASH_PATH || undefined
  });
  const modelRuntime = await ModelRuntime.create({
    credentials: new InMemoryCredentialStore(), modelsStore: new InMemoryModelsStore(),
    modelsPath: null, refreshOnCreate: false
  });
  modelRuntime.registerNativeProvider(createCopilotProvider(browser, { onStatus: webUI.onStatus }));
  const services = await createAgentSessionServices({
    cwd, agentDir, settingsManager, modelRuntime,
    resourceLoaderOptions: {
      noExtensions: true, disabledBuiltinExtensions: ['mcp'], extensionFactories: [webUI.extension],
      noSkills: true, noPromptTemplates: true, noThemes: true,
      appendSystemPrompt: [
        // The model needs the shell platform and tool behavior, not our search implementation or wire budget.
        'Use Bash syntax and forward-slash paths on Windows. ' +
        'Use grep and find to search project files; use bash for commands, builds and tests. ' +
        'Keep tool output focused: read files with offset/limit and bound command output. Inspect large files in sections.'
      ]
    }
  });
  return {
    ...(await createAgentSessionFromServices({
      services, sessionManager, sessionStartEvent,
      model: modelRuntime.getModel('copilot-web', 'm365-web'), thinkingLevel: 'off',
      // Override Pi's rg/fd implementations with the Git-backed tools required by the offline target.
      tools: ['read', 'bash', 'edit', 'write', 'grep', 'find'], customTools: createSearchTools(cwd)
    })),
    services, diagnostics: services.diagnostics
  };
}

export async function runAgent({ cwd, agentDir, prompt, interactive }) {
  const runtime = await createAgentSessionRuntime(createRuntime, { cwd, agentDir, sessionManager: SessionManager.inMemory(cwd) });
  let mode;
  const interrupt = () => { process.exitCode = 130; void runtime.session.abort(); };
  try {
    if (interactive) {
      mode = new CopilotInteractiveMode(runtime, { initialMessage: prompt || undefined });
      await mode.run();
    } else {
      const session = runtime.session;
      await session.bindExtensions({});
      session.subscribe(event => {
        if (event.type === 'message_update' && event.assistantMessageEvent.type === 'text_delta') {
          process.stdout.write(event.assistantMessageEvent.delta);
        }
        if (event.type === 'tool_execution_start') console.error(`[tool] ${event.toolName}`);
        if (event.type === 'tool_execution_end' && event.isError) console.error(`[tool] ${event.toolName} failed`);
      });
      process.on('SIGINT', interrupt);
      process.on('SIGTERM', interrupt);
      await session.prompt(prompt);
      const last = session.state.messages.findLast(message => message.role === 'assistant');
      if (process.exitCode !== 130 && (!last || last.stopReason !== 'stop')) {
        throw new Error(last?.errorMessage ?? 'The task did not complete.');
      }
      process.stdout.write('\n');
    }
  } finally {
    process.off('SIGINT', interrupt);
    process.off('SIGTERM', interrupt);
    mode?.stop();
    await runtime.dispose();
  }
}
