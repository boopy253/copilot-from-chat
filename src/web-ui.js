import { basename } from 'node:path';
import { truncateToWidth } from '@earendil-works/pi-tui';

// Only terminal integration lives here; Pi owns editing and tool rendering.
export function createWebUI(browser) {
  let ui, workingMessage;
  const onStatus = message => {
    if (!ui || message === workingMessage) return;
    workingMessage = message;
    ui.setWorkingMessage(message);
  };
  return {
    onStatus,
    extension(pi) {
      pi.registerToolRenderer((name, next) => {
        const renderer = next();
        if (name !== 'bash' || !renderer?.renderCall) return renderer;
        return {
          ...renderer,
          // Pi 1.5's Bash renderCall uses executionStarted only to start its UI stopwatch.
          // In regular mode, its per-second updates replay all history when the row is off-screen
          // (e.g. earlier parallel tools). Suppress that clock only in the render context: actual
          // execution, live output, and the final recorded duration still use Pi's native behavior.
          renderCall: (args, theme, context) => renderer.renderCall(args, theme, { ...context, executionStarted: false })
        };
      });
      pi.on('session_start', (_event, ctx) => {
        if (ctx.mode !== 'tui') return;
        ui = ctx.ui;
        workingMessage = undefined;
        // Web replies are buffered, so an 80 ms spinner produces terminal writes without new text.
        // A single ASCII frame disables Pi's animation timer and preserves Git Bash scrollback.
        ui.setWorkingIndicator({ frames: ['*'] });
        ui.setTitle(`Copilot Chat - ${basename(ctx.cwd)}`);
        ui.setHeader((_tui, theme) => ({
          invalidate() {},
          render: width => [
            theme.fg('accent', 'Copilot Chat'),
            theme.fg('dim', 'Enter to send · Esc to stop · /new · /quit')
          ].map(line => truncateToWidth(line, width))
        }));
        ui.setFooter((tui, theme, data) => {
          const unsubscribe = data.onBranchChange(() => tui.requestRender());
          return {
            dispose: unsubscribe, invalidate() {},
            // Web chat exposes no reliable token usage. Omit that counter instead of adding a permanent notice.
            render: width => [
              theme.fg('dim', `${ctx.cwd}${data.getGitBranch() ? ` (${data.getGitBranch()})` : ''}`)
            ].map(line => truncateToWidth(line, width))
          };
        });
      });
      pi.on('agent_start', () => onStatus('Waiting for Copilot...'));
      pi.on('tool_execution_start', () => onStatus('Running tools...'));
      pi.on('agent_end', () => onStatus());
      pi.on('session_shutdown', async () => { ui = undefined; workingMessage = undefined; await browser.close(); });
      // No real usage/limit is exposed by the webpage. Token-based compaction is unreliable here.
      pi.on('session_before_compact', (_event, ctx) => {
        ctx.ui.notify('Compaction is unavailable. Use /new to start a new task.', 'warning');
        return { cancel: true };
      });
      pi.registerCommand('web', {
        description: 'Show workspace and tools',
        handler: async (_args, ctx) => {
          ctx.ui.notify([
            `Workspace: ${ctx.cwd}`,
            'Tools: read, bash, edit, write, grep, find'
          ].join('\n'), 'info');
        }
      });
    }
  };
}
