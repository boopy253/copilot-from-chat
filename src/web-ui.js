import { basename } from 'node:path';
import { truncateToWidth } from '@earendil-works/pi-tui';

// Only terminal integration lives here; Pi owns editing and tool rendering.
export function createWebUI(browser) {
  let ui;
  return {
    onStatus(message) { ui?.setWorkingMessage(message); },
    extension(pi) {
      pi.on('session_start', (_event, ctx) => {
        if (ctx.mode !== 'tui') return;
        ui = ctx.ui;
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
      pi.on('agent_start', () => { ui?.setWorkingMessage('Waiting for Copilot...'); });
      pi.on('tool_execution_start', () => { ui?.setWorkingMessage('Running tools...'); });
      pi.on('agent_end', () => { ui?.setWorkingMessage(); });
      pi.on('session_shutdown', async () => { ui = undefined; await browser.close(); });
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
