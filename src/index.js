import { CopilotBrowser } from './browser.js';
import { fileURLToPath } from 'node:url';

const browser = new CopilotBrowser();
let lastText = '', previewActive = false;
const report = text => {
  if (previewActive && !lastText.endsWith('\n')) process.stdout.write('\n');
  previewActive = false;
  console.log(text);
};

try {
  report(`浏览器验证入口：${fileURLToPath(import.meta.url)}`);
  if (process.argv.length > 2 || process.env.npm_lifecycle_event === 'agent') {
    throw new Error('当前入口只做固定的浏览器连通验证。执行任务请用 npm run agent -- "任务内容"，并确认 package.json 的 agent 指向 node src/agent.js。');
  }
  await browser.ask('请只回复：浏览器连接成功', {
    onStatus: report,
    onSnapshot: ({ text }) => {
      const append = text.startsWith(lastText);
      if (!lastText) report('5. 已读取回复（实时预览，尚未确认结束）：');
      else if (!append) report('[网页已改写回复，以下全文替换之前的预览]');
      else if (!previewActive) report('[继续输出]');
      process.stdout.write(append ? text.slice(lastText.length) : text);
      lastText = text;
      previewActive = true;
    }
  });
  report('6. 已确认生成结束。');
  report('请核对网页与终端的回复。关闭 Edge 窗口结束本次验证。');
} catch (error) {
  console.error('\n验证中断：', error.message);
  if (browser.context) console.error('请检查网页；关闭 Edge 窗口结束本次验证。');
  process.exitCode = 1;
} finally {
  await browser.closed;
  await browser.close();
}
