import { chromium } from 'playwright-core';
import { fileURLToPath } from 'node:url';
import { setTimeout as delay } from 'node:timers/promises';

const profile = fileURLToPath(new URL('../.local/browser-profile', import.meta.url));
const input = '#m365-chat-editor-target-element:visible, [contenteditable="true"][data-lexical-editor="true"]:visible';
const responses = '[data-testid="markdown-reply"]:visible, [data-message-author-role="assistant"]:visible, [data-author="assistant"]:visible';
const message = '请只回复：浏览器连接成功';
const normalizeInput = text => text.replace(/\r\n?/g, '\n').replace(/^[\s\u200B\u200C\uFEFF]+|[\s\u200B\u200C\uFEFF]+$/g, '');

// Persistent launch adapted from https://github.com/alexistb2904/m365-copilot-browser-provider
// Reference: e029fe19891621684e6b4172ca79fa3990a52325; MIT license: licenses/m365-copilot-browser-provider.LICENSE.
let context;
let closed;
try {
  context = await chromium.launchPersistentContext(profile, {
    channel: 'msedge', headless: false, viewport: null, chromiumSandbox: true,
    ignoreDefaultArgs: ['--unsafely-disable-devtools-self-xss-warnings', '--enable-unsafe-swiftshader']
  });
  console.log('1. Edge 已启动。');
  closed = new Promise(resolve => context.once('close', resolve));
  const page = context.pages()[0] ?? await context.newPage();
  await page.goto('https://m365.cloud.microsoft/chat', { waitUntil: 'domcontentloaded', timeout: 60000 });
  console.log('2. 页面已打开，请在浏览器中完成登录。');
  const composer = page.locator(input).first();
  await composer.waitFor({ state: 'visible', timeout: 300000 });
  console.log('3. 已识别聊天输入框。');

  const stop = page.locator('button[data-testid="stop-button"]')
    .or(page.getByRole('button', { name: /^(stop(?: generating(?: response)?| responding)?|停止(?:生成|响应|回應)?)(?:\s*\([^)]*\))?$/i }))
    .filter({ visible: true }).first();
  const readInput = () => composer.evaluate(el => el.matches('input, textarea') ? el.value : (el.innerText ?? el.textContent ?? ''));
  if (await stop.isVisible()) throw new Error('网页仍在生成，请等待完成后重新验证。');
  if (normalizeInput(await readInput())) throw new Error('输入框已有草稿，请在网页处理后重新验证。');
  const replies = page.locator(responses);
  const previousCount = await replies.count();
  const previous = previousCount ? await replies.last().elementHandle() : null;
  const previousText = previous ? (await previous.innerText()).replace(/\r\n?/g, '\n') : '';
  await composer.fill(message);
  let actual = '';
  const inputDeadline = Date.now() + 3000;
  do {
    actual = await readInput();
    if (normalizeInput(actual) === message) break;
    await delay(100);
  } while (Date.now() < inputDeadline);
  if (normalizeInput(actual) !== message) {
    const codes = Array.from(actual, char => 'U+' + char.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')).join(' ');
    throw new Error(`输入校验未通过，没有点击发送。\n期望：${JSON.stringify(message)}\n读回：${JSON.stringify(actual)}\n字符码点：${codes}`);
  }
  console.log('输入内容已核对，准备发送。');
  await page.getByRole('button', { name: /^(send(?: message)?|发送|傳送)$/i })
    .filter({ visible: true }).first().click({ timeout: 15000 });
  console.log('4. 已点击发送：' + message);

  let sawGenerating = false, lastText = '', changedAt = Date.now(), completed = false;
  let candidate = '', candidateSince = 0, nextStatusAt = Date.now() + 10000;
  let previewActive = false, waitingNotice = false;
  const report = text => {
    if (previewActive && !lastText.endsWith('\n')) process.stdout.write('\n');
    previewActive = false;
    console.log(text);
  };
  const deadline = Date.now() + 180000;
  while (Date.now() < deadline) {
    const reply = await replies.evaluateAll((elements, previous) => {
      const last = elements.at(-1);
      if (!last) return { count: 0, text: '', busy: false, finished: false };
      const owner = last.closest('[data-message-author-role="assistant"], [data-author="assistant"]');
      const nodes = owner && owner !== last ? [last, owner] : [last];
      const states = nodes.flatMap(node => ['aria-busy', 'data-is-streaming', 'data-streaming'].map(name => node.getAttribute(name)));
      return {
        count: elements.length, fresh: last !== previous, text: (last.innerText ?? '').replace(/\r\n?/g, '\n'),
        busy: states.includes('true'), finished: states.includes('false')
      };
    }, previous);
    const stopVisible = await stop.isVisible();
    const generating = stopVisible || reply.busy;
    sawGenerating ||= generating;
    const current = reply.text.trim() && (reply.count > previousCount || reply.text !== previousText || (sawGenerating && reply.fresh));
    if (current && reply.text !== lastText) {
      const append = reply.text.startsWith(lastText);
      if (!lastText) report('5. 已读取回复（实时预览，尚未确认结束）：');
      else if (!append) report('[网页已改写回复，以下全文替换之前的预览]');
      else if (!previewActive) report('[继续输出]');
      process.stdout.write(append ? reply.text.slice(lastText.length) : reply.text);
      previewActive = true;
      lastText = reply.text;
      changedAt = Date.now();
      waitingNotice = false;
    }
    // Completion needs UI evidence and an unchanged reply while generation stays idle.
    if (current && !generating && (sawGenerating || reply.finished)) {
      if (candidate !== reply.text) { candidate = reply.text; candidateSince = Date.now(); }
      else if (Date.now() - candidateSince >= 1500) { completed = true; break; }
    } else {
      candidate = '';
      candidateSince = 0;
    }
    if (current && !generating && !sawGenerating && !reply.finished && !waitingNotice && Date.now() - changedAt >= 5000) {
      report('正文已读取，尚无结束信号；继续监听后续内容。');
      waitingNotice = true;
    }
    if (Date.now() >= nextStatusAt) {
      report(`回复检测：区域=${reply.count}，字数=${reply.text.length}，停止按钮=${stopVisible}，生成标记=${reply.busy}，结束标记=${reply.finished}`);
      nextStatusAt = Date.now() + 10000;
    }
    await delay(300);
  }
  if (!completed) throw new Error(lastText
    ? '已显示实时预览，但等待生成结束超时，完整性尚未确认；没有重新发送。'
    : '没有定位到本轮回复，等待超时；没有重新发送。请检查网页和回复检测日志。');
  report('6. 已确认生成结束。');
  report('请核对网页与终端的回复。关闭 Edge 窗口结束本次验证。');
} catch (error) {
  console.error('\n验证中断：', error.message);
  if (context) console.error('请检查网页；关闭 Edge 窗口结束本次验证。');
  process.exitCode = 1;
} finally {
  await closed;
  await context?.close();
}
