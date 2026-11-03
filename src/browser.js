import { chromium } from 'playwright-core';
import { fileURLToPath } from 'node:url';
import { setTimeout as delay } from 'node:timers/promises';
import { mkdir, writeFile } from 'node:fs/promises';

const CHAT_URL = 'https://m365.cloud.microsoft/chat';
const PROFILE = fileURLToPath(new URL('../.local/browser-profile', import.meta.url));
const INPUT_DIAGNOSTIC = new URL('../.local/input-diagnostic.json', import.meta.url);
const INPUT = '#m365-chat-editor-target-element:visible, [contenteditable="true"][data-lexical-editor="true"]:visible';
const RESPONSES = '[data-testid="markdown-reply"]:visible, [data-message-author-role="assistant"]:visible, [data-author="assistant"]:visible';
const normalize = text => text.replace(/\r\n?/g, '\n').replace(/^[\s\u200B\u200C\uFEFF]+|[\s\u200B\u200C\uFEFF]+$/g, '');

// Runs inside the page. innerText adds visual paragraph spacing; textContent loses <br> boundaries.
function readComposerText(el) {
  if (el.matches('input, textarea')) return el.value;
  const read = node => {
    if (node.nodeType === Node.TEXT_NODE) return node.textContent;
    if (node.nodeType !== Node.ELEMENT_NODE) return '';
    if (node.tagName === 'BR') return '\n';
    const children = [...node.childNodes].filter(child => child.nodeType === Node.ELEMENT_NODE
      || (child.nodeType === Node.TEXT_NODE && child.textContent !== ''));
    // Editable blocks use a final <br> as a caret placeholder. A preceding <br> remains a real newline.
    if ((node === el || /^(P|DIV)$/.test(node.tagName)) && children.at(-1)?.nodeName === 'BR') children.pop();
    let text = '', previousBlock = false;
    for (const [index, child] of children.entries()) {
      const block = child.nodeType === Node.ELEMENT_NODE && /^(P|DIV)$/.test(child.nodeName);
      if (index && (block || previousBlock)) text += '\n';
      text += read(child);
      previousBlock = block;
    }
    return text;
  };
  return read(el);
}

function compareInput(actual, expected, contentEditable) {
  const ignored = {}, wanted = expected.split('\n');
  const received = normalize(actual).split('\n').map((line, index) => {
    // Accept only extra suffix markers AFTER an exact match of the expected line, including its spaces.
    // Within normalized lines, expected zero-width characters stay mandatory; line breaks also stay exact.
    if (!contentEditable || wanted[index] === undefined || !line.startsWith(wanted[index])) return line;
    const extra = line.slice(wanted[index].length);
    if (!/^[\u200B\u200C\uFEFF]+$/.test(extra)) return line;
    for (const char of extra) {
      const point = 'U+' + char.codePointAt(0).toString(16).toUpperCase().padStart(4, '0');
      ignored[point] = (ignored[point] ?? 0) + 1;
    }
    return wanted[index];
  }).join('\n');
  return { received, matches: received === expected, ignored };
}

// Launch adapted from alexistb2904/m365-copilot-browser-provider, e029fe19891621684e6b4172ca79fa3990a52325.
// MIT license: licenses/m365-copilot-browser-provider.LICENSE.
export class CopilotBrowser {
  context = null;
  page = null;
  closed = Promise.resolve();
  busy = false;
  uncertain = false;
  used = false;

  async open({ signal, onStatus = () => {} } = {}) {
    signal?.throwIfAborted();
    if (this.context && !this.page?.isClosed()) return;
    this.context = await chromium.launchPersistentContext(PROFILE, {
      channel: 'msedge', headless: false, viewport: null, chromiumSandbox: true,
      ignoreDefaultArgs: ['--unsafely-disable-devtools-self-xss-warnings', '--enable-unsafe-swiftshader']
    });
    const context = this.context;
    this.closed = new Promise(resolve => context.once('close', resolve));
    try {
      signal?.throwIfAborted();
      this.context.setDefaultTimeout(5000);
      this.page = context.pages()[0] ?? await context.newPage();
      onStatus('1. Edge 已启动。');
      await this.page.goto(CHAT_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
      onStatus('2. 页面已打开，请在浏览器中完成登录。');
      await this.page.locator(INPUT).first().waitFor({ state: 'visible', timeout: 300000 });
      onStatus('3. 已识别聊天输入框。');
    } catch (error) {
      await this.close();
      throw error;
    }
  }

  async ask(message, { signal, fresh = false, onSnapshot = () => {}, onStatus = () => {} } = {}) {
    if (this.busy) throw new Error('浏览器已有请求正在进行。');
    if (this.uncertain) throw new Error('上一轮发送结果未确认，请检查网页并重启程序；禁止自动重发。');
    if (typeof message !== 'string' || !message.trim()) throw new Error('消息不能为空。');
    this.busy = true;
    let sent = false;
    const abort = () => { void this.close().catch(() => {}); };
    signal?.addEventListener('abort', abort, { once: true });
    try {
      await this.open({ signal, onStatus });
      signal?.throwIfAborted();
      if (fresh && (this.used || await this.page.locator(RESPONSES).count())) {
        const old = this.page;
        this.page = await this.context.newPage();
        await this.page.goto(CHAT_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
        await this.page.locator(INPUT).first().waitFor({ state: 'visible', timeout: 300000 });
        await old.close();
        if (await this.page.locator(RESPONSES).count()) {
          const reset = this.page.getByRole('button', { name: /^(?:(?:start (?:a )?)?new (?:chat|conversation)|新建聊天|新建对话|新聊天|新对话|开始新聊天)(?:\s*\([^)]*\))?$/i }).filter({ visible: true }).first();
          await reset.click();
          await this.page.locator(RESPONSES).first().waitFor({ state: 'hidden', timeout: 15000 });
        }
      }
      const page = this.page;
      const composer = page.locator(INPUT).first();
      const replies = page.locator(RESPONSES);
      const stop = page.locator('button[data-testid="stop-button"]')
        .or(page.getByRole('button', { name: /^(stop(?: generating(?: response)?| responding)?|停止(?:生成|响应|回應)?)(?:\s*\([^)]*\))?$/i }))
        .filter({ visible: true }).first();
      const readInput = () => composer.evaluate(readComposerText);
      if (await stop.isVisible()) throw new Error('网页仍在生成，请等待完成后重新验证。');
      if (normalize(await readInput())) throw new Error('输入框已有草稿，请在网页处理后重新验证。');
      const previousCount = await replies.count();
      const previous = previousCount ? await replies.last().elementHandle() : null;
      const previousText = previous ? (await previous.innerText()).replace(/\r\n?/g, '\n') : '';
      const lines = message.replace(/\r\n?/g, '\n').split('\n');
      const contentEditable = await composer.evaluate(el => el.isContentEditable);
      const canUseLineBreaks = lines.length > 1 && contentEditable;
      const expected = normalize(message), attempts = [];
      let actual = '', verified = false;
      let comparison;
      for (const method of canUseLineBreaks ? ['fill', 'lines'] : ['fill']) {
        signal?.throwIfAborted();
        if (method === 'lines') {
          // Only replace an unsent draft; never retry after the page appears to have submitted it.
          const count = await replies.count();
          if (await stop.isVisible() || count !== previousCount
            || (count && (await replies.last().innerText()).replace(/\r\n?/g, '\n') !== previousText)
            || !normalize(await readInput())) {
            this.uncertain = true;
            throw new Error('输入期间网页状态发生变化，停止重新输入；请检查是否已发送。');
          }
          onStatus(`整段输入未通过回读校验，改为逐行输入（${lines.length} 行，${message.length} 字符）。`);
          await composer.fill(lines[0]);
          for (const line of lines.slice(1)) {
            signal?.throwIfAborted();
            await composer.press('Shift+Enter');
            if (line) await page.keyboard.insertText(line);
          }
        } else {
          onStatus(`正在输入并校验 ${message.length} 字符。`);
          await composer.fill(message);
        }
        const started = Date.now(), samples = [];
        let matchingSince = 0;
        do {
          signal?.throwIfAborted();
          actual = await readInput();
          comparison = compareInput(actual, expected, contentEditable);
          const { matches, ignored } = comparison;
          samples.push({ elapsedMs: Date.now() - started, length: actual.length, matches, ignored });
          matchingSince = matches ? matchingSince || Date.now() : 0;
          if (matchingSince && Date.now() - matchingSince >= 500) { verified = true; break; }
          await delay(100);
        } while (Date.now() - started < 3000);
        const dom = verified ? undefined : await composer.evaluate(el => ({
          tag: el.tagName, contentEditable: el.isContentEditable, focused: el === document.activeElement || el.contains(document.activeElement),
          whiteSpace: getComputedStyle(el).whiteSpace, textTransform: getComputedStyle(el).textTransform, maxLength: el.getAttribute('maxlength'),
          paragraphs: el.querySelectorAll('p').length, lineBreaks: el.querySelectorAll('br').length,
          value: el.matches('input, textarea') ? el.value : null,
          innerText: el.innerText, textContent: el.textContent, html: el.outerHTML
        }));
        attempts.push({ method, verified, actual, comparison, samples, dom });
        if (!verified) {
          // Capture original browser evidence before changing the draft, even if the fallback later fails.
          try {
            await mkdir(new URL('.', INPUT_DIAGNOSTIC), { recursive: true });
            await writeFile(INPUT_DIAGNOSTIC, JSON.stringify({
              capturedAt: new Date().toISOString(), expected: message, reader: 'value for input/textarea; DOM text with paragraph and br boundaries for contenteditable',
              normalization: 'CRLF to LF; trim document boundaries; for contenteditable, ignore extra U+200B/U+200C/U+FEFF suffixes only after an exact match of each expected line', attempts
            }, null, 2), 'utf8');
            onStatus(`输入诊断已保存：${fileURLToPath(INPUT_DIAGNOSTIC)}`);
          } catch (error) { onStatus(`保存输入诊断失败：${error.message}`); }
          onStatus(`输入框结构：P=${dom.paragraphs}，BR=${dom.lineBreaks}，显示文本=${dom.innerText?.length ?? 0} 字符，逻辑文本=${actual.length} 字符。`);
        }
        if (verified) break;
      }
      if (!verified) {
        const { received } = comparison;
        if (received === expected) throw new Error('最后一次读回内容一致，但未持续稳定 500 毫秒；没有点击发送，请查看输入诊断。');
        let offset = 0;
        while (offset < expected.length && expected[offset] === received[offset]) offset++;
        const excerpt = text => JSON.stringify(text.slice(Math.max(0, offset - 30), offset + 90));
        const point = text => offset < text.length ? 'U+' + text.codePointAt(offset).toString(16).toUpperCase().padStart(4, '0') : 'EOF';
        const codes = text => Array.from(text.slice(offset)).slice(0, 8).map(char => 'U+' + char.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')).join(' ') || 'EOF';
        throw new Error(`输入校验未通过，没有点击发送。期望 ${expected.length} 字符，原始读回 ${normalize(actual).length} 字符，比较后 ${received.length} 字符。\n首个差异：偏移 ${offset}（从 0 开始），期望 ${point(expected)}，读回 ${point(received)}。\n差异后码点：期望 ${codes(expected)}；读回 ${codes(received)}。\n期望片段：${excerpt(expected)}\n读回片段：${excerpt(received)}`);
      }
      const ignored = Object.entries(comparison.ignored).map(([point, count]) => `${point} × ${count}`).join('，');
      if (ignored) onStatus(`回读校验已忽略行尾额外零宽标记：${ignored}。`);
      onStatus('输入内容已核对，准备发送。');
      signal?.throwIfAborted();
      sent = true;
      this.used = true;
      await page.getByRole('button', { name: /^(send(?: message)?|发送|傳送)$/i }).filter({ visible: true }).first().click({ timeout: 15000 });
      onStatus('4. 已点击发送。');

      let sawGenerating = false, lastText = '', changedAt = Date.now();
      let candidate = '', candidateSince = 0, nextStatusAt = Date.now() + 10000, waitingNotice = false;
      const deadline = Date.now() + 180000;
      while (Date.now() < deadline) {
        signal?.throwIfAborted();
        const reply = await replies.evaluateAll((elements, previous) => {
          const last = elements.at(-1);
          if (!last) return { count: 0, text: '', codeBlocks: [], busy: false, finished: false };
          const owner = last.closest('[data-message-author-role="assistant"], [data-author="assistant"]');
          const nodes = owner && owner !== last ? [last, owner] : [last];
          const states = nodes.flatMap(node => ['aria-busy', 'data-is-streaming', 'data-streaming'].map(name => node.getAttribute(name)));
          const body = owner ?? last;
          const visible = node => node.getClientRects().length > 0 && getComputedStyle(node).visibility !== 'hidden';
          const pres = [...body.querySelectorAll('pre')];
          if (body.matches('pre')) pres.unshift(body);
          const blocks = new Set(pres.filter(visible).map(pre => pre.querySelector('code') ?? pre));
          for (const code of body.querySelectorAll('code')) {
            if (!code.closest('pre') && visible(code) && /^\s*[\[{]/.test(code.textContent ?? '')) blocks.add(code);
          }
          const readCode = node => {
            if (node.nodeType === Node.TEXT_NODE) return node.textContent;
            if (node.nodeType !== Node.ELEMENT_NODE || /^(BUTTON|STYLE|SCRIPT)$/.test(node.nodeName)) return '';
            if (node.nodeName === 'BR') return '\n';
            let text = '', previousBlock = false;
            for (const child of node.childNodes) {
              if (child.nodeType === Node.COMMENT_NODE || /^(BUTTON|STYLE|SCRIPT)$/.test(child.nodeName)) continue;
              const block = /^(P|DIV)$/.test(child.nodeName);
              if (text && (block || previousBlock) && !text.endsWith('\n')) text += '\n';
              text += readCode(child);
              previousBlock = block;
            }
            return text;
          };
          return {
            count: elements.length, fresh: last !== previous, text: (last.innerText ?? '').replace(/\r\n?/g, '\n'),
            protocolText: (body.innerText ?? '').replace(/\r\n?/g, '\n'),
            codeBlocks: [...blocks].map(readCode),
            busy: states.includes('true'), finished: states.includes('false')
          };
        }, previous);
        const stopVisible = await stop.isVisible();
        const generating = stopVisible || reply.busy;
        sawGenerating ||= generating;
        const current = reply.text.trim() && (reply.count > previousCount || reply.text !== previousText || (sawGenerating && reply.fresh));
        if (current && reply.text !== lastText) {
          await onSnapshot(reply);
          lastText = reply.text;
          changedAt = Date.now();
          waitingNotice = false;
        }
        if (current && !generating && (sawGenerating || reply.finished)) {
          const signature = JSON.stringify([reply.text, reply.protocolText, reply.codeBlocks]);
          if (candidate !== signature) { candidate = signature; candidateSince = Date.now(); }
          else if (Date.now() - candidateSince >= 1500) {
            signal?.throwIfAborted();
            return { text: reply.text, protocolText: reply.protocolText, codeBlocks: reply.codeBlocks, completed: true };
          }
        } else { candidate = ''; candidateSince = 0; }
        if (current && !generating && !sawGenerating && !reply.finished && !waitingNotice && Date.now() - changedAt >= 5000) {
          onStatus('正文已读取，尚无结束信号；继续监听后续内容。');
          waitingNotice = true;
        }
        if (Date.now() >= nextStatusAt) {
          onStatus(`回复检测：区域=${reply.count}，字数=${reply.text.length}，停止按钮=${stopVisible}，生成标记=${reply.busy}，结束标记=${reply.finished}`);
          nextStatusAt = Date.now() + 10000;
        }
        await delay(300);
      }
      throw new Error(lastText ? '已读取正文，但等待生成结束超时，完整性尚未确认。' : '没有定位到本轮回复，等待超时。');
    } catch (error) {
      if (sent) this.uncertain = true;
      signal?.throwIfAborted();
      throw error;
    } finally {
      signal?.removeEventListener('abort', abort);
      this.busy = false;
    }
  }

  async close() {
    const context = this.context;
    this.context = null;
    await context?.close();
  }
}
