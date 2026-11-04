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
const MODEL_NAME = 'GPT-5.6 Sol Think deeper';
const MODEL_LABEL = /^GPT[\s\u2010-\u2015-]*5\.6\s+Sol\s+Think\s+deeper$/i;

// M365 controls can have an accessible name such as "Auto, Model Selector", while menu rows
// include a title AND a description. Read their UI fields separately; never search chat prose.
// DOM evidence: CodeWebChat/CWC apps/browser/src/content-scripts/send-prompt-content-script/chatbots/copilot.ts
// and SerK21/CopilotSelecter src/content/applier.js. These are UI selectors, not hidden model/API IDs.
function readModelControls(elements) {
  const clean = text => (text ?? '').replace(/[\u200B\u200C\uFEFF]/g, '').replace(/[\u2010-\u2015\u2212]/g, '-')
    .replace(/\s+/g, ' ').replace(/,?\s*Model Selector\s*$/i, '').trim();
  return elements.map(el => {
    const lines = (el.innerText ?? '').split('\n').map(clean).filter(Boolean);
    const heading = el.querySelector('h4, .fai-CapabilityPickerMenuItem__primaryContentWrapper');
    const state = el.getAttribute('aria-checked') ?? el.getAttribute('aria-selected') ?? el.getAttribute('aria-pressed');
    const visibleCheck = [...el.querySelectorAll('svg[data-testid^="checkmark"]')].some(svg => {
      if (!svg.getClientRects().length) return false;
      // Some M365 versions render a checkmark in EVERY row and hide unselected ones on an ancestor.
      // Presence of the SVG alone is not proof; an explicit false ARIA state also takes precedence.
      for (let node = svg; node && node !== el.parentElement; node = node.parentElement) {
        const style = getComputedStyle(node);
        if (style.display === 'none' || style.visibility !== 'visible' || Number(style.opacity) === 0) return false;
      }
      return true;
    });
    return {
      labels: [...new Set([clean(heading?.innerText), lines[0], clean(lines.slice(0, 2).join(' ')), clean(el.getAttribute('aria-label'))].filter(Boolean))],
      text: clean(lines.join(' ')).slice(0, 160),
      submenu: /^(?:true|menu|listbox)$/.test(el.getAttribute('aria-haspopup') ?? ''),
      testId: el.getAttribute('data-testid') ?? el.getAttribute('data-test-id'),
      selected: state === null ? visibleCheck : state === 'true',
      disabled: el.disabled === true || el.getAttribute('aria-disabled') === 'true' || el.getAttribute('aria-busy') === 'true'
    };
  });
}

// Runs inside the page for both streaming reads and conversation continuity checks.
function readReply(elements, previous) {
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
}

const replyFingerprint = reply => JSON.stringify(reply.codeBlocks?.length ? reply.codeBlocks : [reply.protocolText ?? reply.text ?? '']);

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
  closing = Promise.resolve();
  busy = false;
  uncertain = false;
  used = false;
  lastReply = null;

  async hasContinuation() {
    if (!this.context || !this.page || this.page.isClosed() || this.lastReply?.page !== this.page) return false;
    const reply = await this.page.locator(RESPONSES).evaluateAll(readReply, null);
    return !reply.busy && replyFingerprint(reply) === this.lastReply.signature;
  }

  async selectModel({ signal } = {}) {
    // This is a webpage choice, not Pi model metadata or an API model ID. M365 can reset it to Auto
    // for every new chat. Use the user's menu path and verify the selection before entering a task.
    // Account instructions/memory settings are deliberately left to the user; they affect other chats.
    const page = this.page;
    const pickerNames = /model selector|select (?:a )?model|选择模型|模型选择|^(?:Auto|自动|Quick response|Think deeper|GPT(?:[\s\u2010-\u2015-].*)?|Claude(?:\s.*)?)$/i;
    const menus = page.locator('[role="menu"]:visible, [role="listbox"]:visible, .fui-MenuPopover:visible');
    const items = page.locator('[role="menuitem"], [role="menuitemradio"], [role="menuitemcheckbox"], [role="option"], button, [role="button"]')
      .and(page.locator('[role="menu"] *, [role="listbox"] *, .fui-MenuPopover *')).filter({ visible: true });
    let picker, step = 'picker', touchedMenu = false;
    const read = async locator => (await locator.evaluateAll(readModelControls))[0];
    const matches = info => info?.labels.some(label => MODEL_LABEL.test(label));
    const wait = async (find, timeout = 5000) => {
      const deadline = Date.now() + timeout;
      do {
        signal?.throwIfAborted();
        const value = await find();
        if (value) return value;
        await delay(100, undefined, { signal });
      } while (Date.now() < deadline);
      return null;
    };
    const findItem = async predicate => {
      const choices = (await items.evaluateAll(readModelControls)).map((info, index) => ({ info, locator: items.nth(index) }))
        .filter(({ info }) => !info.disabled && predicate(info));
      if (choices.length > 1) throw new Error('Multiple matching model controls are visible.');
      return choices[0];
    };
    const target = () => findItem(info => !info.submenu && matches(info));
    const family = () => findItem(info => info.submenu && (info.testId === 'gptSubMenuModelTrigger-OpenAI'
      || info.labels.some(label => /^(?:GPT(?:\s+OpenAI)?|OpenAI)$/i.test(label))));
    const openOptions = async () => {
      step = 'menu';
      touchedMenu = true;
      if (!await items.count()) await picker.click();
      const choice = await wait(async () => await target() ?? await family());
      if (!choice) throw new Error('The model menu has no matching GPT entry.');
      if (!choice.info.submenu) return choice;
      step = 'gpt-submenu';
      // The family row is often "GPT OpenAI" and its submenu is portalled beside the main menu.
      // Do not depend on DOM ancestry or click an already expanded family a second time.
      await choice.locator.hover();
      let leaf = await wait(target, 800);
      if (!leaf && await choice.locator.getAttribute('aria-expanded') !== 'true') {
        await choice.locator.click();
        leaf = await wait(target, 800);
      }
      if (!leaf) {
        await choice.locator.press('ArrowRight');
        leaf = await wait(target);
      }
      if (!leaf) throw new Error('The GPT submenu has no matching model title.');
      return leaf;
    };
    try {
      picker = await wait(async () => {
        const byId = page.locator('#gptModeSwitcher:visible');
        const candidates = await byId.count() ? byId : page.getByRole('button', { name: pickerNames })
          .or(page.getByRole('combobox', { name: pickerNames })).filter({ visible: true });
        const count = await candidates.count();
        if (count > 1) throw new Error('Multiple model selectors are visible.');
        const info = count ? await read(candidates) : null;
        return info && !info.disabled ? candidates : null;
      });
      if (!picker) throw new Error('The model selector is missing or not ready.');
      if (matches(await read(picker))) return;
      const choice = await openOptions();
      step = 'select';
      if (!choice.info.selected) await choice.locator.click();
      step = 'verify';
      const confirmed = async () => matches(await read(picker)) || (await target())?.info.selected;
      if (await wait(confirmed, 800)) return;
      // Compact toolbar labels omit the effort. Reopen and check selection state without clicking
      // the model again. Menu closure is not required: some builds keep the selected row visible.
      await openOptions();
      step = 'verify';
      if (!await wait(confirmed)) throw new Error('The requested model is not shown as selected.');
    } catch (cause) {
      signal?.throwIfAborted();
      let evidence = '';
      try {
        const current = picker ? (await read(picker))?.text : 'not found';
        const labels = (await items.evaluateAll(readModelControls)).slice(0, 8).map(info => (info.text || info.labels[0] || '').slice(0, 100));
        evidence = ` Picker=${JSON.stringify(current)}; options=${JSON.stringify(labels)}.`;
      } catch { /* The page may have closed; retain the original failure. */ }
      // Short control-only evidence is enough to report from the restricted machine. Do not dump
      // chat content, HTML, account settings or a long Playwright call log into the terminal.
      throw new Error(`Could not confirm ${MODEL_NAME} (step=${step}). ${cause.message.split('\n')[0]}${evidence} Nothing was entered or sent.`, { cause });
    } finally {
      if (touchedMenu) {
        try {
          for (let level = 0; level < 3 && await menus.count(); level++) {
            await page.keyboard.press('Escape');
            await delay(50);
          }
        } catch { /* Closing a menu must not hide the selection failure or an abort. */ }
      }
    }
  }

  async open({ signal, onStatus = () => {} } = {}) {
    await this.closing;
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
      await this.page.goto(CHAT_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
      if (!await this.page.locator(INPUT).first().isVisible()) onStatus('Sign in to Microsoft 365 in the Edge window.');
      await this.page.locator(INPUT).first().waitFor({ state: 'visible', timeout: 300000 });
    } catch (error) {
      await this.close();
      throw error;
    }
  }

  async ask(message, { signal, fresh = false, requireContinuation = false, onSnapshot = () => {}, onStatus = () => {} } = {}) {
    if (this.busy) throw new Error('A browser request is already in progress.');
    if (this.uncertain && !fresh) throw new Error('The previous exchange is uncertain. Start a new web conversation before sending another request.');
    if (typeof message !== 'string' || !message.trim()) throw new Error('The message cannot be empty.');
    this.busy = true;
    let sent = false;
    const abort = () => { void this.close().catch(() => {}); };
    signal?.addEventListener('abort', abort, { once: true });
    try {
      signal?.throwIfAborted();
      if (requireContinuation && !await this.hasContinuation()) {
        throw Object.assign(new Error('The previous web conversation is no longer active. Nothing was entered or sent.'), { code: 'COPILOT_CONTEXT_LOST' });
      }
      if (fresh) this.lastReply = null;
      await this.open({ signal, onStatus });
      signal?.throwIfAborted();
      const draft = normalize(await this.page.locator(INPUT).first().evaluate(readComposerText));
      if (fresh && (this.used || this.uncertain || await this.page.locator(RESPONSES).count()
        || draft)) {
        const old = this.page;
        this.page = await this.context.newPage();
        await this.page.goto(CHAT_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
        await this.page.locator(INPUT).first().waitFor({ state: 'visible', timeout: 300000 });
        // Keep an existing unsent draft available for inspection in its original tab.
        if (!draft) await old.close();
        if (await this.page.locator(RESPONSES).count()) {
          const reset = this.page.getByRole('button', { name: /^(?:(?:start (?:a )?)?new (?:chat|conversation)|新建聊天|新建对话|新聊天|新对话|开始新聊天)(?:\s*\([^)]*\))?$/i }).filter({ visible: true }).first();
          await reset.click();
          await this.page.locator(RESPONSES).first().waitFor({ state: 'hidden', timeout: 15000 });
        }
      }
      if (fresh || !this.used) await this.selectModel({ signal });
      if (fresh) this.uncertain = false;
      const page = this.page;
      const composer = page.locator(INPUT).first();
      const replies = page.locator(RESPONSES);
      const stop = page.locator('button[data-testid="stop-button"]')
        .or(page.getByRole('button', { name: /^(stop(?: generating(?: response)?| responding)?|停止(?:生成|响应|回應)?)(?:\s*\([^)]*\))?$/i }))
        .filter({ visible: true }).first();
      const readInput = () => composer.evaluate(readComposerText);
      if (await stop.isVisible()) throw new Error('The webpage is still generating. Wait for it to finish before trying again.');
      if (normalize(await readInput())) throw new Error('The webpage contains an unsent draft. Resolve it in the browser before trying again.');
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
            throw new Error('The webpage changed during input. Stopped editing the draft; check whether it was submitted.');
          }
          onStatus('Retrying input with explicit line breaks...');
          await composer.fill(lines[0]);
          for (const line of lines.slice(1)) {
            signal?.throwIfAborted();
            await composer.press('Shift+Enter');
            if (line) await page.keyboard.insertText(line);
          }
        } else {
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
            onStatus(`Input diagnostics saved: ${fileURLToPath(INPUT_DIAGNOSTIC)}`);
          } catch (error) { onStatus(`Could not save input diagnostics: ${error.message}`); }
        }
        if (verified) break;
      }
      if (!verified) {
        const { received } = comparison;
        if (received === expected) throw new Error('Input matched but did not remain stable for 500 ms. Nothing was sent; check the input diagnostics.');
        let offset = 0;
        while (offset < expected.length && expected[offset] === received[offset]) offset++;
        const excerpt = text => JSON.stringify(text.slice(Math.max(0, offset - 30), offset + 90));
        const point = text => offset < text.length ? 'U+' + text.codePointAt(offset).toString(16).toUpperCase().padStart(4, '0') : 'EOF';
        const codes = text => Array.from(text.slice(offset)).slice(0, 8).map(char => 'U+' + char.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')).join(' ') || 'EOF';
        throw new Error(`Input verification failed; nothing was sent. Expected ${expected.length} characters, read ${normalize(actual).length}, compared ${received.length}.\nFirst difference at offset ${offset} (zero-based): expected ${point(expected)}, read ${point(received)}.\nCode points: expected ${codes(expected)}; read ${codes(received)}.\nExpected: ${excerpt(expected)}\nRead: ${excerpt(received)}\nDiagnostics: ${fileURLToPath(INPUT_DIAGNOSTIC)}`);
      }
      signal?.throwIfAborted();
      if (requireContinuation && (!await this.hasContinuation() || await stop.isVisible())) {
        this.uncertain = true;
        throw new Error('The web conversation changed during input. Nothing was sent; inspect the draft before continuing.');
      }
      sent = true;
      this.used = true;
      await page.getByRole('button', { name: /^(send(?: message)?|发送|傳送)$/i }).filter({ visible: true }).first().click({ timeout: 15000 });
      onStatus('Waiting for Copilot...');

      let sawGenerating = false, lastText = '';
      let candidate = '', candidateSince = 0;
      // Think deeper can spend several minutes reasoning before visible answer text appears.
      // Keep the wait cancellable and still require the normal generation-complete checks.
      const deadline = Date.now() + 600000;
      while (Date.now() < deadline) {
        signal?.throwIfAborted();
        const reply = await replies.evaluateAll(readReply, previous);
        const stopVisible = await stop.isVisible();
        const generating = stopVisible || reply.busy;
        sawGenerating ||= generating;
        const current = reply.text.trim() && (reply.count > previousCount || reply.text !== previousText || (sawGenerating && reply.fresh));
        if (current && reply.text !== lastText) {
          await onSnapshot(reply);
          lastText = reply.text;
        }
        if (current && !generating && (sawGenerating || reply.finished)) {
          const signature = JSON.stringify([reply.text, reply.protocolText, reply.codeBlocks]);
          if (candidate !== signature) { candidate = signature; candidateSince = Date.now(); }
          else if (Date.now() - candidateSince >= 1500) {
            signal?.throwIfAborted();
            this.lastReply = { page, signature: replyFingerprint(reply) };
            return { text: reply.text, protocolText: reply.protocolText, codeBlocks: reply.codeBlocks, completed: true };
          }
        } else { candidate = ''; candidateSince = 0; }
        await delay(300);
      }
      throw new Error(lastText ? 'Reply text was found, but generation did not finish before the timeout. Its completeness is unconfirmed.' : 'Timed out waiting for the current reply.');
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
    this.lastReply = null;
    const context = this.context;
    if (context) {
      this.context = null;
      this.page = null;
      this.closing = context.close();
    }
    await this.closing;
  }
}
