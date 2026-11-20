import test, { before, after } from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from 'playwright-core';
import { CopilotBrowser } from '../src/browser.js';

let engine;
before(async () => {
  try { engine = await chromium.launch({ channel: process.platform === 'win32' ? 'msedge' : 'chrome', headless: true }); }
  catch (error) {
    if (!/Executable doesn't exist|distribution .* is not found/.test(error.message)) throw error;
    // Use an installed browser only; tests must not download one in the restricted target environment.
  }
});
after(async () => { await engine?.close(); });

async function fixture(t, config = {}) {
  if (!engine) { t.skip('An installed Edge/Chrome browser is required.'); return; }
  const context = await engine.newContext();
  context.setDefaultTimeout(1500);
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body>
    <button id="picker" aria-haspopup="menu">Auto</button>
    <div id="menu" role="menu" hidden><div id="gpt" role="menuitem" tabindex="0" aria-haspopup="menu">GPT</div></div>
    <div id="submenu" role="menu" hidden>
      <div id="target" role="menuitemradio" tabindex="0" aria-checked="false">GPT-5.6 Sol <span>Think deeper</span></div>
      <div role="menuitemradio" tabindex="0" aria-checked="false">GPT-5.6 Sol Quick response</div>
    </div>
    <div id="m365-chat-editor-target-element" contenteditable="true" data-lexical-editor="true"></div>
    <button id="send">Send</button><button id="stop" data-testid="stop-button" hidden>Stop</button>
    <main id="replies"></main>
    <script>
      const config = ${JSON.stringify(config)};
      const picker = document.getElementById('picker'), menu = document.getElementById('menu');
      const submenu = document.getElementById('submenu'), target = document.getElementById('target');
      let composer = document.getElementById('m365-chat-editor-target-element');
      let selected = !!config.selected;
      window.actions = [];
      window.sent = [];
      window.inputSelections = [];
      const label = () => config.auto ? (config.autoLabel || 'Auto') : config.compact ? 'GPT-5.6 Sol' : 'GPT-5.6 Sol Think deeper';
      if (config.referenceMarkup) {
        picker.id = 'gptModeSwitcher';
        picker.setAttribute('aria-label', 'Auto, Model Selector');
        document.getElementById('gpt').innerHTML = '<h4>GPT</h4><span>OpenAI</span>';
        document.getElementById('gpt').setAttribute('data-testid', 'gptSubMenuModelTrigger-OpenAI');
        // M365 accessible names include descriptions, and family test IDs can also appear on leaves.
        target.setAttribute('data-testid', 'gptSubMenuModelTrigger-OpenAI');
        target.innerHTML = '<div class="fai-CapabilityPickerMenuItem__primaryContentWrapper"><h4>GPT-5.6 Sol</h4></div><div>Think deeper</div><p>Spend more time on complex tasks.</p>';
      }
      if (config.labelledPicker) picker.setAttribute('aria-label', 'Auto, Model Selector');
      if (config.auto) {
        picker.textContent = 'GPT-5.6 Sol Think deeper';
        if (config.labelledPicker) picker.setAttribute('aria-label', 'GPT-5.6 Sol Think deeper, Model Selector');
        target.innerHTML = '<h4>' + label() + '</h4><p>Choose the best model automatically.</p>';
        menu.append(target);
      }
      if (config.checkmark) {
        target.removeAttribute('aria-checked');
        const mark = document.createElement('span');
        mark.id = 'check'; mark.style.visibility = 'hidden';
        mark.innerHTML = '<svg data-testid="checkmark-model" width="10" height="10"><path d="M0 5 L5 10 L10 0" /></svg>';
        target.append(mark);
      }
      if (selected) {
        picker.textContent = label(); target.setAttribute('aria-checked', 'true');
        if (config.labelledPicker) picker.setAttribute('aria-label', label() + ', Model Selector');
      }
      picker.onclick = () => { window.actions.push('picker'); menu.hidden = !menu.hidden; submenu.hidden = true; };
      const expand = () => { window.actions.push('gpt'); submenu.hidden = false; };
      document.getElementById('gpt').onclick = expand;
      if (config.hover) document.getElementById('gpt').onpointerenter = () => setTimeout(expand, 100);
      if (config.missing) target.remove();
      target.onclick = () => {
        window.actions.push('target');
        if (!config.ignore) {
          selected = true; picker.textContent = label();
          if (config.referenceMarkup || config.labelledPicker) picker.setAttribute('aria-label', label() + ', Model Selector');
          if (config.checkmark) document.getElementById('check').style.visibility = 'visible';
          else target.setAttribute('aria-checked', 'true');
        }
        if (!config.keepMenu) { menu.hidden = true; submenu.hidden = true; }
      };
      document.addEventListener('keydown', event => {
        if (event.key !== 'Escape') return;
        if (!submenu.hidden) submenu.hidden = true; else menu.hidden = true;
      });
      const onInput = () => {
        if (!composer.textContent) return;
        window.inputSelections.push(selected);
        if (!config.rebuildEditor || window.rebuilt) return;
        window.rebuilt = true;
        const replacement = composer.cloneNode(true);
        if (config.roleEditor) {
          replacement.removeAttribute('id');
          replacement.removeAttribute('data-lexical-editor');
          replacement.setAttribute('role', 'textbox');
          replacement.setAttribute('aria-label', 'Message Copilot');
        }
        replacement.oninput = onInput;
        const placeholder = document.createComment('editor remount');
        composer.replaceWith(placeholder);
        setTimeout(() => { placeholder.replaceWith(replacement); composer = replacement; }, config.rebuildDelay ?? 0);
      };
      composer.oninput = onInput;
      document.getElementById('send').onclick = () => {
        // Simulate the submitted lines; textContent drops the editor's <br> boundaries.
        window.sent.push({ text: composer.innerText, selected });
        composer.textContent = '';
        const stop = document.getElementById('stop');
        stop.hidden = false;
        const reply = document.createElement('article');
        reply.setAttribute('data-message-author-role', 'assistant');
        reply.setAttribute('aria-busy', 'true');
        reply.textContent = 'Thinking';
        document.getElementById('replies').append(reply);
        setTimeout(() => {
          reply.textContent = 'Reply ' + window.sent.length;
          reply.setAttribute('aria-busy', 'false');
          stop.hidden = true;
        }, 80);
      };
    </script></body></html>`;
  await context.route('https://m365.cloud.microsoft/chat', route => route.fulfill({ contentType: 'text/html', body: html }));
  const browser = new CopilotBrowser({ model: config.auto ? 'Auto' : 'GPT-5.6 Sol Think deeper' });
  browser.context = context;
  browser.page = await context.newPage();
  t.after(() => browser.close());
  await browser.page.goto('https://m365.cloud.microsoft/chat');
  return browser;
}

test('new chats select the requested model before input; continuations preserve the same chat', async t => {
  const browser = await fixture(t);
  if (!browser) return;
  assert.equal((await browser.ask('First task.', { fresh: true })).completed, true);
  const page = browser.page;
  assert.deepEqual(await page.evaluate(() => window.sent), [{ text: 'First task.', selected: true }]);
  assert.deepEqual(await page.evaluate(() => window.actions), ['picker', 'gpt', 'target']);
  assert.ok((await page.evaluate(() => window.inputSelections)).every(Boolean));
  assert.equal((await browser.ask('Tool result.', { requireContinuation: true })).completed, true);
  assert.equal(browser.page, page);
  assert.deepEqual(await page.evaluate(() => window.actions), ['picker', 'gpt', 'target']);
  assert.equal((await browser.ask('New task.', { fresh: true })).completed, true);
  assert.notEqual(browser.page, page);
  assert.deepEqual(await browser.page.evaluate(() => window.sent), [{ text: 'New task.', selected: true }]);
  assert.deepEqual(await browser.page.evaluate(() => window.actions), ['picker', 'gpt', 'target']);
});

test('a model already shown as selected is left unchanged', async t => {
  const browser = await fixture(t, { selected: true });
  if (!browser) return;
  await browser.selectModel();
  assert.deepEqual(await browser.page.evaluate(() => window.actions), []);
});

test('the web model can be configured through the environment', t => {
  const previous = process.env.COPILOT_WEB_MODEL;
  t.after(() => {
    if (previous === undefined) delete process.env.COPILOT_WEB_MODEL;
    else process.env.COPILOT_WEB_MODEL = previous;
  });
  process.env.COPILOT_WEB_MODEL = 'auto';
  assert.equal(new CopilotBrowser().modelName, 'Auto');
  process.env.COPILOT_WEB_MODEL = 'unknown';
  assert.throws(() => new CopilotBrowser(), /Unsupported COPILOT_WEB_MODEL/);
  delete process.env.COPILOT_WEB_MODEL;
  assert.equal(new CopilotBrowser().modelName, 'Auto');
});

for (const autoLabel of ['Auto', '自动']) {
  test(`one-level ${autoLabel} is selected before input and on each new chat`, async t => {
    const browser = await fixture(t, { auto: true, autoLabel, labelledPicker: true });
    if (!browser) return;
    for (const text of ['First Auto task.', 'New Auto task.']) {
      assert.equal((await browser.ask(text, { fresh: true })).completed, true);
      assert.deepEqual(await browser.page.evaluate(() => window.actions), ['picker', 'target']);
      assert.deepEqual(await browser.page.evaluate(() => window.sent), [{ text, selected: true }]);
      assert.ok((await browser.page.evaluate(() => window.inputSelections)).every(Boolean));
      assert.equal(await browser.page.getByRole('menu').count(), 0);
    }
  });
}

test('Auto already selected does not open the menu', async t => {
  const browser = await fixture(t, { auto: true, selected: true, labelledPicker: true });
  if (!browser) return;
  await browser.selectModel();
  assert.deepEqual(await browser.page.evaluate(() => window.actions), []);
});

for (const scenario of ['missing', 'ignore']) {
  test(`Auto selection failure (${scenario}) stops before input without entering the GPT submenu`, async t => {
    const browser = await fixture(t, { auto: true, [scenario]: true });
    if (!browser) return;
    await assert.rejects(browser.ask('Do not send.', { fresh: true }), /Could not confirm Auto/);
    assert.deepEqual(await browser.page.evaluate(() => window.sent), []);
    assert.deepEqual(await browser.page.evaluate(() => window.inputSelections), []);
    assert.equal(await browser.page.evaluate(() => window.actions.includes('gpt')), false);
  });
}

test('hover submenus and a compact toolbar label use the checked model option for confirmation', async t => {
  const browser = await fixture(t, { hover: true, compact: true });
  if (!browser) return;
  await browser.selectModel();
  assert.equal(await browser.page.getByRole('button', { name: 'GPT-5.6 Sol', exact: true }).count(), 1);
  assert.deepEqual(await browser.page.evaluate(() => window.actions), ['picker', 'gpt', 'target', 'picker', 'gpt']);
  assert.equal(await browser.page.getByRole('menu').count(), 0);
});

for (const scenario of ['missing', 'ignore']) {
  test(`an unavailable or unconfirmed model (${scenario}) stops before input and Send`, async t => {
    const browser = await fixture(t, { [scenario]: true });
    if (!browser) return;
    // Text elsewhere in the page must not be mistaken for a selectable/selected model.
    await browser.page.evaluate(() => {
      const text = document.createElement('p'); text.textContent = 'GPT-5.6 Sol Think deeper'; document.body.append(text);
    });
    await assert.rejects(browser.ask('Do not send with Auto.', { fresh: true }), error => {
      assert.match(error.message, /Could not confirm GPT-5.6 Sol Think deeper/);
      assert.match(error.message, scenario === 'missing' ? /step=gpt-submenu/ : /step=verify/);
      assert.match(error.message, /Picker=.*options=/);
      assert.doesNotMatch(error.message, /Do not send with Auto/);
      assert.ok(error.message.length < 1500);
      return true;
    });
    assert.deepEqual(await browser.page.evaluate(() => window.sent), []);
    assert.deepEqual(await browser.page.evaluate(() => window.inputSelections), []);
  });
}

test('cancellation does not change the selected model', async t => {
  const browser = await fixture(t);
  if (!browser) return;
  await assert.rejects(browser.selectModel({ signal: AbortSignal.abort() }), { name: 'AbortError' });
  assert.deepEqual(await browser.page.evaluate(() => window.actions), []);
});

for (const config of [
  { referenceMarkup: true },
  { referenceMarkup: true, compact: true },
  { labelledPicker: true },
  { compact: true, checkmark: true, keepMenu: true }
]) {
  test(`M365 model controls work with ${Object.keys(config).join(', ')}`, async t => {
    const browser = await fixture(t, config);
    if (!browser) return;
    await browser.selectModel();
    assert.equal(await browser.page.evaluate(() => window.actions.filter(action => action === 'target').length), 1);
    assert.equal(await browser.page.getByRole('menu').count(), 0);
  });
}

test('a hidden checkmark is not accepted as selection and errors do not enter a draft', async t => {
  const browser = await fixture(t, { checkmark: true, compact: true, keepMenu: true, ignore: true });
  if (!browser) return;
  await assert.rejects(browser.ask('Do not send with an unselected model.', { fresh: true }), /step=verify/);
  assert.deepEqual(await browser.page.evaluate(() => window.sent), []);
  assert.deepEqual(await browser.page.evaluate(() => window.inputSelections), []);
});

test('input read-back follows a rebuilt Message Copilot textbox without refilling or resending', async t => {
  const browser = await fixture(t, { selected: true, rebuildEditor: true, roleEditor: true });
  if (!browser) return;
  const text = 'Inspect C:/project/example.js.\n保留这一行。';
  const statuses = [];
  assert.equal((await browser.ask(text, { fresh: true, onStatus: status => statuses.push(status) })).completed, true);
  assert.deepEqual(await browser.page.evaluate(() => window.sent), [{ text, selected: true }]);
  // One multiline fill can emit several input events; the retry status identifies a second attempt.
  assert.equal(statuses.some(status => status.startsWith('Retrying input')), false);
});

test('input read-back survives an editor remount longer than the five-second locator timeout', async t => {
  const browser = await fixture(t, { selected: true, rebuildEditor: true, rebuildDelay: 5500 });
  if (!browser) return;
  browser.context.setDefaultTimeout(5000);
  const text = 'Send the existing draft exactly once.';
  assert.equal((await browser.ask(text, { fresh: true })).completed, true);
  assert.deepEqual(await browser.page.evaluate(() => window.sent), [{ text, selected: true }]);
  assert.deepEqual(await browser.page.evaluate(() => window.inputSelections), [true]);
});

test('input read-back stops without refilling or sending if the editor never returns', async t => {
  const browser = await fixture(t, { selected: true, rebuildEditor: true, rebuildDelay: 60000 });
  if (!browser) return;
  const statuses = [];
  await assert.rejects(browser.ask('Keep this draft.\nDo not resend it.', { fresh: true, onStatus: status => statuses.push(status) }),
    /step=verify-input, timeout=30000ms.*Send was not clicked/);
  assert.deepEqual(await browser.page.evaluate(() => window.sent), []);
  assert.equal(statuses.some(status => status.startsWith('Retrying input')), false);
  assert.equal(browser.busy, false);
});

test('multiple matching editors stop input instead of choosing the first one', async t => {
  const browser = await fixture(t, { selected: true });
  if (!browser) return;
  await browser.page.evaluate(() => {
    const second = document.getElementById('m365-chat-editor-target-element').cloneNode(true);
    second.removeAttribute('id');
    document.body.append(second);
  });
  await assert.rejects(browser.ask('Do not enter into an ambiguous editor.', { fresh: true }), /Multiple Copilot input boxes/);
  assert.deepEqual(await browser.page.evaluate(() => window.sent), []);
  assert.deepEqual(await browser.page.evaluate(() => window.inputSelections), []);
});
