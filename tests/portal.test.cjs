const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createRuntime } = require('../tools/gas-local.cjs');
const original = fs.readFileSync(path.join(__dirname, 'fixtures/original.html'), 'utf8');
const index = fs.readFileSync(path.join(__dirname, '../apps-script/Index.html'), 'utf8');
const plain = value => JSON.parse(JSON.stringify(value));
const allLinks = portal => portal.sections.flatMap(section => section.groups.flatMap(group => group.links));

test('添付HTMLの全75件のURL、名前、改行、アイコン、説明とCSSを保持する', () => {
  const { seed } = createRuntime();
  const anchors = [...original.matchAll(/<a\s+([^>]*class="link-card"[^>]*)>([\s\S]*?)<\/a>/g)];
  assert.equal(anchors.length, 75);
  assert.equal(seed.links.length, anchors.length);
  anchors.forEach((anchor, i) => {
    assert.equal(seed.links[i][5], anchor[1].match(/href="([^"]*)"/)[1]);
    const title = anchor[2].match(/<span class="card-title">([\s\S]*?)<\/span>/)[1].replace(/<br\s*\/?\s*>/g, '\n');
    assert.equal(seed.links[i][4], title);
    assert.equal(seed.links[i][6], anchor[2].match(/data-feather="([^"]*)"/)[1]);
    assert.equal(seed.links[i][7], anchor[1].match(/data-title="([^"]*)"/)?.[1] || '');
  });
  assert.equal(index.match(/<style>([\s\S]*?)<\/style>/)[1], original.match(/<style>([\s\S]*?)<\/style>/)[1]);
});

test('5セクション・10小分類・75カード・過去リンクを描画する', () => {
  const { context, errors, seed } = createRuntime();
  const portal = context.readPortal_();
  assert.deepEqual(plain(portal.sections.map(section => section.name)), ['庶務関連申請', '購買関連', '営業アシスタント', '財務経理関連', 'その他']);
  assert.equal(allLinks(portal).length, 75);
  const output = context.doGet();
  assert.equal(errors.length, 0);
  assert.equal(output.title, '社内リンク集');
  assert.equal((output.html.match(/class="link-card"/g) || []).length, 75);
  assert.equal((output.html.match(/class="subsection"/g) || []).length, 10);
  assert.equal((output.html.match(/<br>/g) || []).length, 10);
  assert.equal(portal.pastLink.url, seed.settings.find(row => row[0] === '過去リンクURL')[1]);
  assert.ok(output.html.includes('target="_blank" rel="noopener noreferrer"'));
});

test('シート編集でURL・表示名・追加・削除・非表示・表示順が次の読込に反映される', () => {
  const { context, sheets, seed } = createRuntime();
  const rows = sheets.get('リンク').rows;
  rows[1][4] = '更新したリンク';
  rows[1][5] = 'https://example.com/updated?x=1&y=2#fragment';
  rows[2][9] = false;
  rows.splice(3, 1);
  rows.push(['section-new', '新分類', 'star', '', '追加リンク', 'https://example.com/new', 'plus', '新しい説明', 0, true]);
  const portal = context.readPortal_();
  const links = allLinks(portal);
  assert.equal(portal.sections[0].name, '新分類');
  assert.equal(links.length, 74);
  assert.equal(links.find(link => link.name === '更新したリンク').url, rows[1][5]);
  assert.ok(!links.some(link => link.name === seed.links[1][4] || link.name === seed.links[2][4]));
  assert.ok(context.doGet().html.includes('https://example.com/updated?x=1&amp;y=2#fragment'));
});

test('表示順が同じ場合はシートの行順を保ち、空行は無視する', () => {
  const { context, sheets } = createRuntime();
  const sheet = sheets.get('リンク');
  sheet.rows[1][8] = 1;
  sheet.rows[2][8] = 1;
  sheet.rows.push(Array(10).fill(''));
  const links = allLinks(context.readPortal_());
  assert.equal(links.length, 75);
  assert.equal(links[0].name, sheet.rows[1][4]);
  assert.equal(links[1].name, sheet.rows[2][4]);
});

test('初期設定は75件と設定を作成し、再実行で管理者の編集・削除を上書きしない', () => {
  const { context, sheets, properties } = createRuntime({ seedSheets: false, initialized: false });
  context.setup_();
  assert.equal(properties.get('LINKS_SPREADSHEET_ID'), 'local-sheet');
  assert.equal(sheets.get('リンク').rows.length, 76);
  const rows = sheets.get('リンク').rows;
  rows[1][5] = 'https://example.com/keep';
  rows[2][9] = false;
  rows.pop();
  sheets.get('設定').rows[1][1] = '変更したタイトル';
  const before = JSON.stringify([...sheets]);
  context.setup_();
  assert.equal(JSON.stringify([...sheets]), before);
  assert.equal(allLinks(context.readPortal_()).length, 73);
});

test('設定シートからタイトルと過去リンクを変更・非表示にできる', () => {
  const { context, sheets } = createRuntime();
  const rows = sheets.get('設定').rows;
  rows.find(row => row[0] === 'ページタイトル')[1] = '部内ポータル';
  rows.find(row => row[0] === '過去リンクURL')[1] = 'https://example.com/archive';
  rows.find(row => row[0] === '過去リンク名')[1] = '旧ページ';
  let portal = context.readPortal_();
  assert.equal(portal.title, '部内ポータル');
  assert.equal(portal.pastLink.url, 'https://example.com/archive');
  assert.equal(portal.pastLink.name, '旧ページ');
  rows.find(row => row[0] === '過去リンク表示')[1] = 'FALSE';
  assert.equal(context.readPortal_().pastLink, null);
});

test('HTML入力は文字列として描画し、実行可能なURLは拒否する', () => {
  const { context, sheets } = createRuntime();
  sheets.get('リンク').rows[1][4] = '<img src=x onerror=alert(1)>\n次の行';
  sheets.get('リンク').rows[1][7] = '"><script>alert(1)</script>';
  const html = context.doGet().html;
  assert.ok(html.includes('&lt;img src=x onerror=alert(1)&gt;<br>次の行'));
  assert.ok(!html.includes('<script>alert(1)</script>'));
  for (const value of ['javascript:alert(1)', 'data:text/html,x', '//example.com', 'https://', 'https://good.example/\nbad']) {
    sheets.get('リンク').rows[1][5] = value;
    assert.throws(() => context.readPortal_(), /URL/);
  }
});

test('見出し破損、分類の矛盾、無効な表示・順序を検出する', () => {
  const cases = [
    [rows => { rows[0][5] = '別の見出し'; }, /見出し/],
    [rows => { rows[2][1] = '矛盾したセクション'; }, /統一/],
    [rows => { rows[1][9] = 'maybe'; }, /TRUE/],
    [rows => { rows[1][8] = '後で'; }, /表示順/]
  ];
  for (const [change, message] of cases) {
    const { context, sheets } = createRuntime();
    change(sheets.get('リンク').rows);
    assert.throws(() => context.readPortal_(), message);
  }
});

test('未設定・シート欠落を管理者ログに残し、利用者にエラーを表示する', () => {
  const runtime = createRuntime({ initialized: false });
  assert.throws(() => runtime.context.readPortal_(), /setup_/);
  const output = runtime.context.doGet();
  assert.equal(runtime.errors.length, 1);
  assert.ok(output.html.includes('role="alert"'));
  assert.ok(!output.html.includes('class="link-card"'));
  const missing = createRuntime();
  missing.sheets.delete('リンク');
  assert.throws(() => missing.context.readPortal_(), /シート/);
});

test('全件非表示を空の状態として表示し、非表示行の未入力URLは許容する', () => {
  const { context, sheets } = createRuntime();
  for (const row of sheets.get('リンク').rows.slice(1)) { row[9] = false; row[5] = ''; }
  assert.equal(context.readPortal_().sections.length, 0);
  assert.ok(context.doGet().html.includes('表示するリンクはありません。'));
});

test('検索は日本語・大文字小文字・空白・解除に対応する', () => {
  const script = createRuntime().context.doGet().html.match(/<script>([\s\S]*?)<\/script>/)[1];
  const cards = ['セキュリティカード', 'AIライセンス 経費使用申請', 'Oracle割引受験チケット申し込み'].map(textContent => ({ textContent, style: {} }));
  const browser = vm.createContext({ document: { querySelectorAll: () => cards } });
  vm.runInContext(script, browser); // CDN が使えなくても検索は動く。
  browser.filterLinks('  ai  ');
  assert.deepEqual(cards.map(card => card.style.display), ['none', 'flex', 'none']);
  browser.filterLinks('カード');
  assert.deepEqual(cards.map(card => card.style.display), ['flex', 'none', 'none']);
  browser.filterLinks('');
  assert.ok(cards.every(card => card.style.display === 'flex'));
});
