const LINK_SHEET_ = 'リンク';
const SETTINGS_SHEET_ = '設定';
const SPREADSHEET_PROPERTY_ = 'LINKS_SPREADSHEET_ID';
const LINK_HEADERS_ = [
  'セクションID', 'セクション名', 'セクションアイコン', '小分類',
  'リンク名', 'URL', 'アイコン', '説明', '表示順', '表示'
];

/** エディタで選択する実行用関数。シートを開いた管理画面からのみ実行する。 */
function setup() {
  SpreadsheetApp.getUi();
  return setup_();
}

function diagnose() {
  SpreadsheetApp.getUi();
  return diagnose_();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('リンク集')
    .addItem('初期設定（既存データは保持）', 'setup')
    .addItem('診断（実行ログに結果を表示）', 'diagnose')
    .addToUi();
}

/** スプレッドシートに紐づくエディタで一度実行。末尾 _ はブラウザからの呼出しを防ぐ。 */
function setup_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('管理用スプレッドシートの「拡張機能 → Apps Script」から実行してください。');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const seed = initialData_();
    initializeSheet_(spreadsheet, LINK_SHEET_, [LINK_HEADERS_].concat(seed.links), true);
    initializeSheet_(spreadsheet, SETTINGS_SHEET_, [['項目', '値']].concat(seed.settings), false);
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().setProperty(SPREADSHEET_PROPERTY_, spreadsheet.getId());
    const result = diagnose_();
    console.log('初期設定完了。ウェブアプリを再読み込みしてください。');
    return result;
  } finally {
    lock.releaseLock();
  }
}

/** エディタまたはスプレッドシートのメニューから実行。ID・URL・セル値は出力しない。 */
function diagnose_() {
  const portal = readPortal_();
  const result = {
    sections: portal.sections.length,
    links: portal.sections.reduce((count, section) =>
      count + section.groups.reduce((sum, group) => sum + group.links.length, 0), 0)
  };
  console.log('診断OK: ' + result.links + '件のリンク、' + result.sections + 'セクションを読み込めました。');
  return result;
}

function initializeSheet_(spreadsheet, name, values, isLinks) {
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  if (sheet.getLastRow() > 0) return; // 再実行してもユーザーの編集・追加・削除を戻さない。
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, values[0].length).setFontWeight('bold').setBackground('#fff0d9');
  if (isLinks) {
    sheet.getRange(2, 10, values.length - 1, 1).insertCheckboxes();
    sheet.getRange(2, 10, values.length - 1, 1).setValues(values.slice(1).map(row => [row[9]]));
    sheet.setColumnWidth(5, 300);
    sheet.setColumnWidth(6, 420);
    sheet.setColumnWidth(8, 280);
  } else {
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 500);
  }
}

function doGet() {
  let portal;
  let errorMessage = '';
  try {
    portal = readPortal_();
  } catch (error) {
    console.error(error); // 詳細は管理者の実行ログに残す。
    portal = { title: '各種リンク集', browserTitle: '社内リンク集', sections: [], pastLink: null };
    errorMessage = error && error.portalCode === 'SETUP_REQUIRED'
      ? '初期設定が完了していません。管理用スプレッドシートの「拡張機能 → Apps Script」で setup を一度実行し、権限を承認してください。'
      : 'リンク集を読み込めませんでした。管理者は Apps Script で diagnose を実行し、表示されたエラーを確認してください。';
  }
  return HtmlService.createHtmlOutput(renderPage_(portal, errorMessage)).setTitle(portal.browserTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/** 毎回シートを読み込むため、編集は次のページ再読込で反映される。 */
function readPortal_() {
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_PROPERTY_);
  if (!id) {
    const error = new Error('初期設定が完了していません。管理用スプレッドシートの「拡張機能 → Apps Script」で setup を実行してください。');
    error.portalCode = 'SETUP_REQUIRED';
    throw error;
  }
  const spreadsheet = SpreadsheetApp.openById(id);
  const rows = readSheet_(spreadsheet, LINK_SHEET_, LINK_HEADERS_);
  const settings = readSettings_(spreadsheet);
  const links = [];
  rows.forEach((row, index) => {
    const line = index + 2;
    if (row.every(cell => text_(cell) === '')) return;
    if (!boolean_(row[9], true, 'リンク ' + line + '行目の表示')) return;
    const sectionId = text_(row[0]);
    const sectionName = text_(row[1]);
    const name = text_(row[4]);
    if (!/^[a-zA-Z][a-zA-Z0-9_-]*$/.test(sectionId) || !sectionName || !name) {
      throw new Error('リンク ' + line + '行目: セクションID（半角英数字・ハイフン・アンダースコア）、セクション名、リンク名を確認してください。');
    }
    const orderText = text_(row[8]);
    const order = orderText === '' ? line : Number(orderText);
    if (!Number.isFinite(order)) throw new Error('リンク ' + line + '行目: 表示順は数値にしてください。');
    links.push({
      sectionId, sectionName,
      sectionIcon: icon_(row[2], 'clipboard'),
      group: text_(row[3]), name,
      url: url_(row[5], 'リンク ' + line + '行目のURL'),
      icon: icon_(row[6], 'link'), description: text_(row[7]), order, line
    });
  });
  links.sort((a, b) => a.order - b.order || a.line - b.line);
  const sections = [];
  const sectionMap = new Map();
  links.forEach(link => {
    let section = sectionMap.get(link.sectionId);
    if (!section) {
      section = { id: link.sectionId, name: link.sectionName, icon: link.sectionIcon, groups: [] };
      sectionMap.set(link.sectionId, section);
      sections.push(section);
    } else if (section.name !== link.sectionName || section.icon !== link.sectionIcon) {
      throw new Error('リンク ' + link.line + '行目: 同じセクションIDの名前・アイコンを統一してください。');
    }
    let group = section.groups.find(item => item.name === link.group);
    if (!group) {
      group = { name: link.group, links: [] };
      section.groups.push(group);
    }
    group.links.push(link);
  });
  const pastVisible = boolean_(settings['過去リンク表示'], true, '設定: 過去リンク表示');
  const pastUrl = text_(settings['過去リンクURL']);
  return {
    title: text_(settings['ページタイトル']) || '各種リンク集',
    browserTitle: text_(settings['ブラウザータイトル']) || '社内リンク集',
    sections,
    pastLink: pastVisible && pastUrl ? {
      name: text_(settings['過去リンク名']) || '過去リンク',
      url: url_(pastUrl, '設定: 過去リンクURL')
    } : null
  };
}

function readSheet_(spreadsheet, name, headers) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet || sheet.getLastRow() === 0) throw new Error(name + ' シートがありません。setup を実行してください。');
  const values = sheet.getDataRange().getDisplayValues();
  const actual = values.shift().map(text_);
  if (headers.some((header, index) => actual[index] !== header)) {
    throw new Error(name + ' シートの見出しを確認してください: ' + headers.join(' / '));
  }
  return values;
}

function readSettings_(spreadsheet) {
  const settings = Object.create(null);
  readSheet_(spreadsheet, SETTINGS_SHEET_, ['項目', '値']).forEach(row => {
    const key = text_(row[0]);
    if (!key) return;
    if (Object.prototype.hasOwnProperty.call(settings, key)) throw new Error('設定の項目が重複しています: ' + key);
    settings[key] = row[1];
  });
  return settings;
}

function text_(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function boolean_(value, fallback, label) {
  const text = text_(value).toLowerCase();
  if (text === '') return fallback;
  if (['true', '1', '表示'].includes(text)) return true;
  if (['false', '0', '非表示'].includes(text)) return false;
  throw new Error(label + ': TRUE / FALSE またはチェックボックスを使用してください。');
}

function icon_(value, fallback) {
  const name = text_(value);
  return /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(name) ? name : fallback;
}

function url_(value, label) {
  const url = text_(value);
  // 表示時にHTMLをエスケープする。実行可能なスキームは禁止。
  if (/[\u0000-\u001f\u007f]/.test(url) || !/^https?:\/\/[^\s/?#<>"'`]+(?:[/?#][\s\S]*)?$/i.test(url)) {
    throw new Error(label + ': http:// または https:// で始まるURLを入力してください。');
  }
  return url;
}

/** 最初に提供されたHTMLを、そのままIndex.htmlに保存して使用する。 */
function sourceHtml_() {
  return HtmlService.createHtmlOutputFromFile('Index').getContent();
}

function initialData_() {
  const html = sourceHtml_();
  const links = [];
  const sections = Array.from(html.matchAll(/<section\b([^>]*)>([\s\S]*?)<\/section>/gi));
  sections.forEach(section => {
    const heading = (section[2].match(/<h2\b[^>]*>([\s\S]*?)<\/h2>/i) || [])[1] || '';
    const sectionId = attribute_(section[1], 'id');
    const sectionName = plainText_(heading);
    const sectionIcon = attribute_((heading.match(/<i\b([^>]*)>/i) || [])[1] || '', 'data-feather') || 'clipboard';
    section[2].split(/<div\s+class=["']subsection["']\s*>/i).forEach(group => {
      const groupName = plainText_((group.match(/<h3\b[^>]*>([\s\S]*?)<\/h3>/i) || [])[1] || '');
      Array.from(group.matchAll(/<a\b([^>]*)>([\s\S]*?)<\/a>/gi)).forEach(anchor => {
        if (!attribute_(anchor[1], 'class').split(/\s+/).includes('link-card')) return;
        const title = (anchor[2].match(/<span\b[^>]*class=["']card-title["'][^>]*>([\s\S]*?)<\/span>/i) || [])[1];
        if (title === undefined) throw new Error('元HTMLにカード名がありません。');
        const iconTag = (anchor[2].match(/<i\b([^>]*)>/i) || [])[1] || '';
        links.push([sectionId, sectionName, sectionIcon, groupName, plainText_(title),
          attribute_(anchor[1], 'href'), attribute_(iconTag, 'data-feather') || 'link',
          attribute_(anchor[1], 'data-title'), links.length + 1, true]);
      });
    });
  });
  if (!links.length) throw new Error('Index.htmlに最初のHTML全文を貼り付けてください。');
  const past = Array.from(html.matchAll(/<a\b([^>]*)>([\s\S]*?)<\/a>/gi))
    .find(anchor => attribute_(anchor[1], 'class').split(/\s+/).includes('fixed-link'));
  return { links, settings: [
    ['ページタイトル', plainText_((html.match(/<h1\b[^>]*>([\s\S]*?)<\/h1>/i) || [])[1] || '各種リンク集')],
    ['ブラウザータイトル', plainText_((html.match(/<title>([\s\S]*?)<\/title>/i) || [])[1] || '社内リンク集')],
    ['過去リンク名', past ? plainText_(past[2]) : '過去リンク'],
    ['過去リンクURL', past ? attribute_(past[1], 'href') : ''],
    ['過去リンク表示', 'TRUE']
  ] };
}

function attribute_(tag, name) {
  const match = tag.match(new RegExp('(?:^|\\s)' + name + '\\s*=\\s*(["\'])([\\s\\S]*?)\\1', 'i'));
  return match ? decodeHtml_(match[2]) : '';
}

function plainText_(html) {
  return decodeHtml_(html.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>/g, '')).trim();
}

function decodeHtml_(text) {
  const entities = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0' };
  return text.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos|nbsp);/gi, (match, entity) => {
    const key = entity.toLowerCase();
    if (key[0] !== '#') return entities[key];
    const code = key[1] === 'x' ? parseInt(key.slice(2), 16) : Number(key.slice(1));
    return code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : match;
  });
}

function escapeHtml_(value) {
  const entities = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  return String(value).replace(/[&<>"']/g, character => entities[character]);
}

function renderPage_(portal, errorMessage) {
  let html = sourceHtml_();
  const blocks = Array.from(html.matchAll(/<section\b[^>]*>[\s\S]*?<\/section>/gi));
  if (!blocks.length) throw new Error('Index.htmlには、最初のHTML全文を保存してください。');
  let content;
  if (errorMessage) {
    content = '<p class="section" role="alert">' + escapeHtml_(errorMessage) + '</p>';
  } else if (!portal.sections.length) {
    content = '<p class="section" role="status">表示するリンクはありません。</p>';
  } else {
    content = portal.sections.map(section => {
      const groups = section.groups.map(group => {
        const cards = group.links.map(link => {
          const tooltip = link.description ? ' data-title="' + escapeHtml_(link.description) + '"' : '';
          return '<a class="link-card" href="' + escapeHtml_(link.url) + '" target="_blank" rel="noopener noreferrer"' + tooltip + '>' +
            '<i class="' + (group.name ? 'card-icon' : '') + '" data-feather="' + escapeHtml_(link.icon) + '"></i>' +
            '<span class="card-title">' + link.name.split(/\r?\n/).map(escapeHtml_).join('<br>') + '</span></a>';
        }).join('\n');
        const grid = '<div class="links-grid">' + cards + '</div>';
        return group.name ? '<div class="subsection"><h3 class="subsection-title">' + escapeHtml_(group.name) + '</h3>' + grid + '</div>' : grid;
      }).join('\n');
      return '<section class="section" id="' + escapeHtml_(section.id) + '"><h2 class="section-title"><i data-feather="' +
        escapeHtml_(section.icon) + '"></i> ' + escapeHtml_(section.name) + '</h2>' + groups + '</section>';
    }).join('\n');
  }
  const last = blocks[blocks.length - 1];
  html = html.slice(0, blocks[0].index) + content + html.slice(last.index + last[0].length);
  html = html.replace(/(<h1\b[^>]*>)[\s\S]*?(<\/h1>)/i, (match, open, close) => open + escapeHtml_(portal.title) + close);
  html = html.replace(/<title>[\s\S]*?<\/title>/i, () => '<title>' + escapeHtml_(portal.browserTitle) + '</title>');
  const past = portal.pastLink;
  html = html.replace(/<a\b[^>]*class=["']fixed-link["'][^>]*>[\s\S]*?<\/a>/i, () => past ?
    '<a class="fixed-link" href="' + escapeHtml_(past.url) + '" target="_blank" rel="noopener noreferrer">' + escapeHtml_(past.name) + '</a>' : '');
  html = html.replace('onkeyup="filterLinks(this.value)"', 'oninput="filterLinks(this.value)"');
  html = html.replace('feather.replace();', "if (typeof feather !== 'undefined') feather.replace();");
  return html;
}
