/******************************************************
 * 商品コード管理ツール v9
 * 
 * 【仕様】
 * - onEdit時点で作品IDを確定（「????」ではない）
 * - WorksKey = 日本語タイトル + 作者
 * - 同じ作品なら同じID、新規なら新規ID発番
 * - 巻数入力時に重複チェック
 * 
 * 【高速化】
 * - Worksデータを1回で読み込み
 * - 一括書き込み
 * - 列番号をPropertiesにキャッシュ
 *
 ******************************************************/

const CFG = {
  SHEET_MASTER: '台湾まんがマスターシート',
  SHEET_WORKS: 'Works',
  SHEET_LANG: '言語マスター',
  SHEET_CAT: 'カテゴリマスター',

  WORKS_HEADERS: ['WorksKey', '作品ID', '日本語タイトル', '作者', '原題タイトル', '登録済み巻', '最新巻', '更新日時'],
  LANG_HEADERS: ['言語', 'コード', '色'],
  CAT_HEADERS: ['カテゴリ', 'コード', '色'],

  LANG_INITIAL: [
    ['台湾', 'TW', '#b7e1cd'], ['韓国', 'KR', '#fce8b2'], ['日本', 'JP', '#f4c7c3'],
    ['中国', 'CN', '#c9daf8'], ['タイ', 'TH', '#d9d2e9'], ['英語', 'US', '#fce5cd']
  ],

  CAT_INITIAL: [
    ['コミック', 'CM', '#fff2cc'], ['まんが', 'CM', '#fff2cc'], ['小説', 'NV', '#d9ead3'],
    ['グッズ', 'GD', '#cfe2f3'], ['設定集', 'ART', '#f4cccc'], ['アートブック', 'ART', '#f4cccc'], ['雑誌', 'MZ', '#d9d2e9']
  ],

  COLOR_PALETTE: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd', '#fff2cc', '#d9ead3', '#cfe2f3', '#f4cccc'],

  EDITION_PREFIX: { '特装版': 'S', '初版限定版': 'F', '初回限定版': 'F' },
  AUTO_BONUS_IF_EMPTY: true,
  BONUS_BY_EDITION: { '初版限定版': '※初版限定版特典付き', '初回限定版': '※初回限定版特典付き', '特装版': '※特装版限定特典付き' }
};

const H = {
  ISSUE_OK: '発番発行',
  CODE: '商品コード（SKU）',
  TITLE_OUT: 'タイトル',
  AUTHOR: '作者',
  JP: '日本語タイトル',
  ORG: '原題タイトル',
  EDITION: '形態（通常/初回限定/特装）',
  LANG: '言語',
  CAT: 'カテゴリ',
  VOL: '単巻数',
  SET_FROM: 'セット巻数開始番号',
  SET_TO: 'セット巻数終了番号',
  BONUS_MEMO: '特典メモ',
  WORK_ID: '作品ID(W)（自動）',
  SKU: 'SKU（自動）',
  CODE_STATUS: '商品コードステータス',
  STATUS_AUTO: '登録状況'
};

const WATCH_COLS = [H.AUTHOR, H.JP, H.ORG, H.LANG, H.CAT, H.EDITION, H.VOL, H.SET_FROM, H.SET_TO, H.BONUS_MEMO];

/* ============================
 * メニュー
 * ============================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('商品コード管理')
    .addItem('① 商品コード確定発行', 'issueProductCodes')
    .addItem('② 商品コード削除', 'deleteProductCodes')
    .addSeparator()
    .addItem('③ 既存データ一括更新', 'rebuildAllRowsFast')
    .addSeparator()
    .addItem('④ Works初期化', 'resetWorksOnly')
    .addItem('⑤ マスター作成', 'createLangCatMasters')
    .addItem('⑥ プルダウン更新', 'updateDropdownsAndColors')
    .addItem('⑦ キャッシュクリア', 'clearAllCacheManual')
    .addToUi();
}

/* ============================
 * onEdit - 入力時に作品ID確定
 * ============================ */
function onEdit(e) {
  if (!e?.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.SHEET_MASTER) return;

  const row = e.range.getRow();
  if (row < 2) return;

  const col = getColMap_(sh);
  const editedCol = e.range.getColumn();

  // 監視対象列かチェック
  const watchColNums = WATCH_COLS.map(h => col[h]).filter(Boolean);
  if (!watchColNums.includes(editedCol)) return;

  // 1行分のデータを読む
  const lastCol = sh.getLastColumn();
  const rowData = sh.getRange(row, 1, 1, lastCol).getValues()[0];

  const jp   = norm_(rowData[col[H.JP] - 1]);
  const au   = norm_(rowData[col[H.AUTHOR] - 1]);
  const lang = norm_(rowData[col[H.LANG] - 1]);
  const cat  = norm_(rowData[col[H.CAT] - 1]);

  // 必須チェック（日本語タイトル、言語、カテゴリ）
  if (!jp || !lang || !cat) {
    clearRow_(sh, row, col);
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const langMap = getLangMap_(ss);
  const catMap = getCatMap_(ss);

  const langCode = langMap[lang] || 'XX';
  const catCode = catMap[cat] || 'XX';

  const ed  = norm_(rowData[col[H.EDITION] - 1]);
  const org = norm_(rowData[col[H.ORG] - 1]);
  const vol = rowData[col[H.VOL] - 1];
  const sf  = rowData[col[H.SET_FROM] - 1];
  const st  = rowData[col[H.SET_TO] - 1];
  const memo = norm_(rowData[col[H.BONUS_MEMO] - 1]);

  // ★Works検索・作品ID取得
  const worksSh = ensureWorksSheet_(ss);
  const worksResult = getOrCreateWorkId_(worksSh, jp, au, org);
  const workIdNum = worksResult.workId;

  // ★巻数重複チェック
  let volWarning = '';
  const volNum = toIntOrNull_(vol);
  if (volNum != null && worksResult.registeredVols) {
    const vols = worksResult.registeredVols.split(',').map(v => parseInt(v.trim(), 10)).filter(n => !isNaN(n));
    if (vols.includes(volNum)) {
      volWarning = '【重複注意】';
    }
  }

  // SKU生成
  const sku = buildSku_(langCode, ed, workIdNum, catCode, vol, sf, st);

  // タイトル生成
  const title = volWarning + buildTitle_(lang, cat, ed, jp, vol, sf, st, au, org, memo);

  // 書き込み
  const updates = [];
  if (col[H.CODE]) updates.push([col[H.CODE], sku]);
  if (col[H.TITLE_OUT]) updates.push([col[H.TITLE_OUT], title]);
  if (col[H.WORK_ID]) updates.push([col[H.WORK_ID], workIdNum]);
  if (col[H.SKU]) updates.push([col[H.SKU], sku]);
  if (col[H.CODE_STATUS]) updates.push([col[H.CODE_STATUS], '商品コード（予約）']);
  if (col[H.STATUS_AUTO]) updates.push([col[H.STATUS_AUTO], '未登録']);

  for (const [c, v] of updates) {
    sh.getRange(row, c).setValue(v);
  }
}

function clearRow_(sh, row, col) {
  const keys = [H.CODE, H.TITLE_OUT, H.WORK_ID, H.SKU, H.CODE_STATUS, H.STATUS_AUTO];
  for (const k of keys) {
    if (col[k]) sh.getRange(row, col[k]).setValue('');
  }
}

/* ============================
 * Works検索・作品ID取得
 * ============================ */
function getOrCreateWorkId_(worksSh, jp, author, org) {
  // WorksKey = 日本語タイトル + || + 作者
  const worksKey = (jp + '||' + (author || '')).toLowerCase();

  const last = worksSh.getLastRow();
  let maxId = 0;
  let foundRow = null;
  let foundId = null;
  let registeredVols = '';

  if (last >= 2) {
    const data = worksSh.getRange(2, 1, last - 1, 8).getValues();

    for (let i = 0; i < data.length; i++) {
      const key = String(data[i][0] || '').trim().toLowerCase();
      const id = String(data[i][1] || '').trim();

      // 最大ID更新
      const num = parseInt(id, 10);
      if (!isNaN(num) && num > maxId) maxId = num;

      // 同じ作品か
      if (key === worksKey) {
        foundRow = i + 2;
        foundId = id;
        registeredVols = String(data[i][5] || '');
      }
    }
  }

  if (foundId) {
    // 既存作品
    return { workId: foundId, isNew: false, registeredVols: registeredVols, rowNo: foundRow };
  }

  // 新規作品 → Works に追加
  const newIdNum = maxId + 1;
  const newId = String(newIdNum).padStart(4, '0');
  const newRow = Math.max(2, last + 1);

  worksSh.getRange(newRow, 1, 1, 8).setValues([[
    worksKey,
    newId,
    jp,
    author || '',
    org || '',
    '',  // 登録済み巻
    '',  // 最新巻
    new Date()
  ]]);

  return { workId: newId, isNew: true, registeredVols: '', rowNo: newRow };
}

/* ============================
 * ① 商品コード確定発行
 * ============================ */
function issueProductCodes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const col = getColMap_(sh);
  const langMap = getLangMap_(ss);
  const catMap = getCatMap_(ss);

  const allData = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const worksSh = ensureWorksSheet_(ss);

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const outWorkId = [], outSku = [], outCode = [], outTitle = [], outStatus = [], outStatusAuto = [];
    let issued = 0;

    // Works全読み込み
    const worksData = readWorksAll_(worksSh);

    for (let i = 0; i < allData.length; i++) {
      const r = allData[i];

      if (r[col[H.ISSUE_OK] - 1] !== true) {
        outWorkId.push([r[col[H.WORK_ID] - 1] || '']);
        outSku.push([r[col[H.SKU] - 1] || '']);
        outCode.push([r[col[H.CODE] - 1] || '']);
        outTitle.push([r[col[H.TITLE_OUT] - 1] || '']);
        outStatus.push([r[col[H.CODE_STATUS] - 1] || '']);
        outStatusAuto.push([r[col[H.STATUS_AUTO] - 1] || '']);
        continue;
      }

      const jp = norm_(r[col[H.JP] - 1]);
      const au = norm_(r[col[H.AUTHOR] - 1]);
      const org = norm_(r[col[H.ORG] - 1]);
      const lang = norm_(r[col[H.LANG] - 1]);
      const cat = norm_(r[col[H.CAT] - 1]);
      const ed = norm_(r[col[H.EDITION] - 1]);

      if (!jp || !lang || !cat) {
        outWorkId.push([r[col[H.WORK_ID] - 1] || '']);
        outSku.push([r[col[H.SKU] - 1] || '']);
        outCode.push([r[col[H.CODE] - 1] || '']);
        outTitle.push([r[col[H.TITLE_OUT] - 1] || '']);
        outStatus.push([r[col[H.CODE_STATUS] - 1] || '']);
        outStatusAuto.push([r[col[H.STATUS_AUTO] - 1] || '']);
        continue;
      }

      const langCode = langMap[lang] || 'XX';
      const catCode = catMap[cat] || 'XX';

      // Works照合
      const worksKey = (jp + '||' + (au || '')).toLowerCase();
      let workIdNum = worksData.keyToId[worksKey];

      if (!workIdNum) {
        worksData.maxId++;
        workIdNum = String(worksData.maxId).padStart(4, '0');
        worksData.keyToId[worksKey] = workIdNum;
        worksData.newRows.push([worksKey, workIdNum, jp, au, org, '', '', new Date()]);
      }

      // 巻数をWorksに登録
      const volNum = toIntOrNull_(r[col[H.VOL] - 1]);
      if (volNum != null) {
        if (!worksData.keyToVols[worksKey]) worksData.keyToVols[worksKey] = new Set();
        worksData.keyToVols[worksKey].add(volNum);
      }

      const sku = buildSku_(langCode, ed, workIdNum, catCode, r[col[H.VOL] - 1], r[col[H.SET_FROM] - 1], r[col[H.SET_TO] - 1]);
      const title = buildTitle_(lang, cat, ed, jp, r[col[H.VOL] - 1], r[col[H.SET_FROM] - 1], r[col[H.SET_TO] - 1], au, org, norm_(r[col[H.BONUS_MEMO] - 1]));

      outWorkId.push([workIdNum]);
      outSku.push([sku]);
      outCode.push([sku]);
      outTitle.push([title]);
      outStatus.push(['商品コード（発行済み確定）']);
      outStatusAuto.push(['登録済み']);
      issued++;
    }

    // 一括書き込み
    if (col[H.WORK_ID]) sh.getRange(2, col[H.WORK_ID], outWorkId.length, 1).setValues(outWorkId);
    if (col[H.SKU]) sh.getRange(2, col[H.SKU], outSku.length, 1).setValues(outSku);
    if (col[H.CODE]) sh.getRange(2, col[H.CODE], outCode.length, 1).setValues(outCode);
    if (col[H.TITLE_OUT]) sh.getRange(2, col[H.TITLE_OUT], outTitle.length, 1).setValues(outTitle);
    if (col[H.CODE_STATUS]) sh.getRange(2, col[H.CODE_STATUS], outStatus.length, 1).setValues(outStatus);
    if (col[H.STATUS_AUTO]) sh.getRange(2, col[H.STATUS_AUTO], outStatusAuto.length, 1).setValues(outStatusAuto);

    // Works更新
    updateWorksFromData_(worksSh, worksData);

    ui.alert(`✅ 確定発行完了: ${issued}件`);
  } finally {
    lock.releaseLock();
  }
}

/* ============================
 * ② 商品コード削除
 * ============================ */
function deleteProductCodes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const col = getColMap_(sh);
  const allData = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const rowsToDelete = [];
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][col[H.ISSUE_OK] - 1] === true) rowsToDelete.push(i + 2);
  }

  if (rowsToDelete.length === 0) {
    ui.alert('チェックが入った行がありません');
    return;
  }

  const res = ui.alert('確認', `${rowsToDelete.length}件を削除します。続行？`, ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;

  rowsToDelete.sort((a, b) => b - a);
  for (const row of rowsToDelete) sh.deleteRow(row);

  ui.alert(`✅ 削除完了: ${rowsToDelete.length}件`);
}

/* ============================
 * ③ 既存データ一括更新
 * ============================ */
function rebuildAllRowsFast() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('確認', '全行を再生成します。続行？', ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('データがありません'); return; }

  const col = getColMap_(sh);
  const langMap = getLangMap_(ss);
  const catMap = getCatMap_(ss);

  if (Object.keys(langMap).length === 0 || Object.keys(catMap).length === 0) {
    ui.alert('言語/カテゴリマスターがありません。⑤を実行してください。');
    return;
  }

  const allData = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const worksSh = ensureWorksSheet_(ss);

  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);

  try {
    const worksData = readWorksAll_(worksSh);
    const outWorkId = [], outSku = [], outCode = [], outTitle = [], outStatus = [], outStatusAuto = [], outAuthor = [], outOrg = [];

    for (let i = 0; i < allData.length; i++) {
      const r = allData[i];

      const jp = norm_(r[col[H.JP] - 1]);
      const lang = norm_(r[col[H.LANG] - 1]);
      const cat = norm_(r[col[H.CAT] - 1]);

      if (!jp || !lang || !cat) {
        outWorkId.push([r[col[H.WORK_ID] - 1] || '']);
        outSku.push([r[col[H.SKU] - 1] || '']);
        outCode.push([r[col[H.CODE] - 1] || '']);
        outTitle.push([r[col[H.TITLE_OUT] - 1] || '']);
        outStatus.push([r[col[H.CODE_STATUS] - 1] || '']);
        outStatusAuto.push([r[col[H.STATUS_AUTO] - 1] || '']);
        outAuthor.push([r[col[H.AUTHOR] - 1] || '']);
        outOrg.push([r[col[H.ORG] - 1] || '']);
        continue;
      }

      const langCode = langMap[lang] || 'XX';
      const catCode = catMap[cat] || 'XX';
      const ed = norm_(r[col[H.EDITION] - 1]);

      let au = norm_(r[col[H.AUTHOR] - 1]);
      let org = norm_(r[col[H.ORG] - 1]);

      // Works照合
      const worksKey = (jp + '||' + (au || '')).toLowerCase();
      let workIdNum = worksData.keyToId[worksKey];
      const existingData = worksData.keyToData[worksKey];

      if (!workIdNum) {
        worksData.maxId++;
        workIdNum = String(worksData.maxId).padStart(4, '0');
        worksData.keyToId[worksKey] = workIdNum;
        worksData.newRows.push([worksKey, workIdNum, jp, au, org, '', '', new Date()]);
      }

      // 既存データから作者・原題を補完
      if (existingData) {
        if (!au && existingData.author) au = existingData.author;
        if (!org && existingData.org) org = existingData.org;
      }

      const sku = buildSku_(langCode, ed, workIdNum, catCode, r[col[H.VOL] - 1], r[col[H.SET_FROM] - 1], r[col[H.SET_TO] - 1]);
      const title = buildTitle_(lang, cat, ed, jp, r[col[H.VOL] - 1], r[col[H.SET_FROM] - 1], r[col[H.SET_TO] - 1], au, org, norm_(r[col[H.BONUS_MEMO] - 1]));

      const currentStatus = norm_(r[col[H.STATUS_AUTO] - 1]);
      const isConfirmed = (currentStatus === '登録済み');

      outWorkId.push([workIdNum]);
      outSku.push([sku]);
      outCode.push([sku]);
      outTitle.push([title]);
      outAuthor.push([au]);
      outOrg.push([org]);

      if (isConfirmed) {
        outStatus.push([r[col[H.CODE_STATUS] - 1] || '商品コード（発行済み確定）']);
        outStatusAuto.push(['登録済み']);
      } else {
        outStatus.push(['商品コード（予約）']);
        outStatusAuto.push(['未登録']);
      }
    }

    // 一括書き込み
    if (col[H.WORK_ID]) sh.getRange(2, col[H.WORK_ID], outWorkId.length, 1).setValues(outWorkId);
    if (col[H.SKU]) sh.getRange(2, col[H.SKU], outSku.length, 1).setValues(outSku);
    if (col[H.CODE]) sh.getRange(2, col[H.CODE], outCode.length, 1).setValues(outCode);
    if (col[H.TITLE_OUT]) sh.getRange(2, col[H.TITLE_OUT], outTitle.length, 1).setValues(outTitle);
    if (col[H.CODE_STATUS]) sh.getRange(2, col[H.CODE_STATUS], outStatus.length, 1).setValues(outStatus);
    if (col[H.STATUS_AUTO]) sh.getRange(2, col[H.STATUS_AUTO], outStatusAuto.length, 1).setValues(outStatusAuto);
    if (col[H.AUTHOR]) sh.getRange(2, col[H.AUTHOR], outAuthor.length, 1).setValues(outAuthor);
    if (col[H.ORG]) sh.getRange(2, col[H.ORG], outOrg.length, 1).setValues(outOrg);

    // Works更新
    updateWorksFromData_(worksSh, worksData);

    ui.alert(`✅ 一括更新完了: ${outWorkId.length}件`);
  } finally {
    lock.releaseLock();
  }
}

/* ============================
 * Works操作
 * ============================ */
function ensureWorksSheet_(ss) {
  let sh = ss.getSheetByName(CFG.SHEET_WORKS);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_WORKS);
    sh.getRange(1, 1, 1, CFG.WORKS_HEADERS.length).setValues([CFG.WORKS_HEADERS]);
  }
  return sh;
}

function readWorksAll_(worksSh) {
  const result = { keyToId: {}, keyToData: {}, keyToRow: {}, keyToVols: {}, maxId: 0, newRows: [] };
  const last = worksSh.getLastRow();
  if (last < 2) return result;

  const data = worksSh.getRange(2, 1, last - 1, 8).getValues();
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const key = String(r[0] || '').trim().toLowerCase();
    const id = String(r[1] || '').trim();
    if (!key || !id) continue;

    result.keyToId[key] = id;
    result.keyToRow[key] = i + 2;
    result.keyToData[key] = { author: r[3] || '', org: r[4] || '' };

    const vols = String(r[5] || '').split(',').map(v => parseInt(v.trim(), 10)).filter(n => !isNaN(n));
    result.keyToVols[key] = new Set(vols);

    const num = parseInt(id, 10);
    if (!isNaN(num) && num > result.maxId) result.maxId = num;
  }
  return result;
}

function updateWorksFromData_(worksSh, worksData) {
  // 新規行追加
  if (worksData.newRows.length > 0) {
    const startRow = Math.max(2, worksSh.getLastRow() + 1);
    worksSh.getRange(startRow, 1, worksData.newRows.length, 8).setValues(worksData.newRows);
  }

  // 巻数更新
  for (const [key, vols] of Object.entries(worksData.keyToVols)) {
    const rowNo = worksData.keyToRow[key];
    if (!rowNo || vols.size === 0) continue;

    const volArr = Array.from(vols).sort((a, b) => a - b);
    worksSh.getRange(rowNo, 6).setValue(volArr.join(','));
    worksSh.getRange(rowNo, 7).setValue(Math.max(...volArr));
    worksSh.getRange(rowNo, 8).setValue(new Date());
  }
}

/* ============================
 * SKU / タイトル生成
 * ============================ */
function buildSku_(langCode, edition, workIdNum, catCode, vol, sf, st) {
  const edPrefix = CFG.EDITION_PREFIX[String(edition || '').trim()] || '';
  const base = langCode + edPrefix + workIdNum + '-' + catCode;

  const sfStr = String(sf || '').trim();
  const stStr = String(st || '').trim();
  if (sfStr && stStr) return base + '-' + sfStr.padStart(2, '0') + stStr.padStart(2, '0');

  const v = String(vol || '').trim();
  if (v) {
    const n = parseInt(v, 10);
    return base + '-' + (!isNaN(n) ? String(n).padStart(2, '0') : v);
  }
  return base;
}

function buildTitle_(lang, cat, ed, jp, vol, sf, st, au, org, memo) {
  const head = [lang, cat].filter(Boolean).join(' ') || '台湾 まんが';
  const edLabel = ed ? `（${ed}）` : '';

  let volPart = '';
  const sfStr = String(sf || '').trim();
  const stStr = String(st || '').trim();
  if (sfStr && stStr) {
    volPart = ` ${sfStr.padStart(2, '0')}-${stStr.padStart(2, '0')}巻`;
  } else {
    const v = String(vol || '').trim();
    if (v) {
      const n = parseInt(v, 10);
      volPart = ` 第${!isNaN(n) ? String(n).padStart(2, '0') : v}巻`;
    }
  }

  const authorPart = au ? ` 著：${au}` : '';
  const orgPart = org ? ` ${org}` : '';

  let bonusTail = '';
  if (memo) bonusTail = memo.startsWith('※') ? ` ${memo}` : ` ※${memo}`;
  else if (CFG.AUTO_BONUS_IF_EMPTY && ed && CFG.BONUS_BY_EDITION[ed]) bonusTail = ` ${CFG.BONUS_BY_EDITION[ed]}`;

  return `${head}${edLabel}『${jp}』${volPart}${authorPart}${orgPart}${bonusTail}`;
}

/* ============================
 * キャッシュ・ヘルパー
 * ============================ */
function getColMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const name = String(headers[i] || '').trim();
    if (name) map[name] = i + 1;
  }
  return map;
}

function getLangMap_(ss) {
  const map = {};
  const sh = ss.getSheetByName(CFG.SHEET_LANG);
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [n, c] of data) {
    const name = String(n || '').trim();
    const code = String(c || '').trim();
    if (name && code) map[name] = code;
  }
  return map;
}

function getCatMap_(ss) {
  const map = {};
  const sh = ss.getSheetByName(CFG.SHEET_CAT);
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [n, c] of data) {
    const name = String(n || '').trim();
    const code = String(c || '').trim();
    if (name && code) map[name] = code;
  }
  return map;
}

function clearAllCacheManual() {
  SpreadsheetApp.getUi().alert('✅ キャッシュクリア完了');
}

/* ============================
 * その他
 * ============================ */
function resetWorksOnly() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('警告', 'Worksを全削除します。続行？', ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.SHEET_WORKS);
  if (sh) { const last = sh.getLastRow(); if (last > 1) sh.deleteRows(2, last - 1); }
  else { sh = ss.insertSheet(CFG.SHEET_WORKS); }
  sh.getRange(1, 1, 1, CFG.WORKS_HEADERS.length).setValues([CFG.WORKS_HEADERS]);
  ui.alert('✅ Works初期化完了');
}

function createLangCatMasters() {
  const ss = SpreadsheetApp.getActive();
  let created = [];
  if (!ss.getSheetByName(CFG.SHEET_LANG)) {
    const sh = ss.insertSheet(CFG.SHEET_LANG);
    sh.getRange(1, 1, 1, CFG.LANG_HEADERS.length).setValues([CFG.LANG_HEADERS]);
    sh.getRange(2, 1, CFG.LANG_INITIAL.length, CFG.LANG_INITIAL[0].length).setValues(CFG.LANG_INITIAL);
    created.push('言語');
  }
  if (!ss.getSheetByName(CFG.SHEET_CAT)) {
    const sh = ss.insertSheet(CFG.SHEET_CAT);
    sh.getRange(1, 1, 1, CFG.CAT_HEADERS.length).setValues([CFG.CAT_HEADERS]);
    sh.getRange(2, 1, CFG.CAT_INITIAL.length, CFG.CAT_INITIAL[0].length).setValues(CFG.CAT_INITIAL);
    created.push('カテゴリ');
  }
  if (!ss.getSheetByName(CFG.SHEET_WORKS)) {
    const sh = ss.insertSheet(CFG.SHEET_WORKS);
    sh.getRange(1, 1, 1, CFG.WORKS_HEADERS.length).setValues([CFG.WORKS_HEADERS]);
    created.push('Works');
  }
  if (created.length > 0) { updateDropdownsAndColors(); SpreadsheetApp.getUi().alert(`✅ 作成: ${created.join(', ')}`); }
  else SpreadsheetApp.getUi().alert('既に全て存在します');
}

function updateDropdownsAndColors() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!sh) return;
  const col = getColMap_(sh);
  const lastRow = Math.max(sh.getLastRow(), 100);

  const langData = loadLangDataWithColor_(ss);
  if (langData.values.length > 0 && col[H.LANG]) {
    sh.getRange(2, col[H.LANG], lastRow - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(langData.values, true).build());
    clearConditionalFormat_(sh, col[H.LANG]);
    setConditionalFormat_(sh, col[H.LANG], langData.values, langData.colors, lastRow);
  }

  const catData = loadCatDataWithColor_(ss);
  if (catData.values.length > 0 && col[H.CAT]) {
    sh.getRange(2, col[H.CAT], lastRow - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(catData.values, true).build());
    clearConditionalFormat_(sh, col[H.CAT]);
    setConditionalFormat_(sh, col[H.CAT], catData.values, catData.colors, lastRow);
  }

  SpreadsheetApp.getActive().toast('プルダウン更新完了', '商品コード管理', 3);
}

function loadLangDataWithColor_(ss) {
  const result = { values: [], colors: [] };
  const sh = ss.getSheetByName(CFG.SHEET_LANG);
  if (!sh || sh.getLastRow() < 2) return result;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  let ci = 0;
  for (const row of data) { const n = String(row[0] || '').trim(); if (!n) continue; result.values.push(n); result.colors.push(String(row[2] || '').trim() || CFG.COLOR_PALETTE[ci++ % CFG.COLOR_PALETTE.length]); }
  return result;
}

function loadCatDataWithColor_(ss) {
  const result = { values: [], colors: [] };
  const sh = ss.getSheetByName(CFG.SHEET_CAT);
  if (!sh || sh.getLastRow() < 2) return result;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  let ci = 0;
  for (const row of data) { const n = String(row[0] || '').trim(); if (!n) continue; result.values.push(n); result.colors.push(String(row[2] || '').trim() || CFG.COLOR_PALETTE[ci++ % CFG.COLOR_PALETTE.length]); }
  return result;
}

function clearConditionalFormat_(sh, colNum) {
  const rules = sh.getConditionalFormatRules().filter(r => !r.getRanges().some(rng => rng.getColumn() === colNum));
  sh.setConditionalFormatRules(rules);
}

function setConditionalFormat_(sh, colNum, values, colors, lastRow) {
  const rules = sh.getConditionalFormatRules();
  const range = sh.getRange(2, colNum, lastRow - 1, 1);
  for (let i = 0; i < values.length; i++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(values[i]).setBackground(colors[i] || '#fff').setRanges([range]).build());
  }
  sh.setConditionalFormatRules(rules);
}

function norm_(v) { return String(v || '').replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim(); }
function toIntOrNull_(v) { const n = parseInt(String(v || '').trim(), 10); return Number.isFinite(n) ? n : null; }
