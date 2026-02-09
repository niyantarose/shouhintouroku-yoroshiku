/******************************************************
 * 商品コード管理ツール v10.5（台湾まんが）
 * 
 * 【設計方針】
 * ・入力するだけでタイトル・SKUが段階的に自動生成される
 * ・onEdit時にWorksへ重複登録しない（常に再計算キーで照合）
 * ・一括更新 / 重複削除はセーフティネットとして用意
 * 
 * ✅ v10.5の修正点
 * - onEdit: waitLock(5000)で確実にロック取得
 * - 自己更新フラグにタイムスタンプ（10秒で自動解除）
 * - Works照合: 保存済みキーを信用せず常にタイトル+作者で再計算
 * - 一括更新時に自動で再正規化+ID振り直し
 * - コア処理はUI無し（スクリプト内から呼べる）
 ******************************************************/

/* ============================ 
 * 設定
 * ============================ */
const 設定 = {
  マスターシート名: '台湾まんがマスターシート',
  作品シート名: 'Works',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  
  作品ヘッダー: ['WorksKey', '作品ID', '日本語タイトル', '作者', '原題タイトル', '登録済み巻', '最新巻', '更新日時'],
  言語ヘッダー: ['言語', 'コード', '色'],
  カテゴリヘッダー: ['カテゴリ', 'コード', '色'],
  
  言語初期値: [
    ['台湾', 'TW', '#b7e1cd'],
    ['韓国', 'KR', '#fce8b2'],
    ['日本', 'JP', '#f4c7c3'],
    ['中国', 'CN', '#c9daf8'],
    ['タイ', 'TH', '#d9d2e9'],
    ['英語', 'US', '#fce5cd']
  ],
  
  カテゴリ初期値: [
    ['コミック', 'CM', '#fff2cc'],
    ['まんが', 'CM', '#fff2cc'],
    ['小説', 'NV', '#d9ead3'],
    ['グッズ', 'GD', '#cfe2f3'],
    ['設定集', 'ART', '#f4cccc'],
    ['アートブック', 'ART', '#f4cccc'],
    ['雑誌', 'MZ', '#d9d2e9']
  ],
  
  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd', '#fff2cc', '#d9ead3', '#cfe2f3', '#f4cccc'],
  
  形態プレフィックス: {
    '特装版': 'S',
    '初版限定版': 'F',
    '初回限定版': 'F'
  },
  
  特典自動付与: true,
  形態別特典: {
    '初版限定版': '※初版限定版特典付き',
    '初回限定版': '※初回限定版特典付き',
    '特装版': '※特装版限定特典付き'
  }
};

const 列名 = {
  発行チェック: '発番発行',
  商品コード: '商品コード(SKU)',
  タイトル: 'タイトル',
  作者: '作者',
  日本語タイトル: '日本語タイトル',
  原題: '原題タイトル',
  形態: '形態（通常/初回限定/特装）',
  言語: '言語',
  カテゴリ: 'カテゴリ',
  単巻数: '単巻数',
  セット開始: 'セット巻数開始番号',
  セット終了: 'セット巻数終了番号',
  特典メモ: '特典メモ',
  作品ID: '作品ID(W)（自動）',
  SKU: 'SKU（自動）',
  コードステータス: '商品コードステータス',
  登録状況: '登録状況'
};

const 監視列 = [
  列名.作者, 列名.日本語タイトル, 列名.原題,
  列名.言語, 列名.カテゴリ, 列名.形態,
  列名.単巻数, 列名.セット開始, 列名.セット終了, 列名.特典メモ
];

/* ============================ 
 * メニュー
 * ============================ */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('商品コード管理')
      .addItem('① 商品コード確定発行', 'メニュー_確定発行')
      .addItem('② 商品コード削除', 'メニュー_削除')
      .addSeparator()
      .addItem('③ 既存データ一括更新（Works整理含む）', 'メニュー_一括更新')
      .addSeparator()
      .addItem('④ Works初期化', 'メニュー_Works初期化')
      .addItem('⑤ マスター作成', 'メニュー_マスター作成')
      .addItem('⑥ プルダウン更新', 'メニュー_プルダウン更新')
      .addSeparator()
      .addItem('🔍 Works重複チェック・統合', 'メニュー_重複統合')
      .addItem('🔄 WorksKey再正規化', 'メニュー_WorksKey再正規化')
      .addItem('🔢 Works ID振り直し', 'メニュー_ID振り直し')
      .addToUi();
  } catch (e) {
    // スクリプトエディタから直接実行した場合は無視
    Logger.log('onOpenはスプレッドシートから実行してください: ' + e.message);
  }
}

/* ============================================================
 *  onEdit ― 段階的にタイトル・SKUを自動生成
 *  ✅ waitLock(5000)で確実にロック取得
 *  ✅ 自己更新フラグに10秒タイムアウト
 *  ✅ Works照合は常に再計算キー
 * ============================================================ */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const sh = e.range.getSheet();
  if (sh.getName() !== 設定.マスターシート名) return;
  
  const 行 = e.range.getRow();
  if (行 < 2) return;
  
  // ★ ロックで重複実行を防止（最大5秒待つ）
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    return; // 5秒待ってもダメなら諦める
  }
  
  try {
    // ★ 自己更新チェック（10秒タイムアウト付き）
    if (自己更新中か_()) return;
    
    const 列マップ = 列番号を取得(sh);
    const 編集列 = e.range.getColumn();
    const 監視列番号 = 監視列.map(h => 列マップ[h]).filter(Boolean);
    
    if (!監視列番号.includes(編集列)) return;
    
    自己更新を開始_();
    try {
      const 最終列 = sh.getLastColumn();
      const 行データ = sh.getRange(行, 1, 1, 最終列).getValues()[0];
      
      const 日本語タイトル = 正規化(行データ[(列マップ[列名.日本語タイトル] || 1) - 1]);
      let 作者 = 正規化(行データ[(列マップ[列名.作者] || 1) - 1]);
      let 原題 = 正規化(行データ[(列マップ[列名.原題] || 1) - 1]);
      const 言語 = 正規化(行データ[(列マップ[列名.言語] || 1) - 1]);
      const カテゴリ = 正規化(行データ[(列マップ[列名.カテゴリ] || 1) - 1]);
      const 形態 = 正規化(行データ[(列マップ[列名.形態] || 1) - 1]);
      const 単巻数 = 行データ[(列マップ[列名.単巻数] || 1) - 1];
      const セット開始 = 行データ[(列マップ[列名.セット開始] || 1) - 1];
      const セット終了 = 行データ[(列マップ[列名.セット終了] || 1) - 1];
      const 特典メモ = 正規化(行データ[(列マップ[列名.特典メモ] || 1) - 1]);
      
      const ss = SpreadsheetApp.getActive();
      const 言語マップ = 言語マップを取得(ss);
      const カテゴリマップ = カテゴリマップを取得(ss);
      const 言語コード = 言語 ? (言語マップ[言語] || 'XX') : '';
      const カテゴリコード = カテゴリ ? (カテゴリマップ[カテゴリ] || 'XX') : '';
      
      // ★ 作品ID取得（タイトル+作者が揃ったらWorks登録）
      let 作品ID = '';
      let 巻数警告 = '';
      let 使用作者 = 作者;
      let 使用原題 = 原題;
      
      if (日本語タイトル && 作者) {
        const 作品シート = 作品シートを確保(ss);
        const 作品結果 = 作品IDを取得または作成(作品シート, 日本語タイトル, 作者, 原題);
        作品ID = 作品結果.作品ID;
        
        // 既存作品なら作者・原題を補完
        if (!作品結果.新規 && 作品結果.既存データ) {
          if (!使用作者 && 作品結果.既存データ.作者) {
            使用作者 = 作品結果.既存データ.作者;
            if (列マップ[列名.作者]) sh.getRange(行, 列マップ[列名.作者]).setValue(使用作者);
          }
          if (!使用原題 && 作品結果.既存データ.原題) {
            使用原題 = 作品結果.既存データ.原題;
            if (列マップ[列名.原題]) sh.getRange(行, 列マップ[列名.原題]).setValue(使用原題);
          }
        }
        
        // 巻数重複チェック
        const 巻数 = 数値変換(単巻数);
        if (巻数 != null && 作品結果.登録済み巻) {
          const 既存巻 = String(作品結果.登録済み巻)
            .split(',')
            .map(v => parseInt(String(v).trim(), 10))
            .filter(n => !isNaN(n));
          if (既存巻.includes(巻数)) 巻数警告 = '【重複注意】';
        }
      }
      
      // ★ 段階的SKU生成
      const SKU = SKUを段階生成(
        言語コード, 形態, 作品ID, カテゴリコード,
        単巻数, セット開始, セット終了
      );
      
      // ★ 段階的タイトル生成
      const タイトル = 巻数警告 + タイトルを段階生成(
        言語, カテゴリ, 形態, 日本語タイトル,
        単巻数, セット開始, セット終了,
        使用作者, 使用原題, 特典メモ
      );
      
      // 書き込み（登録状況は触らない）
      if (列マップ[列名.商品コード]) sh.getRange(行, 列マップ[列名.商品コード]).setValue(SKU);
      if (列マップ[列名.タイトル]) sh.getRange(行, 列マップ[列名.タイトル]).setValue(タイトル);
      if (列マップ[列名.作品ID]) sh.getRange(行, 列マップ[列名.作品ID]).setValue(作品ID);
      if (列マップ[列名.SKU]) sh.getRange(行, 列マップ[列名.SKU]).setValue(SKU);
      
      // ステータス
      if (列マップ[列名.コードステータス]) {
        const st = (日本語タイトル && 作者 && 言語 && カテゴリ) ? '商品コード（予約）' : '入力中...';
        sh.getRange(行, 列マップ[列名.コードステータス]).setValue(st);
      }
      
    } finally {
      自己更新を終了_();
    }
  } finally {
    lock.releaseLock();
  }
}

/* ============================ 
 * 作品IDを取得または作成
 * ✅ 保存済みWorksKeyを信用せず常に再計算して照合
 * ============================ */
function 作品IDを取得または作成(作品シート, 日本語タイトル, 作者, 原題) {
  const worksKey = WorksKeyを作る(日本語タイトル, 作者);
  const 最終行 = 作品シート.getLastRow();
  
  let 最大ID = 0;
  let 見つかった行 = null;
  let 見つかったID = null;
  let 登録済み巻 = '';
  let 既存データ = null;
  
  if (最終行 >= 2) {
    const データ = 作品シート.getRange(2, 1, 最終行 - 1, 8).getValues();
    
    for (let i = 0; i < データ.length; i++) {
      const r = データ[i];
      const idRaw = r[1];
      const idStr = String(idRaw == null ? '' : idRaw).trim();
      
      // 最大ID抽出
      const num = parseInt(idStr, 10);
      if (!isNaN(num) && num > 最大ID) 最大ID = num;
      
      // ✅ 常にタイトル・作者から再計算して比較
      const タイトル = 正規化(r[2] || '');
      const 著者 = 正規化(r[3] || '');
      if (!タイトル || !著者) continue;
      
      const 再計算Key = WorksKeyを作る(タイトル, 著者);
      
      if (再計算Key === worksKey) {
        見つかった行 = i + 2;
        見つかったID = idStr ? idStr.padStart(4, '0') : null;
        登録済み巻 = String(r[5] || '');
        既存データ = {
          作者: 正規化(r[3] || ''),
          原題: 正規化(r[4] || '')
        };
        
        // 保存済みWorksKeyが古ければ更新
        const 保存済みKey = String(r[0] || '').trim();
        if (保存済みKey !== worksKey) {
          作品シート.getRange(見つかった行, 1).setValue(worksKey);
        }
        
        break;
      }
    }
  }
  
  if (見つかったID) {
    if (見つかった行 && 見つかったID.length !== 4) {
      作品シート.getRange(見つかった行, 2).setValue(見つかったID);
    }
    return {
      作品ID: 見つかったID,
      新規: false,
      登録済み巻,
      行番号: 見つかった行,
      既存データ
    };
  }
  
  // 新規作品
  const 新ID番号 = 最大ID + 1;
  const 新ID = String(新ID番号).padStart(4, '0');
  const 新規行 = Math.max(2, 最終行 + 1);
  
  作品シート.getRange(新規行, 1, 1, 8).setValues([[
    worksKey,
    新ID,
    正規化(日本語タイトル),
    正規化(作者),
    正規化(原題),
    '',
    '',
    new Date()
  ]]);
  
  return {
    作品ID: 新ID,
    新規: true,
    登録済み巻: '',
    行番号: 新規行,
    既存データ: null
  };
}

/* ============================================================
 *  コア処理（UI無し・スクリプト内部から直接呼べる）
 *  ※ロックは呼び出し側で取得済みの前提
 * ============================================================ */

/**
 * WorksKey再正規化 + 重複統合（コア）
 * @param {Sheet} 作品シート
 * @return {Object} { キー更新数, 統合数, 削除行数 }
 */
function WorksKey再正規化を実行(作品シート) {
  if (!作品シート || 作品シート.getLastRow() < 2) {
    return { キー更新数: 0, 統合数: 0, 削除行数: 0 };
  }
  
  const 最終行 = 作品シート.getLastRow();
  const データ = 作品シート.getRange(2, 1, 最終行 - 1, 8).getValues();
  
  const keyMap = new Map();
  
  for (let i = 0; i < データ.length; i++) {
    const r = データ[i];
    const タイトル = 正規化(r[2] || '');
    const 作者 = 正規化(r[3] || '');
    if (!タイトル && !作者) continue;
    
    const 再計算key = WorksKeyを作る(タイトル, 作者);
    const 元key = String(r[0] || '').trim();
    
    if (!keyMap.has(再計算key)) keyMap.set(再計算key, []);
    keyMap.get(再計算key).push({
      行: i + 2, データ: r, 元key, キー変更: 元key !== 再計算key
    });
  }
  
  let キー更新数 = 0, 統合数 = 0;
  const 削除行 = [];
  
  for (const [key, 行配列] of keyMap.entries()) {
    if (行配列.length === 1) {
      if (行配列[0].キー変更) {
        作品シート.getRange(行配列[0].行, 1).setValue(key);
        キー更新数++;
      }
    } else {
      // 重複あり：若いIDに統合
      行配列.sort((a, b) => {
        const idA = parseInt(String(a.データ[1] || '9999'), 10);
        const idB = parseInt(String(b.データ[1] || '9999'), 10);
        return idA - idB;
      });
      
      const 残す = 行配列[0];
      作品シート.getRange(残す.行, 1).setValue(key);
      
      // 巻数統合
      const 全巻 = new Set();
      for (const item of 行配列) {
        String(item.データ[5] || '').split(',').forEach(v => {
          const n = parseInt(String(v).trim(), 10);
          if (!isNaN(n)) 全巻.add(n);
        });
      }
      if (全巻.size > 0) {
        const 統合巻 = Array.from(全巻).sort((a, b) => a - b);
        作品シート.getRange(残す.行, 6).setValue(統合巻.join(','));
        作品シート.getRange(残す.行, 7).setValue(Math.max(...統合巻));
        作品シート.getRange(残す.行, 8).setValue(new Date());
      }
      
      // 原題補完
      if (!正規化(残す.データ[4] || '')) {
        for (let j = 1; j < 行配列.length; j++) {
          const t = 正規化(行配列[j].データ[4] || '');
          if (t) { 作品シート.getRange(残す.行, 5).setValue(t); break; }
        }
      }
      
      for (let j = 1; j < 行配列.length; j++) 削除行.push(行配列[j].行);
      統合数++;
    }
  }
  
  削除行.sort((a, b) => b - a);
  for (const 行 of 削除行) 作品シート.deleteRow(行);
  
  return { キー更新数, 統合数, 削除行数: 削除行.length };
}

/**
 * Works ID振り直し（コア）
 * 欠番を詰めて1から連番にする
 * @param {Sheet} 作品シート
 * @return {Object} { 変更数, 旧新マップ: { 旧ID: 新ID } }
 */
function WorksID振り直しを実行(作品シート) {
  const result = { 変更数: 0, 旧新マップ: {} };
  if (!作品シート || 作品シート.getLastRow() < 2) return result;
  
  const 最終行 = 作品シート.getLastRow();
  const データ = 作品シート.getRange(2, 1, 最終行 - 1, 8).getValues();
  
  // 現在のID順にソートして連番を振る
  const 行リスト = [];
  for (let i = 0; i < データ.length; i++) {
    const id = parseInt(String(データ[i][1] || '0'), 10);
    行リスト.push({ 行: i + 2, 旧ID: id, データ: データ[i] });
  }
  行リスト.sort((a, b) => a.旧ID - b.旧ID);
  
  const 更新値 = [];
  for (let i = 0; i < 行リスト.length; i++) {
    const 新ID = i + 1;
    const 旧ID = 行リスト[i].旧ID;
    const 新IDstr = String(新ID).padStart(4, '0');
    
    if (旧ID !== 新ID) {
      result.旧新マップ[String(旧ID).padStart(4, '0')] = 新IDstr;
      result.変更数++;
    }
    更新値.push([新IDstr]);
  }
  
  // ID列だけ一括更新（行の並び順が変わるので全行上書き）
  // まずID順にデータを並べ替え
  const 新データ = 行リスト.map((item, i) => {
    const r = item.データ.slice();
    r[1] = String(i + 1).padStart(4, '0');
    return r;
  });
  
  作品シート.getRange(2, 1, 新データ.length, 8).setValues(新データ);
  
  return result;
}

/**
 * Works重複検出のみ（統合しない・レポート用）
 */
function Works重複を検出(作品シート) {
  if (!作品シート || 作品シート.getLastRow() < 2) return [];
  
  const データ = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, 8).getValues();
  const keyMap = new Map();
  
  for (let i = 0; i < データ.length; i++) {
    const r = データ[i];
    const タイトル = 正規化(r[2]);
    const 作者 = 正規化(r[3]);
    if (!タイトル || !作者) continue;
    
    const key = WorksKeyを作る(タイトル, 作者);
    if (!keyMap.has(key)) keyMap.set(key, []);
    keyMap.get(key).push({ 行: i + 2, データ: r });
  }
  
  const 重複リスト = [];
  for (const [key, 行配列] of keyMap.entries()) {
    if (行配列.length > 1) 重複リスト.push({ key, 行配列 });
  }
  return 重複リスト;
}

/* ============================================================
 *  メニューラッパー（UIあり）
 * ============================================================ */

function メニュー_WorksKey再正規化() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(設定.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  
  const res = ui.alert('確認', '全WorksKeyを再計算し重複を統合します。続行？', ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = WorksKey再正規化を実行(作品シート);
    ui.alert(`✅ 完了\nキー更新: ${r.キー更新数}\n重複統合: ${r.統合数}\n削除: ${r.削除行数}` +
      (r.統合数 > 0 ? '\n\n③一括更新を実行してください。' : ''));
  } finally { lock.releaseLock(); }
}

function メニュー_重複統合() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(設定.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 重複リスト = Works重複を検出(作品シート);
    if (重複リスト.length === 0) { ui.alert('✅ 重複はありません！'); return; }
    
    let レポート = `🔍 重複検出: ${重複リスト.length}件\n\n`;
    for (const dup of 重複リスト) {
      レポート += `【${dup.行配列[0].データ[2]}】 著：${dup.行配列[0].データ[3]}\n`;
      for (const item of dup.行配列) {
        レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
      }
      レポート += '\n';
    }
    
    const res = ui.alert('重複チェック結果', レポート + '\n自動統合しますか？', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
    
    const r = WorksKey再正規化を実行(作品シート);
    ui.alert(`✅ 統合完了\n統合: ${r.統合数}件, 削除: ${r.削除行数}行\n\n③一括更新を実行してください。`);
  } finally { lock.releaseLock(); }
}

function メニュー_ID振り直し() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(設定.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  
  const res = ui.alert('確認', 'Works IDを1から連番に振り直します。\n※マスターシートの商品コードも変わります。\n③一括更新とセットで実行してください。\n\n続行？', ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = WorksID振り直しを実行(作品シート);
    ui.alert(`✅ ID振り直し完了\n変更: ${r.変更数}件\n\n③一括更新を実行してください。`);
  } finally { lock.releaseLock(); }
}

/* ============================ 
 * ① 商品コード確定発行
 * ============================ */
function メニュー_確定発行() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.マスターシート名);
  if (!sh) return;
  
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ);
  if (最終行 < 2) return;
  
  const 言語マップ = 言語マップを取得(ss);
  const カテゴリマップ = カテゴリマップを取得(ss);
  if (Object.keys(言語マップ).length === 0 || Object.keys(カテゴリマップ).length === 0) {
    ui.alert('言語/カテゴリマスターがありません。⑤を実行してください。'); return;
  }
  
  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 作品シート = 作品シートを確保(ss);
  
  const lock = LockService.getDocumentLock(); lock.waitLock(60000);
  try {
    // 確定発行前にWorks整理
    WorksKey再正規化を実行(作品シート);
    
    const 作品データ = 全作品データを読み込み(作品シート);
    const out作品ID = [], outSKU = [], out商品コード = [], outタイトル = [], outステータス = [];
    let 発行数 = 0;
    
    for (let i = 0; i < 全データ.length; i++) {
      const r = 全データ[i];
      
      if (r[(列マップ[列名.発行チェック] || 1) - 1] !== true) {
        out作品ID.push([r[(列マップ[列名.作品ID] || 1) - 1] || '']);
        outSKU.push([r[(列マップ[列名.SKU] || 1) - 1] || '']);
        out商品コード.push([r[(列マップ[列名.商品コード] || 1) - 1] || '']);
        outタイトル.push([r[(列マップ[列名.タイトル] || 1) - 1] || '']);
        outステータス.push([r[(列マップ[列名.コードステータス] || 1) - 1] || '']);
        continue;
      }
      
      const 日本語タイトル = 正規化(r[(列マップ[列名.日本語タイトル] || 1) - 1]);
      const 作者 = 正規化(r[(列マップ[列名.作者] || 1) - 1]);
      let 原題 = 正規化(r[(列マップ[列名.原題] || 1) - 1]);
      const 言語 = 正規化(r[(列マップ[列名.言語] || 1) - 1]);
      const カテゴリ = 正規化(r[(列マップ[列名.カテゴリ] || 1) - 1]);
      const 形態 = 正規化(r[(列マップ[列名.形態] || 1) - 1]);
      
      if (!日本語タイトル || !作者 || !言語 || !カテゴリ) {
        out作品ID.push([r[(列マップ[列名.作品ID] || 1) - 1] || '']);
        outSKU.push([r[(列マップ[列名.SKU] || 1) - 1] || '']);
        out商品コード.push([r[(列マップ[列名.商品コード] || 1) - 1] || '']);
        outタイトル.push([r[(列マップ[列名.タイトル] || 1) - 1] || '']);
        outステータス.push(['入力中...']);
        continue;
      }
      
      const 言語コード = 言語マップ[言語] || 'XX';
      const カテゴリコード = カテゴリマップ[カテゴリ] || 'XX';
      const worksKey = WorksKeyを作る(日本語タイトル, 作者);
      
      let 作品ID = 作品データ.keyToId[worksKey];
      const 既存 = 作品データ.keyToData[worksKey];
      if (既存 && !原題 && 既存.原題) 原題 = 正規化(既存.原題);
      
      if (!作品ID) {
        作品データ.maxId++;
        作品ID = String(作品データ.maxId).padStart(4, '0');
        作品データ.keyToId[worksKey] = 作品ID;
        作品データ.newRows.push([worksKey, 作品ID, 日本語タイトル, 作者, 原題, '', '', new Date()]);
      } else {
        作品ID = String(作品ID).padStart(4, '0');
      }
      
      const 巻数 = 数値変換(r[(列マップ[列名.単巻数] || 1) - 1]);
      if (巻数 != null) {
        if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set();
        作品データ.keyToVols[worksKey].add(巻数);
      }
      
      const SKU = SKUを生成(言語コード, 形態, 作品ID, カテゴリコード,
        r[(列マップ[列名.単巻数] || 1) - 1], r[(列マップ[列名.セット開始] || 1) - 1], r[(列マップ[列名.セット終了] || 1) - 1]);
      const タイトル = タイトルを生成(言語, カテゴリ, 形態, 日本語タイトル,
        r[(列マップ[列名.単巻数] || 1) - 1], r[(列マップ[列名.セット開始] || 1) - 1], r[(列マップ[列名.セット終了] || 1) - 1],
        作者, 原題, 正規化(r[(列マップ[列名.特典メモ] || 1) - 1]));
      
      out作品ID.push([作品ID]); outSKU.push([SKU]); out商品コード.push([SKU]);
      outタイトル.push([タイトル]); outステータス.push(['商品コード（発行済み確定）']);
      発行数++;
    }
    
    if (列マップ[列名.作品ID]) sh.getRange(2, 列マップ[列名.作品ID], out作品ID.length, 1).setValues(out作品ID);
    if (列マップ[列名.SKU]) sh.getRange(2, 列マップ[列名.SKU], outSKU.length, 1).setValues(outSKU);
    if (列マップ[列名.商品コード]) sh.getRange(2, 列マップ[列名.商品コード], out商品コード.length, 1).setValues(out商品コード);
    if (列マップ[列名.タイトル]) sh.getRange(2, 列マップ[列名.タイトル], outタイトル.length, 1).setValues(outタイトル);
    if (列マップ[列名.コードステータス]) sh.getRange(2, 列マップ[列名.コードステータス], outステータス.length, 1).setValues(outステータス);
    
    作品データを更新(作品シート, 作品データ);
    ui.alert(`✅ 確定発行完了: ${発行数}件`);
  } finally { lock.releaseLock(); }
}

/* ============================ 
 * ② 商品コード削除
 * ============================ */
function メニュー_削除() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(設定.マスターシート名);
  if (!sh) return;
  
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ);
  if (最終行 < 2) return;
  
  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 削除行 = [];
  for (let i = 0; i < 全データ.length; i++) {
    if (全データ[i][(列マップ[列名.発行チェック] || 1) - 1] === true) 削除行.push(i + 2);
  }
  if (削除行.length === 0) { ui.alert('チェックが入った行がありません'); return; }
  
  const res = ui.alert('確認', `${削除行.length}件を削除します。続行？`, ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  
  削除行.sort((a, b) => b - a);
  for (const 行 of 削除行) sh.deleteRow(行);
  ui.alert(`✅ 削除完了: ${削除行.length}件`);
}

/* ============================ 
 * ③ 既存データ一括更新（Works整理含む）
 * ✅ 再正規化 → ID振り直し → 全行再生成
 * ============================ */
function メニュー_一括更新() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('確認', '全行を再生成します（Works重複整理+ID振り直し含む）。続行？', ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.マスターシート名);
  if (!sh) return;
  
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ);
  if (最終行 < 2) { ui.alert('データがありません'); return; }
  
  const 言語マップ = 言語マップを取得(ss);
  const カテゴリマップ = カテゴリマップを取得(ss);
  if (Object.keys(言語マップ).length === 0 || Object.keys(カテゴリマップ).length === 0) {
    ui.alert('言語/カテゴリマスターがありません。⑤を実行してください。'); return;
  }
  
  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 作品シート = 作品シートを確保(ss);
  
  const lock = LockService.getDocumentLock(); lock.waitLock(60000);
  try {
    // ★ Step1: Works整理（再正規化+重複統合）
    const 正規化結果 = WorksKey再正規化を実行(作品シート);
    
    // ★ Step2: ID振り直し（欠番を詰める）
    const ID結果 = WorksID振り直しを実行(作品シート);
    
    // ★ Step3: 整理後のWorksデータで全行再生成
    const 作品データ = 全作品データを読み込み(作品シート);
    
    const out作品ID = [], outSKU = [], out商品コード = [], outタイトル = [];
    const outステータス = [], out作者 = [], out原題 = [];
    
    for (let i = 0; i < 全データ.length; i++) {
      const r = 全データ[i];
      
      const 日本語タイトル = 正規化(r[(列マップ[列名.日本語タイトル] || 1) - 1]);
      const 言語 = 正規化(r[(列マップ[列名.言語] || 1) - 1]);
      const カテゴリ = 正規化(r[(列マップ[列名.カテゴリ] || 1) - 1]);
      let 作者 = 正規化(r[(列マップ[列名.作者] || 1) - 1]);
      let 原題 = 正規化(r[(列マップ[列名.原題] || 1) - 1]);
      const 形態 = 正規化(r[(列マップ[列名.形態] || 1) - 1]);
      
      if (!日本語タイトル || !言語 || !カテゴリ) {
        out作品ID.push([r[(列マップ[列名.作品ID] || 1) - 1] || '']);
        outSKU.push([r[(列マップ[列名.SKU] || 1) - 1] || '']);
        out商品コード.push([r[(列マップ[列名.商品コード] || 1) - 1] || '']);
        outタイトル.push([r[(列マップ[列名.タイトル] || 1) - 1] || '']);
        outステータス.push([r[(列マップ[列名.コードステータス] || 1) - 1] || '']);
        out作者.push([作者 || '']); out原題.push([原題 || '']);
        continue;
      }
      
      const 言語コード = 言語マップ[言語] || 'XX';
      const カテゴリコード = カテゴリマップ[カテゴリ] || 'XX';
      let 作品ID = '';
      
      if (日本語タイトル && 作者) {
        const worksKey = WorksKeyを作る(日本語タイトル, 作者);
        const 既存 = 作品データ.keyToData[worksKey];
        if (既存) {
          if (!作者 && 既存.作者) 作者 = 正規化(既存.作者);
          if (!原題 && 既存.原題) 原題 = 正規化(既存.原題);
        }
        
        作品ID = 作品データ.keyToId[worksKey];
        if (!作品ID) {
          作品データ.maxId++;
          作品ID = String(作品データ.maxId).padStart(4, '0');
          作品データ.keyToId[worksKey] = 作品ID;
          作品データ.newRows.push([worksKey, 作品ID, 日本語タイトル, 作者, 原題, '', '', new Date()]);
        } else {
          作品ID = String(作品ID).padStart(4, '0');
        }
      }
      
      const SKU = SKUを段階生成(言語コード, 形態, 作品ID, カテゴリコード,
        r[(列マップ[列名.単巻数] || 1) - 1], r[(列マップ[列名.セット開始] || 1) - 1], r[(列マップ[列名.セット終了] || 1) - 1]);
      const タイトル = タイトルを生成(言語, カテゴリ, 形態, 日本語タイトル,
        r[(列マップ[列名.単巻数] || 1) - 1], r[(列マップ[列名.セット開始] || 1) - 1], r[(列マップ[列名.セット終了] || 1) - 1],
        作者, 原題, 正規化(r[(列マップ[列名.特典メモ] || 1) - 1]));
      
      out作品ID.push([作品ID]); outSKU.push([SKU]); out商品コード.push([SKU]);
      outタイトル.push([タイトル]);
      outステータス.push([(日本語タイトル && 作者 && 言語 && カテゴリ) ? '商品コード（予約）' : '入力中...']);
      out作者.push([作者 || '']); out原題.push([原題 || '']);
    }
    
    if (列マップ[列名.作品ID]) sh.getRange(2, 列マップ[列名.作品ID], out作品ID.length, 1).setValues(out作品ID);
    if (列マップ[列名.SKU]) sh.getRange(2, 列マップ[列名.SKU], outSKU.length, 1).setValues(outSKU);
    if (列マップ[列名.商品コード]) sh.getRange(2, 列マップ[列名.商品コード], out商品コード.length, 1).setValues(out商品コード);
    if (列マップ[列名.タイトル]) sh.getRange(2, 列マップ[列名.タイトル], outタイトル.length, 1).setValues(outタイトル);
    if (列マップ[列名.コードステータス]) sh.getRange(2, 列マップ[列名.コードステータス], outステータス.length, 1).setValues(outステータス);
    if (列マップ[列名.作者]) sh.getRange(2, 列マップ[列名.作者], out作者.length, 1).setValues(out作者);
    if (列マップ[列名.原題]) sh.getRange(2, 列マップ[列名.原題], out原題.length, 1).setValues(out原題);
    
    作品データを更新(作品シート, 作品データ);
    
    let msg = `✅ 一括更新完了: ${out作品ID.length}件`;
    if (正規化結果.統合数 > 0) msg += `\nWorks重複統合: ${正規化結果.統合数}件`;
    if (ID結果.変更数 > 0) msg += `\nID振り直し: ${ID結果.変更数}件`;
    ui.alert(msg);
  } finally { lock.releaseLock(); }
}

/* ============================ 
 * 作品シート関連
 * ============================ */
function 作品シートを確保(ss) {
  let sh = ss.getSheetByName(設定.作品シート名);
  if (!sh) {
    sh = ss.insertSheet(設定.作品シート名);
    sh.getRange(1, 1, 1, 設定.作品ヘッダー.length).setValues([設定.作品ヘッダー]);
  }
  return sh;
}

function 全作品データを読み込み(作品シート) {
  const result = {
    keyToId: {}, keyToData: {}, keyToRow: {}, keyToVols: {},
    maxId: 0, newRows: [], keyUpdates: []
  };
  
  const 最終行 = 作品シート.getLastRow();
  if (最終行 < 2) return result;
  
  const データ = 作品シート.getRange(2, 1, 最終行 - 1, 8).getValues();
  
  for (let i = 0; i < データ.length; i++) {
    const r = データ[i];
    const idStr = String(r[1] == null ? '' : r[1]).trim();
    if (!idStr) continue;
    
    const タイトル = 正規化(r[2] || '');
    const 作者 = 正規化(r[3] || '');
    
    let key;
    if (タイトル && 作者) {
      key = WorksKeyを作る(タイトル, 作者);
    } else {
      key = String(r[0] || '').trim();
    }
    if (!key) continue;
    
    const 保存済みKey = String(r[0] || '').trim();
    if (保存済みKey !== key) {
      result.keyUpdates.push({ 行: i + 2, key });
    }
    
    // 重複キーなら若いIDを優先
    if (result.keyToId[key]) {
      const 既存ID = parseInt(result.keyToId[key], 10);
      const 新ID = parseInt(idStr, 10);
      if (!isNaN(新ID) && !isNaN(既存ID) && 新ID >= 既存ID) {
        const vols = String(r[5] || '').split(',').map(v => parseInt(String(v).trim(), 10)).filter(n => !isNaN(n));
        if (!result.keyToVols[key]) result.keyToVols[key] = new Set();
        for (const v of vols) result.keyToVols[key].add(v);
        const num = parseInt(idStr, 10);
        if (!isNaN(num) && num > result.maxId) result.maxId = num;
        continue;
      }
    }
    
    const id = idStr.padStart(4, '0');
    result.keyToId[key] = id;
    result.keyToRow[key] = i + 2;
    result.keyToData[key] = { 作者: 正規化(r[3] || ''), 原題: 正規化(r[4] || '') };
    
    const vols = String(r[5] || '').split(',').map(v => parseInt(String(v).trim(), 10)).filter(n => !isNaN(n));
    if (!result.keyToVols[key]) result.keyToVols[key] = new Set();
    for (const v of vols) result.keyToVols[key].add(v);
    
    const num = parseInt(idStr, 10);
    if (!isNaN(num) && num > result.maxId) result.maxId = num;
  }
  
  return result;
}

function 作品データを更新(作品シート, 作品データ) {
  if (作品データ.keyUpdates && 作品データ.keyUpdates.length > 0) {
    for (const upd of 作品データ.keyUpdates) {
      作品シート.getRange(upd.行, 1).setValue(upd.key);
    }
    作品データ.keyUpdates = [];
  }
  
  if (作品データ.newRows.length > 0) {
    const 開始行 = Math.max(2, 作品シート.getLastRow() + 1);
    作品シート.getRange(開始行, 1, 作品データ.newRows.length, 8).setValues(作品データ.newRows);
    for (let i = 0; i < 作品データ.newRows.length; i++) {
      const key = String(作品データ.newRows[i][0] || '').trim();
      if (key) 作品データ.keyToRow[key] = 開始行 + i;
    }
  }
  
  for (const [key, vols] of Object.entries(作品データ.keyToVols)) {
    const 行番号 = 作品データ.keyToRow[key];
    if (!行番号 || !vols || vols.size === 0) continue;
    const 巻配列 = Array.from(vols).sort((a, b) => a - b);
    作品シート.getRange(行番号, 6).setValue(巻配列.join(','));
    作品シート.getRange(行番号, 7).setValue(Math.max(...巻配列));
    作品シート.getRange(行番号, 8).setValue(new Date());
  }
  
  作品データ.newRows = [];
}

/* ============================ 
 * SKU / タイトル生成
 * ============================ */
function SKUを段階生成(言語コード, 形態, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了) {
  const parts = [];
  if (言語コード) parts.push(言語コード);
  
  const 形態プレ = 設定.形態プレフィックス[String(形態 || '').trim()] || '';
  if (形態プレ) parts.push(形態プレ);
  
  if (作品ID) {
    parts.push(String(作品ID).padStart(4, '0'));
  } else if (言語コード) {
    parts.push('????');
  }
  
  if (カテゴリコード) parts.push('-' + カテゴリコード);
  
  const sf = String(セット開始 || '').trim();
  const st = String(セット終了 || '').trim();
  if (sf && st) {
    parts.push('-' + sf.padStart(2, '0') + st.padStart(2, '0'));
  } else {
    const v = String(単巻数 || '').trim();
    if (v) {
      const n = parseInt(v, 10);
      parts.push('-' + (!isNaN(n) ? String(n).padStart(2, '0') : v));
    }
  }
  return parts.join('');
}

function SKUを生成(言語コード, 形態, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了) {
  const 形態プレ = 設定.形態プレフィックス[String(形態 || '').trim()] || '';
  const base = String(言語コード || '') + String(形態プレ || '') + String(作品ID || '').padStart(4, '0') + '-' + String(カテゴリコード || '');
  const sf = String(セット開始 || '').trim();
  const st = String(セット終了 || '').trim();
  if (sf && st) return base + '-' + sf.padStart(2, '0') + st.padStart(2, '0');
  const v = String(単巻数 || '').trim();
  if (v) {
    const n = parseInt(v, 10);
    return base + '-' + (!isNaN(n) ? String(n).padStart(2, '0') : v);
  }
  return base;
}

function タイトルを段階生成(言語, カテゴリ, 形態, 日本語タイトル, 単巻数, セット開始, セット終了, 作者, 原題, 特典メモ) {
  const parts = [];
  const 言語表示 = 言語 ? `${言語}版` : '';
  const head = [言語表示, カテゴリ].filter(Boolean).join(' ');
  if (head) parts.push(head);
  if (形態) parts.push(`（${形態}）`);
  if (日本語タイトル) parts.push(`『${日本語タイトル}』`);
  
  const sf = String(セット開始 || '').trim();
  const st = String(セット終了 || '').trim();
  if (sf && st) {
    parts.push(`${sf.padStart(2, '0')}-${st.padStart(2, '0')}巻`);
  } else {
    const v = String(単巻数 || '').trim();
    if (v) {
      const n = parseInt(v, 10);
      parts.push(`第${!isNaN(n) ? String(n).padStart(2, '0') : v}巻`);
    }
  }
  
  if (作者) parts.push(`著：${作者}`);
  if (原題) parts.push(原題);
  
  if (特典メモ) {
    parts.push(特典メモ.startsWith('※') ? 特典メモ : `※${特典メモ}`);
  } else if (設定.特典自動付与 && 形態 && 設定.形態別特典[形態]) {
    parts.push(設定.形態別特典[形態]);
  }
  return parts.join(' ');
}

function タイトルを生成(言語, カテゴリ, 形態, 日本語タイトル, 単巻数, セット開始, セット終了, 作者, 原題, 特典メモ) {
  return タイトルを段階生成(言語, カテゴリ, 形態, 日本語タイトル, 単巻数, セット開始, セット終了, 作者, 原題, 特典メモ);
}

/* ============================ 
 * ヘルパー
 * ============================ */
function 列番号を取得(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const name = String(headers[i] || '').trim();
    if (name) map[name] = i + 1;
  }
  return map;
}

function 言語マップを取得(ss) {
  const map = {};
  const sh = ss.getSheetByName(設定.言語マスター名);
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [n, c] of data) {
    const name = String(n || '').trim();
    const code = String(c || '').trim();
    if (name && code) map[name] = code;
  }
  return map;
}

function カテゴリマップを取得(ss) {
  const map = {};
  const sh = ss.getSheetByName(設定.カテゴリマスター名);
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [n, c] of data) {
    const name = String(n || '').trim();
    const code = String(c || '').trim();
    if (name && code) map[name] = code;
  }
  return map;
}

function 正規化(v) {
  return String(v || '').replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim();
}

function 数値変換(v) {
  const n = parseInt(String(v || '').trim(), 10);
  return Number.isFinite(n) ? n : null;
}

function データがある最終行を取得(sh, 列マップ) {
  const 基準列 = 列マップ[列名.日本語タイトル] || 1;
  const 最大行 = sh.getLastRow();
  if (最大行 <= 1) return 1;
  const data = sh.getRange(2, 基準列, 最大行 - 1, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== '' && data[i][0] !== null && data[i][0] !== undefined) return i + 2;
  }
  return 1;
}

/* ============================ 
 * WorksKey 正規化
 * ============================ */
function WorksKeyを作る(日本語タイトル, 作者) {
  return キー用正規化_(日本語タイトル) + '||' + キー用正規化_(作者);
}

function キー用正規化_(v) {
  let s = 正規化(v).toLowerCase();
  s = s
    .replace(/^著[:：]\s*/g, '')
    .replace(/^作[:：]\s*/g, '')
    .replace(/[［\[][^］\]]*[］\]]/g, '')
    .replace(/[（\(][^）\)]*[）\)]/g, '')
    .replace(/[｛\{][^｝\}]*[｝\}]/g, '')
    .replace(/[・･]/g, ' ')
    .replace(/[～〜~]/g, '')
    .replace(/[：:]/g, '')
    .replace(/[、,]/g, '')
    .replace(/[。\.]/g, '')
    .replace(/[！!]/g, '')
    .replace(/[？?]/g, '')
    .replace(/[『』「」]/g, '')
    .replace(/["'"'"]/g, '')
    .replace(/[‐−–—―]/g, '-')
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/\s*[\/／]\s*/g, ' / ')
    .replace(/\s+/g, ' ')
    .replace(/-+/g, '-')
    .trim();
  return s;
}

/* ============================ 
 * ④⑤⑥ マスター管理
 * ============================ */
function メニュー_Works初期化() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('警告', 'Worksを全削除します。続行？', ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(設定.作品シート名);
  if (sh) {
    const last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  } else {
    sh = ss.insertSheet(設定.作品シート名);
  }
  sh.getRange(1, 1, 1, 設定.作品ヘッダー.length).setValues([設定.作品ヘッダー]);
  ui.alert('✅ Works初期化完了');
}

function メニュー_マスター作成() {
  const ss = SpreadsheetApp.getActive();
  const created = [];
  
  if (!ss.getSheetByName(設定.言語マスター名)) {
    const sh = ss.insertSheet(設定.言語マスター名);
    sh.getRange(1, 1, 1, 設定.言語ヘッダー.length).setValues([設定.言語ヘッダー]);
    sh.getRange(2, 1, 設定.言語初期値.length, 設定.言語初期値[0].length).setValues(設定.言語初期値);
    created.push('言語');
  }
  if (!ss.getSheetByName(設定.カテゴリマスター名)) {
    const sh = ss.insertSheet(設定.カテゴリマスター名);
    sh.getRange(1, 1, 1, 設定.カテゴリヘッダー.length).setValues([設定.カテゴリヘッダー]);
    sh.getRange(2, 1, 設定.カテゴリ初期値.length, 設定.カテゴリ初期値[0].length).setValues(設定.カテゴリ初期値);
    created.push('カテゴリ');
  }
  if (!ss.getSheetByName(設定.作品シート名)) {
    const sh = ss.insertSheet(設定.作品シート名);
    sh.getRange(1, 1, 1, 設定.作品ヘッダー.length).setValues([設定.作品ヘッダー]);
    created.push('Works');
  }
  
  if (created.length > 0) {
    メニュー_プルダウン更新();
    SpreadsheetApp.getUi().alert(`✅ 作成: ${created.join(', ')}`);
  } else {
    SpreadsheetApp.getUi().alert('既に全て存在します');
  }
}

function メニュー_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.マスターシート名);
  if (!sh) return;
  
  const 列マップ = 列番号を取得(sh);
  const データ最終行 = データがある最終行を取得(sh, 列マップ);
  const 最終行 = Math.max(データ最終行 + 100, 200);
  
  const 言語データ = 言語データを色付きで取得(ss);
  if (言語データ.values.length > 0 && 列マップ[列名.言語]) {
    sh.getRange(2, 列マップ[列名.言語], 最終行 - 1, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(言語データ.values, true).build());
    条件付き書式をクリア(sh, 列マップ[列名.言語]);
    条件付き書式を設定(sh, 列マップ[列名.言語], 言語データ.values, 言語データ.colors, 最終行);
  }
  
  const カテゴリデータ = カテゴリデータを色付きで取得(ss);
  if (カテゴリデータ.values.length > 0 && 列マップ[列名.カテゴリ]) {
    sh.getRange(2, 列マップ[列名.カテゴリ], 最終行 - 1, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(カテゴリデータ.values, true).build());
    条件付き書式をクリア(sh, 列マップ[列名.カテゴリ]);
    条件付き書式を設定(sh, 列マップ[列名.カテゴリ], カテゴリデータ.values, カテゴリデータ.colors, 最終行);
  }
  
  SpreadsheetApp.getActive().toast('プルダウン更新完了', '商品コード管理', 3);
}

function 言語データを色付きで取得(ss) {
  const result = { values: [], colors: [] };
  const sh = ss.getSheetByName(設定.言語マスター名);
  if (!sh || sh.getLastRow() < 2) return result;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  let ci = 0;
  for (const row of data) {
    const n = String(row[0] || '').trim();
    if (!n) continue;
    result.values.push(n);
    result.colors.push(String(row[2] || '').trim() || 設定.色パレット[ci++ % 設定.色パレット.length]);
  }
  return result;
}

function カテゴリデータを色付きで取得(ss) {
  const result = { values: [], colors: [] };
  const sh = ss.getSheetByName(設定.カテゴリマスター名);
  if (!sh || sh.getLastRow() < 2) return result;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  let ci = 0;
  for (const row of data) {
    const n = String(row[0] || '').trim();
    if (!n) continue;
    result.values.push(n);
    result.colors.push(String(row[2] || '').trim() || 設定.色パレット[ci++ % 設定.色パレット.length]);
  }
  return result;
}

function 条件付き書式をクリア(sh, colNum) {
  const rules = sh.getConditionalFormatRules();
  const kept = [];
  for (const rule of rules) {
    const hit = rule.getRanges().some(rng => {
      const c1 = rng.getColumn();
      return (c1 <= colNum && colNum <= c1 + rng.getNumColumns() - 1);
    });
    if (!hit) kept.push(rule);
  }
  sh.setConditionalFormatRules(kept);
}

function 条件付き書式を設定(sh, colNum, values, colors, lastRow) {
  const rules = sh.getConditionalFormatRules();
  const range = sh.getRange(2, colNum, lastRow - 1, 1);
  for (let i = 0; i < values.length; i++) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(values[i])
        .setBackground(colors[i] || '#fff')
        .setRanges([range])
        .build()
    );
  }
  sh.setConditionalFormatRules(rules);
}

/* ============================ 
 * 自己書き込みループ防止（タイムスタンプ付き）
 * ============================ */
function 自己更新中か_() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('__SELF_EDIT_LOCK__') !== '1') return false;
  
  // ★ 10秒以上前のフラグは古いとみなしてクリア
  const ts = Number(props.getProperty('__SELF_EDIT_TS__') || '0');
  if (ts > 0 && (Date.now() - ts) > 10000) {
    props.deleteProperty('__SELF_EDIT_LOCK__');
    props.deleteProperty('__SELF_EDIT_TS__');
    return false;
  }
  return true;
}

function 自己更新を開始_() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('__SELF_EDIT_LOCK__', '1');
  props.setProperty('__SELF_EDIT_TS__', String(Date.now()));
}

function 自己更新を終了_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('__SELF_EDIT_LOCK__');
  props.deleteProperty('__SELF_EDIT_TS__');
}
