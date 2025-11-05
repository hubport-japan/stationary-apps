const BATCH_SIZE = 50; // 一度に処理するURLの数

const ERROR_LOG_SHEET_NAME = 'エラーログ'; // エラーログを書き出すシート名

const CANCEL_RETURN_LOG_SHEET_NAME = 'キャンセル・返品ログ'; // キャンセル・返品ログを書き出すシート名

// 複数の関数で利用する定数をグローバルスコープに定義

const MASTER_SHEET_ID = '1ljdiPI2zRbyMnGMTIx5YKm4YGtwkeDz6njpJxxdYIz4'; // 利用者シートIDが記載されたマスターシート

const TARGET_SHEET_NAME = '配送管理'; // 各利用者シート内の対象シート名

/**
 * メインの実行関数。
 * 確認ダイアログを表示し、OKなら全件処理を開始する。
 */
function runFullUpdate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'キャンセル・返品チェック',
    'すべての利用者シートをチェックし、ステータスを更新します。\n処理には数分かかる場合があります。開始しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  // 実行記録をリセットして最初から開始
  PropertiesService.getScriptProperties().deleteProperty('LAST_PROCESSED_INDEX');
  
  // 1回目のバッチ処理を即時実行
  updateOrderStatus();
}

/**
 * 注文IDをキーにして、各担当者の管理シートから注文ステータスを取得し、
 * アクティブシートに転記します。（バッチ処理本体）
 */
function updateOrderStatus() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastProcessedIndex = parseInt(scriptProperties.getProperty('LAST_PROCESSED_INDEX') || '0');

  if (lastProcessedIndex === 0) {
    console.log('キャンセル・返品チェック処理を最初から開始します。');
    prepareErrorLogSheet();
    prepareCancelReturnLogSheet();
    // 最初の処理開始時に、アクティブシートのAA列をクリア
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = activeSheet.getLastRow();
    if (lastRow > 1) {
      activeSheet.getRange(2, 27, lastRow - 1, 1).clearContent(); // AA列をクリア
    }
  } else {
    console.log(`キャンセル・返品チェック処理を ${lastProcessedIndex} 件目の続きから開始します。`);
  }

  const orderStatusMap = new Map();
  const cancelReturnOrders = []; // キャンセル・返品の注文IDを記録する配列
  let endIndex = 0;
  let totalIds = 0;
  let idData = [];
  const failedIds = [];

  // 前回のバッチで収集したキャンセル・返品情報を読み込む（続きから処理する場合）
  let previousOrderStatusMap = new Map();
  if (lastProcessedIndex > 0) {
    const storedMapJson = scriptProperties.getProperty('ORDER_STATUS_MAP');
    if (storedMapJson) {
      try {
        const storedMap = JSON.parse(storedMapJson);
        storedMap.forEach(([orderId, status]) => {
          previousOrderStatusMap.set(orderId, status);
        });
        console.log(`前回のバッチから ${previousOrderStatusMap.size} 件のキャンセル・返品情報を読み込みました。`);
      } catch (e) {
        console.warn('前回のキャンセル・返品情報の読み込みに失敗しました。');
      }
    }
  }

  // 前回の情報を現在のMapに統合
  previousOrderStatusMap.forEach((status, orderId) => {
    orderStatusMap.set(orderId, status);
  });

  try {
    const masterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const masterSheet = masterSs.getSheets()[0]; 
    idData = masterSheet.getRange('D2:D' + masterSheet.getLastRow()).getValues();
    totalIds = idData.filter(row => row[0]).length;

    console.log(`全有効ID数: ${totalIds}件`);

    if (lastProcessedIndex >= idData.length && idData.length > 0) {
      SpreadsheetApp.getUi().alert('すべてのチェックが完了しました。');
      return;
    }

    endIndex = Math.min(lastProcessedIndex + BATCH_SIZE, idData.length);
    console.log(`${lastProcessedIndex + 1}行目から${endIndex}行目までを処理します。`);

    for (let i = lastProcessedIndex; i < endIndex; i++) {
      const id = idData[i][0] ? String(idData[i][0]).trim() : '';
      if (!id) continue;

      try {
        const targetSs = SpreadsheetApp.openById(id);
        const targetSheet = targetSs.getSheets().find(sheet => sheet.getName().trim() === TARGET_SHEET_NAME);

        if (!targetSheet) {
          console.warn(`シートが見つかりません: スプレッドシート名「${targetSs.getName()}」`);
          failedIds.push([new Date(), id, `シート「${TARGET_SHEET_NAME}」が見つかりません`]);
          continue;
        }

        const data = targetSheet.getDataRange().getValues();
        const displayData = targetSheet.getDataRange().getDisplayValues(); // 表示されている値も取得
        const sheetName = targetSs.getName(); // シート名を取得

        for (let j = 1; j < data.length; j++) {
          // C列（インデックス2）が注文ID、K列（インデックス10）がステータス
          // 表示値と実際の値の両方を確認
          let orderId = displayData[j][2] ? String(displayData[j][2]).trim() : null;
          let status = displayData[j][10] ? String(displayData[j][10]).trim() : null;
          
          // 表示値が空の場合は実際の値を使用
          if (!orderId || orderId === '') {
            orderId = data[j][2] ? String(data[j][2]).trim() : null;
          }
          if (!status || status === '') {
            status = data[j][10] ? String(data[j][10]).trim() : null;
          }

          // 注文IDとステータスの正規化（改行、タブ、その他の空白文字を除去）
          if (orderId) {
            orderId = orderId.replace(/[\r\n\t]/g, '').trim();
          }
          if (status) {
            status = status.replace(/[\r\n\t]/g, '').trim();
          }

          if (!orderId || !status) continue;

          // キャンセル・返品の判定（大文字小文字を区別しない、部分一致）
          const statusLower = status.toLowerCase();
          if (statusLower.includes('キャンセル') || statusLower.includes('返品') || 
              statusLower.includes('cancel') || statusLower.includes('返品済')) {
            orderStatusMap.set(orderId, status);
            
            // キャンセル・返品の注文IDを記録（シート名、注文ID、ステータス、日時）
            cancelReturnOrders.push([
              new Date(),
              sheetName,
              orderId,
              status
            ]);
            
            // 特定の注文IDが見つかった場合にデバッグログを出力
            if (orderId.includes('249-2182985-0003060')) {
              console.log(`*** デバッグ: 注文ID [${orderId}] を検出しました。ステータス: [${status}] ***`);
            }
            
            console.log(`追加: 注文ID [${orderId}] のステータス [${status}] を収集しました。`);
          }
        }
      } catch (e) {
        console.error(`IDの処理中にエラーが発生しました。ID: ${id} - エラー: ${e.message}`);
        failedIds.push([new Date(), id, e.message]);
      }
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert(`マスタースプレッドシート(ID: ${MASTER_SHEET_ID})を開けませんでした。\n\nエラー: ${e.message}`);
    console.error(`マスタースプレッドシートを開けませんでした。エラー: ${e.message}`);
    return;
  }

  if (failedIds.length > 0) {
    logErrorsToSheet(failedIds);
  }

  // キャンセル・返品の注文IDをログシートに記録
  if (cancelReturnOrders.length > 0) {
    logCancelReturnOrders(cancelReturnOrders);
  }

  // --- 次の処理の予約 ---
  if (endIndex < idData.length) {
    // 次のバッチ処理のために、現在収集したキャンセル・返品情報を保存
    const mapArray = Array.from(orderStatusMap.entries());
    scriptProperties.setProperty('ORDER_STATUS_MAP', JSON.stringify(mapArray));
    scriptProperties.setProperty('LAST_PROCESSED_INDEX', endIndex);
    
    // 既存のトリガーを削除してから新しいトリガーを作成
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'updateOrderStatus') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 1秒後に次のバッチ処理を自動実行する
    ScriptApp.newTrigger('updateOrderStatus').timeBased().after(1000).create();
    
    console.log(`バッチ処理進行中: ${endIndex}/${idData.length}件処理済み。現在までに${orderStatusMap.size}件のキャンセル・返品を検出。`);
  } else {
    // すべての処理が完了したら、アクティブシートに最終結果を書き込む
    console.log('全利用者シートのチェックが完了しました。アクティブシートに結果を書き込みます...');
    updateActiveSheet(orderStatusMap);
    
    // 保存した情報をクリア
    scriptProperties.deleteProperty('LAST_PROCESSED_INDEX');
    scriptProperties.deleteProperty('ORDER_STATUS_MAP');
    
    // 完了時にトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'updateOrderStatus') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    console.log('すべての行のチェックが完了しました。');
    
    const finalMessage = `すべてのチェックが完了しました。\nキャンセル・返品の注文IDはAA列に表示されています。\n詳細は実行ログや「${ERROR_LOG_SHEET_NAME}」シート、「${CANCEL_RETURN_LOG_SHEET_NAME}」シートを確認してください。`;
    SpreadsheetApp.getUi().alert(finalMessage);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('住所チェック実行', 'checkAddressData')
    .addSeparator()
    .addItem('キャンセル・返品チェック', 'runFullUpdate')
    .addToUi();
}

/**
 * エラーログシートを準備する
 */
function prepareErrorLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ERROR_LOG_SHEET_NAME);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(ERROR_LOG_SHEET_NAME);
  }
  sheet.getRange('A1:C1').setValues([['日時', 'エラーが発生したID', 'エラー内容']]).setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 300);
}

/**
 * キャンセル・返品ログシートを準備する
 */
function prepareCancelReturnLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CANCEL_RETURN_LOG_SHEET_NAME);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CANCEL_RETURN_LOG_SHEET_NAME);
  }
  sheet.getRange('A1:D1').setValues([['日時', '利用者シート名', '注文ID', 'ステータス']]).setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 150);
  
  // 見やすくするためにフィルタを設定
  sheet.getRange(1, 1, 1, 4).setBackground('#4285f4').setFontColor('#ffffff');
}

/**
 * エラーをログシートに記録する
 */
function logErrorsToSheet(errors) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ERROR_LOG_SHEET_NAME);
  if (sheet && errors.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, errors.length, 3).setValues(errors);
  }
}

/**
 * キャンセル・返品の注文IDをログシートに記録する
 */
function logCancelReturnOrders(orders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CANCEL_RETURN_LOG_SHEET_NAME);
  if (sheet && orders.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, orders.length, 4).setValues(orders);
  }
}

/**
 * アクティブシートに注文ステータスを更新する
 * アクティブシートのS列（インデックス18）の注文IDがキャンセル・返品かどうかを判定し、AA列（27列目）に結果を書き込む
 */
function updateActiveSheet(orderStatusMap) {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = activeSheet.getDataRange();
  const values = range.getValues();
  const displayValues = range.getDisplayValues(); // 表示されている値も取得
  let updatedCount = 0;
  let cancelReturnCount = 0;

  console.log('アクティブシートの更新を開始します...');
  console.log(`収集したキャンセル・返品の注文ID数: ${orderStatusMap.size}件`);

  // AA列に結果を書き込むための配列を準備
  const results = [];
  
  // デバッグ用：収集した注文IDのリストを出力
  const collectedOrderIds = Array.from(orderStatusMap.keys());
  console.log(`収集した注文IDのサンプル（最初の10件）: ${collectedOrderIds.slice(0, 10).join(', ')}`);
  
  for (let i = 1; i < values.length; i++) {
    // S列（インデックス18）が注文ID - 表示値と実際の値の両方を確認
    let orderIdFromSheet = displayValues[i][18] ? String(displayValues[i][18]).trim() : null;
    
    // 表示値が空の場合は実際の値を使用
    if (!orderIdFromSheet || orderIdFromSheet === '') {
      orderIdFromSheet = values[i][18] ? String(values[i][18]).trim() : null;
    }
    
    // 注文IDの正規化（改行、タブ、その他の空白文字を除去）
    if (orderIdFromSheet) {
      orderIdFromSheet = orderIdFromSheet.replace(/[\r\n\t]/g, '').trim();
    }
    
    if (!orderIdFromSheet) {
      // 注文IDが空の場合は空文字を設定
      results.push(['']);
      continue;
    }

    // orderStatusMapに注文IDが存在するかチェック
    const isFound = orderStatusMap.has(orderIdFromSheet);
    
    if (isFound) {
      const status = orderStatusMap.get(orderIdFromSheet);
      results.push([status]);
      cancelReturnCount++;
      console.log(`更新: ${i + 1}行目の注文ID [${orderIdFromSheet}] のステータスを [${status}] に更新しました。`);
    } else {
      // キャンセル・返品でない場合は空文字を設定
      results.push(['']);
      // デバッグ用：特定の注文IDが見つからない場合にログ出力
      if (orderIdFromSheet === '249-2182985-0003060') {
        console.warn(`デバッグ: 注文ID [${orderIdFromSheet}] がorderStatusMapに見つかりませんでした。`);
        console.warn(`orderStatusMapには ${orderStatusMap.size} 件のエントリがあります。`);
        console.warn(`類似の注文IDがあるか確認: ${collectedOrderIds.filter(id => id.includes('249-2182985-0003060')).join(', ')}`);
      }
    }
    updatedCount++;
  }

  // AA列に一括で書き込み
  if (results.length > 0) {
    activeSheet.getRange(2, 27, results.length, 1).setValues(results);
  }

  console.log(`処理完了。全${updatedCount}件チェック済み。うち${cancelReturnCount}件がキャンセル・返品でした。`);
  return updatedCount;
}

function checkAddressData() {
  const SPREADSHEET_ID_MASTER = '1XP8DIqBi-pGK8Xp44KGe0CQOOoc7TW6WrvClR3gMg7Q';
  const SHEET_NAME_MASTER = '注文マスタ';

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const ui = SpreadsheetApp.getUi();

    const masterSs = SpreadsheetApp.openById(SPREADSHEET_ID_MASTER);
    const masterSheet = masterSs.getSheetByName(SHEET_NAME_MASTER);

    if (!masterSheet) {
      ui.alert(`エラー: 照合先のシート「${SHEET_NAME_MASTER}」が見つかりません。`);
      return;
    }

    const activeLastRow = activeSheet.getLastRow();
    if (activeLastRow < 2) {
      ui.alert('アクティブシートにチェック対象のデータがありません。（ヘッダー行を除く）');
      return;
    }

    const activeData = activeSheet.getRange(2, 14, activeLastRow - 1, 11).getValues();
    const masterLastRow = masterSheet.getLastRow();

    if (masterLastRow < 2) {
      ui.alert('注文マスタにデータがありません。');
      return;
    }

    const masterData = masterSheet.getRange(2, 1, masterLastRow - 1, masterSheet.getLastColumn()).getValues();
    const masterMap = new Map();

    masterData.forEach(row => {
      const orderId = row[2];
      if (orderId) {
        masterMap.set(String(orderId).trim(), {
          productName: String(row[5]).trim(),
          name: String(row[34]).trim(),
          zip: String(row[35]).trim(),
          address: String(row[36]).trim(),
          phone: String(row[37]).trim()
        });
      }
    });

    // --- 新規追加ロジック: 各利用者の配送管理シートからN列の値を収集 ---
    const userSheetNColumnMap = new Map();
    let userIdData = [];

    try {
      const userIdMasterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
      const userIdMasterSheet = userIdMasterSs.getSheets()[0];
      userIdData = userIdMasterSheet.getRange('D2:D' + userIdMasterSheet.getLastRow()).getValues();

      console.log(`ユーザーシートIDマスタから ${userIdData.length} 件のIDを読み込みました。`);

      for (const row of userIdData) {
        const userId = row[0] ? String(row[0]).trim() : '';
        if (!userId) continue;

        try {
          const targetSs = SpreadsheetApp.openById(userId);
          const targetSheet = targetSs.getSheets().find(sheet => sheet.getName().trim() === TARGET_SHEET_NAME);

          if (!targetSheet) {
            console.warn(`ユーザーシートが見つかりません: スプレッドシート名「${targetSs.getName()}」, シート名「${TARGET_SHEET_NAME}」`);
            continue;
          }

          const dataRange = targetSheet.getDataRange();
          const values = dataRange.getValues();
          const displayValues = dataRange.getDisplayValues();

          for (let j = 1; j < values.length; j++) {
            const orderId = values[j][2] ? String(values[j][2]).trim() : null;
            const nColumnValue = displayValues[j][13] ? String(displayValues[j][13]).trim() : null;

            if (orderId && nColumnValue) {
              userSheetNColumnMap.set(orderId, nColumnValue);
            }
          }
        } catch (e) {
          console.error(`ユーザーシートIDの処理中にエラーが発生しました。ID: ${userId} - エラー: ${e.message}`);
        }
      }
    } catch (e) {
      console.error(`ユーザーシートIDマスタースプレッドシート(ID: ${MASTER_SHEET_ID})を開けませんでした。エラー: ${e.message}`);
    }
    // --- 新規追加ロジック終了 ---

    const results = [];

    for (const row of activeData) {
      const productName = String(row[0]).trim();
      const orderId = String(row[5]).trim();
      const name = String(row[7]).trim();
      const zip = String(row[8]).trim();
      const address = String(row[9]).trim();
      const phone = String(row[10]).trim();

      if (!orderId) {
        results.push(['注文IDが空です', '']);
        continue;
      }

      const masterRecord = masterMap.get(orderId);

      if (!masterRecord) {
        results.push(['注文IDなし', '']);
      } else {
        const mismatchedItems = [];

        const normalizePhone = (phoneStr) => {
          if (!phoneStr) return '';
          let normalized = phoneStr.replace(/[-－ー\s]/g, '');
          if (normalized.length === 10 && !normalized.startsWith('0')) {
            normalized = '0' + normalized;
          }
          return normalized;
        };
        
        const normalizeAddress = (addrStr) => {
          if (!addrStr) return '';
          return addrStr.normalize('NFKC').replace(/\s+/g, '');
        };

        const activePhoneNormalized = normalizePhone(phone);
        const masterPhoneNormalized = normalizePhone(masterRecord.phone);
        const activeAddressNormalized = normalizeAddress(address);
        const masterAddressNormalized = normalizeAddress(masterRecord.address);
        
        if (masterRecord.productName !== productName) mismatchedItems.push('商品名');
        if (masterRecord.name !== name) mismatchedItems.push('氏名');
        if (masterRecord.zip !== zip) mismatchedItems.push('郵便番号');
        if (masterAddressNormalized !== masterAddressNormalized) mismatchedItems.push('住所');
        if (activePhoneNormalized !== masterPhoneNormalized) mismatchedItems.push('電話番号');

        if (mismatchedItems.length === 0) {
          results.push(['一致', '']);
        } else {
          results.push(['不一致', mismatchedItems.join(', ')]);
        }
      }
    }
    
    if (results.length > 0) {
      activeSheet.getRange(2, 25, results.length, 2).setValues(results);
    }
    
    // --- 新規追加ロジック: AA列にN列の値を転記 ---
    const valuesForAA = [];

    for (let i = 0; i < activeData.length; i++) {
      const orderId = String(activeData[i][5]).trim();
      valuesForAA.push([userSheetNColumnMap.get(orderId) || '']);
    }

    if (valuesForAA.length > 0) {
      const targetRange = activeSheet.getRange(2, 27, valuesForAA.length, 1);
      targetRange.setNumberFormat('@');
      targetRange.setValues(valuesForAA);
    }
    // --- 新規追加ロジック終了 ---

    ui.alert('住所チェックが完了しました。');
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert('処理中にエラーが発生しました。\n' + e.message);
  }
}

