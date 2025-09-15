/**
 * kintone to スプレッドシート書き出し - 使用例
 * 
 * 事前に以下をスクリプトプロパティで設定してください：
 * - KINTONE_SUBDOMAIN
 * - KINTONE_APP_ID  
 * - KINTONE_API_TOKEN
 */

// === 設定値の取得 ===
const CONFIG = {
  subdomain: PropertiesService.getScriptProperties().getProperty('KINTONE_SUBDOMAIN'),
  appId: PropertiesService.getScriptProperties().getProperty('KINTONE_APP_ID'),
  apiToken: PropertiesService.getScriptProperties().getProperty('KINTONE_API_TOKEN')
};

/**
 * 基本的な使用例：全データを出力
 */
async function exportAllData() {
  const exporter = new KintoneSpreadsheetExporter(CONFIG);
  
  try {
    const result = await exporter.exportToSheet();
    console.log(`エクスポート完了: ${result.recordCount}件`);
  } catch (error) {
    console.error('エクスポートエラー:', error);
    // メール通知などのエラー処理をここに
  }
}

/**
 * 特定フィールドのみを出力
 */
function exportSelectedFields() {
  const exporter = new KintoneSpreadsheetExporter(CONFIG, {
    sheetName: '限定フィールド'
  });

  // 必要なフィールドのみ指定
  const fields = ['タイトル', 'ステータス', '担当者', '期日'];

  exporter.export(fields);
}

/**
 * カスタムオプション付きの例
 */
function exportWithOptions() {
  const exporter = new KintoneToSheetExporter(CONFIG, {
    sheetName: 'kintone_backup',
    batchSize: 300,        // 一回の取得件数を300件に
    sleepMs: 200          // API間隔を200msに
  });

  exporter.export();
}

/**
 * 複数アプリを順次処理する例
 */
function exportMultipleApps() {
  const apps = [
    { appId: '123', sheetName: 'タスク管理' },
    { appId: '456', sheetName: '顧客管理' },
    { appId: '789', sheetName: '商品マスタ' }
  ];

  apps.forEach(app => {
    try {
      const config = { ...CONFIG, appId: app.appId };
      const exporter = new KintoneToSheetExporter(config, {
        sheetName: app.sheetName
      });

      const result = exporter.export();
      console.log(`${app.sheetName}: ${result.recordCount}件完了`);

      // アプリ間で少し待機
      Utilities.sleep(1000);

    } catch (error) {
      console.error(`${app.sheetName}でエラー:`, error);
      // 一つ失敗しても続行
    }
  });
}

/**
 * エラーハンドリング強化版
 */
function exportWithErrorHandling() {
  const exporter = new KintoneToSheetExporter(CONFIG);

  try {
    // 設定確認
    exporter.showConfig();

    // 実行
    const result = exporter.export();

    // 成功通知
    const message = `kintone同期完了\n件数: ${result.recordCount}\n時間: ${result.duration}秒`;
    console.log(message);

    // 必要に応じてSlackやメール通知
    // sendNotification(message);

  } catch (error) {
    const errorMessage = `kintone同期エラー\n${error.message}\n\nスタック:\n${error.stack}`;
    console.error(errorMessage);

    // エラー通知
    sendErrorNotification(errorMessage);
  }
}

/**
 * 定期実行用（トリガーから呼ばれる想定）
 */
function scheduledExport() {
  // 実行時間を制限（GASの実行時間制限対策）
  const startTime = new Date();
  const maxDuration = 5 * 60 * 1000; // 5分

  try {
    const exporter = new KintoneToSheetExporter(CONFIG, {
      sheetName: `データ_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm')}`
    });

    const result = exporter.export();

    const endTime = new Date();
    const duration = endTime - startTime;

    if (duration > maxDuration * 0.8) { // 80%超えたら警告
      console.warn(`実行時間注意: ${duration / 1000}秒`);
    }

    console.log(`定期実行完了: ${result.recordCount}件`);

  } catch (error) {
    console.error('定期実行エラー:', error);
    // 重要なデータなので管理者に即座に通知
    sendUrgentNotification(error);
  }
}

// === 初期設定用の関数 ===

/**
 * スクリプトプロパティにAPIトークン等を設定
 * ※実際の値に置き換えて実行してください
 */
function setupProperties() {
  const props = PropertiesService.getScriptProperties();

  // 実際の値に置き換えてください
  props.setProperties({
    'KINTONE_SUBDOMAIN': 'your-subdomain',
    'KINTONE_APP_ID': '123',
    'KINTONE_API_TOKEN': 'your-api-token-here'
  });

  console.log('スクリプトプロパティを設定しました');
}

/**
 * 定期実行のトリガーを設定
 */
function setupScheduledTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scheduledExport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎日午前9時に実行するトリガーを作成
  ScriptApp.newTrigger('scheduledExport')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  console.log('定期実行トリガーを設定しました（毎日9時）');
}

// === 通知関数（実装例）===

function sendErrorNotification(message) {
  // Gmail通知の例
  try {
    GmailApp.sendEmail(
      'admin@yourcompany.com',
      'kintone同期エラー',
      message
    );
  } catch (e) {
    console.error('メール送信失敗:', e);
  }
}

function sendUrgentNotification(error) {
  // Slack Webhook通知の例
  try {
    const payload = {
      text: `🚨 kintone同期で緊急エラー\n${error.message}`,
      channel: '#alerts'
    };

    // Slack Webhook URL（実際のURLに置き換え）
    const webhookUrl = 'https://hooks.slack.com/services/YOUR/SLACK/WEBHOOK';

    UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
  } catch (e) {
    console.error('Slack通知失敗:', e);
  }
}