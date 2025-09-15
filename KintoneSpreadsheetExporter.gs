/**
 * kintone から Google Spreadsheet にデータを書き出すクラス（レコード閲覧権限のみ対応）
 * 
 * 使用例:
 * const exporter = new KintoneSpreadsheetExporter({
 *   subdomain: 'your-subdomain',
 *   appId: '123',
 *   apiToken: 'your-api-token'
 * });
 * 
 * await exporter.exportToSheet();
 */
class KintoneSpreadsheetExporter {
  /**
   * @param {Object} config - kintone接続設定
   * @param {string} config.subdomain - kintoneのサブドメイン
   * @param {string|number} config.appId - アプリID
   * @param {string} config.apiToken - APIトークン
   * @param {Object} [options] - オプション設定
   * @param {string} [options.sheetName='kintoneデータ'] - 出力先シート名
   * @param {number} [options.batchSize=500] - 一回の取得件数
   * @param {number} [options.sleepMs=100] - API呼び出し間隔(ms)
   * @param {boolean} [options.enableStyling=true] - ヘッダー装飾の有効/無効
   * @param {string[]} [options.excludeFields] - 除外するフィールド（システムフィールドなど）
   */
  constructor(config, options = {}) {
    this._validateConfig(config);

    this.config = {
      subdomain: config.subdomain,
      appId: String(config.appId),
      apiToken: config.apiToken,
      baseUrl: `https://${config.subdomain}.cybozu.com`
    };

    this.options = {
      sheetName: options.sheetName || 'kintoneデータ',
      batchSize: Math.min(options.batchSize || 500, 500),
      sleepMs: options.sleepMs || 100,
      enableStyling: options.enableStyling !== false,
      // デフォルトで除外するシステムフィールド
      excludeFields: options.excludeFields || [
        '$revision', '__REVISION__'
      ],
      ...options
    };

    // キャッシュ用プライベートプロパティ
    this._totalCount = null;
    this._fieldMap = null; // レコードから動的に構築

    this._log('KintoneSpreadsheetExporter initialized (ReadOnly mode)');
  }

  /**
   * メイン実行関数：kintoneからデータを取得してスプレッドシートに書き出し
   * @param {string[]} [fieldCodes] - 取得するフィールドコードの配列
   * @returns {Promise<ExportResult>}
   */
  async exportToSheet(fieldCodes = null) {
    try {
      this._log('=== Export Started ===');
      const startTime = new Date();

      // 1. スプレッドシートの準備
      const sheet = this._prepareSheet();

      // 2. データ取得（フィールド情報取得をスキップ）
      this._log('データ取得開始...');
      const records = await this.getAllRecords(fieldCodes);

      if (records.length === 0) {
        this._log('取得できるレコードがありません');
        return this._createResult(0, 0, 0);
      }

      // 3. レコードからフィールドマップを構築
      const fieldMap = this._buildFieldMapFromRecords(records);

      // 4. 対象フィールドの決定
      const targetFields = this._determineTargetFields(fieldMap, fieldCodes);
      this._log(`対象フィールド: ${targetFields.length}個`);

      // 5. データ変換と書き込み
      this._writeRecordsToSheet(sheet, records, fieldMap, targetFields);

      const endTime = new Date();
      const duration = (endTime - startTime) / 1000;
      const result = this._createResult(records.length, targetFields.length, duration);

      this._log(`=== Export Completed ===`);
      this._log(`処理時間: ${duration}秒, レコード数: ${records.length}, フィールド数: ${targetFields.length}`);

      return result;

    } catch (error) {
      this._log(`Export失敗: ${error.message}`);
      throw new Error(`kintone書き出し処理でエラーが発生しました: ${error.message}`);
    }
  }

  /**
   * 指定したフィールドのみでエクスポート
   * @param {string[]} fieldCodes - フィールドコードの配列
   * @returns {Promise<ExportResult>}
   */
  async exportSelectedFields(fieldCodes) {
    if (!Array.isArray(fieldCodes) || fieldCodes.length === 0) {
      throw new Error('フィールドコードの配列を指定してください');
    }
    return this.exportToSheet(fieldCodes);
  }

  /**
   * 増分取得（更新日時ベース）
   * @param {Date|string} since - この日時以降のレコードを取得
   * @param {string[]} [fieldCodes] - 取得するフィールドコードの配列
   * @returns {Promise<ExportResult>}
   */
  async exportIncremental(since, fieldCodes = null) {
    try {
      const sinceISO = since instanceof Date ? since.toISOString() : since;
      this._log(`増分取得開始: ${sinceISO} 以降`);

      const sheet = this._prepareSheet();

      // 更新日時条件付きでレコード取得
      const query = `更新日時 > "${sinceISO}" order by 更新日時 asc, $id asc`;
      const records = await this._getRecordsByQuery(query, fieldCodes);

      if (records.length === 0) {
        this._log(`${sinceISO} 以降に更新されたレコードはありません`);
        return this._createResult(0, 0, 0);
      }

      const fieldMap = this._buildFieldMapFromRecords(records);
      const targetFields = this._determineTargetFields(fieldMap, fieldCodes);

      this._writeRecordsToSheet(sheet, records, fieldMap, targetFields);
      this._log(`増分取得完了: ${records.length}件`);

      return this._createResult(records.length, targetFields.length, 0);

    } catch (error) {
      throw new Error(`増分取得でエラーが発生しました: ${error.message}`);
    }
  }

  /**
   * 設定情報の表示
   */
  showConfig() {
    this._log('=== Current Configuration ===');
    this._log(`Subdomain: ${this.config.subdomain}`);
    this._log(`App ID: ${this.config.appId}`);
    this._log(`Sheet Name: ${this.options.sheetName}`);
    this._log(`Batch Size: ${this.options.batchSize}`);
    this._log(`Sleep: ${this.options.sleepMs}ms`);
    this._log(`Styling: ${this.options.enableStyling ? 'Enabled' : 'Disabled'}`);
    this._log(`Exclude Fields: ${this.options.excludeFields.join(', ')}`);
  }

  /**
   * kintoneから全レコードを取得（自動的にページネーション方式を選択）
   * @param {string[]} [fieldCodes] - 取得するフィールドコードの配列
   * @returns {Promise<Object[]>} レコードの配列
   */
  async getAllRecords(fieldCodes = null) {
    const totalCount = await this.getTotalCount();
    this._log(`総レコード数: ${totalCount}件`);

    if (totalCount <= 10000) {
      return this._getAllRecordsByOffset(fieldCodes);
    } else {
      return this._getAllRecordsByIdSeek(fieldCodes);
    }
  }

  /**
   * 総レコード数を取得（キャッシュ付き）
   * @returns {Promise<number>}
   */
  async getTotalCount() {
    if (this._totalCount !== null) {
      return this._totalCount;
    }

    const { body } = await this._kintoneRequest('/k/v1/records.json', {
      app: this.config.appId,
      query: 'limit 1',
      totalCount: true
    });

    this._totalCount = parseInt(body.totalCount || 0);
    return this._totalCount;
  }

  // === プライベートメソッド ===

  /**
   * 設定値の検証
   * @private
   */
  _validateConfig(config) {
    if (!config || typeof config !== 'object') {
      throw new Error('設定オブジェクトが必要です');
    }

    const required = ['subdomain', 'appId', 'apiToken'];
    const missing = required.filter(key => !config[key]);

    if (missing.length > 0) {
      throw new Error(`必須設定が不足しています: ${missing.join(', ')}`);
    }
  }

  /**
   * 結果オブジェクトの生成
   * @private
   */
  _createResult(recordCount, fieldCount, duration) {
    return {
      recordCount,
      fieldCount,
      duration,
      success: true,
      timestamp: new Date()
    };
  }

  /**
   * スプレッドシートの準備（取得または作成）
   * @private
   */
  _prepareSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(this.options.sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(this.options.sheetName);
      this._log(`新しいシート "${this.options.sheetName}" を作成しました`);
    } else {
      sheet.clear();
      this._log(`既存シート "${this.options.sheetName}" をクリアしました`);
    }

    return sheet;
  }

  /**
   * レコード群からフィールドマップを動的に構築
   * @private
   */
  _buildFieldMapFromRecords(records) {
    if (this._fieldMap) return this._fieldMap;

    const fieldMap = {};

    // 全レコードを調べてフィールド情報を収集
    records.forEach(record => {
      Object.keys(record).forEach(fieldCode => {
        if (!fieldMap[fieldCode]) {
          const field = record[fieldCode];
          fieldMap[fieldCode] = {
            code: fieldCode,
            type: field.type,
            // ラベルはフィールドコードをそのまま使用（権限不足でfield情報が取得できないため）
            label: this._generateFieldLabel(fieldCode, field.type)
          };
        }
      });
    });

    this._fieldMap = fieldMap;
    this._log(`フィールドマップ構築完了: ${Object.keys(fieldMap).length}個`);
    return fieldMap;
  }

  /**
   * フィールドコードから表示用ラベルを生成
   * @private
   */
  _generateFieldLabel(fieldCode, fieldType) {
    // システムフィールドの日本語名マッピング
    const systemFieldLabels = {
      '$id': 'レコードID',
      '更新日時': '更新日時',
      '作成日時': '作成日時',
      '更新者': '更新者',
      '作成者': '作成者',
      'レコード番号': 'レコード番号',
      '$revision': 'リビジョン'
    };

    return systemFieldLabels[fieldCode] || fieldCode;
  }

  /**
   * 取得対象フィールドの決定
   * @private
   */
  _determineTargetFields(fieldMap, specifiedFields) {
    const allFieldCodes = Object.keys(fieldMap).filter(
      fieldCode => !this.options.excludeFields.includes(fieldCode)
    );

    if (specifiedFields && Array.isArray(specifiedFields)) {
      const validFields = specifiedFields.filter(field => allFieldCodes.includes(field));
      const invalidFields = specifiedFields.filter(field => !allFieldCodes.includes(field));

      if (invalidFields.length > 0) {
        this._log(`警告: 存在しないフィールドをスキップ: ${invalidFields.join(', ')}`);
      }

      if (validFields.length === 0) {
        throw new Error('有効なフィールドが指定されていません');
      }

      return validFields;
    }

    // デフォルト: レコードID、更新日時を先頭にして、それ以外を追加
    const priorityFields = ['$id', '更新日時'];
    const otherFields = allFieldCodes.filter(code => !priorityFields.includes(code));

    return [...priorityFields.filter(field => allFieldCodes.includes(field)), ...otherFields];
  }

  /**
   * offset方式でレコード取得（10,000件以下）
   * @private
   */
  async _getAllRecordsByOffset(fieldCodes = null) {
    const limit = this.options.batchSize;
    let offset = 0;
    const allRecords = [];

    while (true) {
      const params = {
        app: this.config.appId,
        query: `order by $id asc limit ${limit} offset ${offset}`
      };

      if (fieldCodes && fieldCodes.length > 0) {
        params.fields = fieldCodes;
      }

      const { body } = await this._kintoneRequest('/k/v1/records.json', params);
      const chunk = body.records || [];
      allRecords.push(...chunk);

      this._log(`取得済み: ${allRecords.length}件 (offset: ${offset})`);

      if (chunk.length < limit) break;
      offset += limit;
      if (offset >= 10000) {
        this._log('警告: offset制限(10,000)に達しました');
        break;
      }

      await this._sleep(this.options.sleepMs);
    }

    this._log(`offset方式完了: ${allRecords.length}件`);
    return allRecords;
  }

  /**
   * レコードID方式でレコード取得（10,000件超対応）
   * @private
   */
  async _getAllRecordsByIdSeek(fieldCodes = null) {
    const limit = this.options.batchSize;
    let lastRecordId = 0;
    const allRecords = [];

    while (true) {
      const params = {
        app: this.config.appId,
        query: `$id > ${lastRecordId} order by $id asc limit ${limit}`
      };

      if (fieldCodes && fieldCodes.length > 0) {
        params.fields = fieldCodes;
      }

      const { body } = await this._kintoneRequest('/k/v1/records.json', params);
      const chunk = body.records || [];
      if (chunk.length === 0) break;

      allRecords.push(...chunk);
      lastRecordId = parseInt(chunk[chunk.length - 1].$id.value);

      this._log(`取得済み: ${allRecords.length}件 (lastId: ${lastRecordId})`);

      if (chunk.length < limit) break;
      await this._sleep(this.options.sleepMs);
    }

    this._log(`レコードID方式完了: ${allRecords.length}件`);
    return allRecords;
  }

  /**
   * クエリ指定でレコード取得
   * @private
   */
  async _getRecordsByQuery(query, fieldCodes = null) {
    const limit = this.options.batchSize;
    const allRecords = [];
    let offset = 0;

    while (true) {
      const params = {
        app: this.config.appId,
        query: `${query} limit ${limit} offset ${offset}`
      };

      if (fieldCodes && fieldCodes.length > 0) {
        params.fields = fieldCodes;
      }

      const { body } = await this._kintoneRequest('/k/v1/records.json', params);
      const chunk = body.records || [];
      if (chunk.length === 0) break;

      allRecords.push(...chunk);
      this._log(`取得済み: ${allRecords.length}件`);

      if (chunk.length < limit) break;
      offset += limit;
      await this._sleep(this.options.sleepMs);
    }

    return allRecords;
  }

  /**
   * スプレッドシートにデータを書き込み
   * @private
   */
  _writeRecordsToSheet(sheet, records, fieldMap, targetFields) {
    // ヘッダー行の作成
    const headers = this._buildHeaders(fieldMap, targetFields);

    // データ行の作成
    const dataRows = records.map(record => this._buildDataRow(record, fieldMap, targetFields));

    // 一括書き込み
    const allData = [headers, ...dataRows];
    const range = sheet.getRange(1, 1, allData.length, allData[0].length);
    range.setValues(allData);

    // ヘッダー行の装飾
    if (this.options.enableStyling) {
      this._styleHeaderRow(sheet, headers.length);
    }

    // 列幅の自動調整
    sheet.autoResizeColumns(1, headers.length);

    this._log(`スプレッドシート書き込み完了: ${dataRows.length}行 × ${headers.length}列`);
  }

  /**
   * ヘッダー行の構築
   * @private
   */
  _buildHeaders(fieldMap, targetFields) {
    return targetFields.map(fieldCode => {
      return fieldMap[fieldCode]?.label || fieldCode;
    });
  }

  /**
   * データ行の構築
   * @private
   */
  _buildDataRow(record, fieldMap, targetFields) {
    const row = [];

    targetFields.forEach(fieldCode => {
      const field = record[fieldCode];
      const fieldType = fieldMap[fieldCode]?.type;
      const value = this._formatFieldValue(field, fieldType);
      row.push(value);
    });

    return row;
  }

  /**
   * ヘッダー行のスタイル設定
   * @private
   */
  _styleHeaderRow(sheet, columnCount) {
    const headerRange = sheet.getRange(1, 1, 1, columnCount);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
  }

  /**
   * フィールドタイプに応じた値のフォーマット
   * @private
   */
  _formatFieldValue(field, fieldType) {
    if (!field || field.value === null || field.value === undefined) {
      return '';
    }

    switch (fieldType) {
      case 'SINGLE_LINE_TEXT':
      case 'MULTI_LINE_TEXT':
      case 'RICH_TEXT':
      case 'NUMBER':
      case 'CALC':
      case 'LINK':
      case 'RECORD_NUMBER':
        return field.value;

      case 'DATE':
      case 'DATETIME':
      case 'CREATED_TIME':
      case 'UPDATED_TIME':
        return field.value ? new Date(field.value) : '';

      case 'TIME':
        return field.value;

      case 'DROP_DOWN':
      case 'RADIO_BUTTON':
        return field.value;

      case 'CHECK_BOX':
      case 'MULTI_SELECT':
        return Array.isArray(field.value) ? field.value.join(', ') : field.value;

      case 'USER_SELECT':
      case 'ORGANIZATION_SELECT':
      case 'GROUP_SELECT':
      case 'CREATOR':
      case 'MODIFIER':
        if (Array.isArray(field.value)) {
          return field.value.map(item => item.name || item.code).join(', ');
        }
        return field.value?.name || field.value?.code || '';

      case 'FILE':
        if (Array.isArray(field.value)) {
          return field.value.map(file => file.name).join(', ');
        }
        return '';

      case 'SUBTABLE':
        return `[サブテーブル: ${field.value?.length || 0}行]`;

      case '__ID__':
      case '__REVISION__':
        return field.value;

      default:
        return field.value?.toString() || '';
    }
  }

  /**
   * kintone REST API呼び出し共通処理
   * @private
   */
  async _kintoneRequest(path, params = {}, method = 'GET') {
    const normalized = { ...params };
    if (Array.isArray(normalized.fields)) {
      normalized.fields = normalized.fields.join(',');
    }

    const url = method === 'GET'
      ? `${this.config.baseUrl}${path}?` + Object.keys(normalized)
        .map(key => `${key}=${encodeURIComponent(normalized[key])}`)
        .join('&')
      : `${this.config.baseUrl}${path}`;

    const options = {
      method,
      headers: {
        'X-Cybozu-API-Token': this.config.apiToken
        // 'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    // POSTの場合のみContent-Typeを追加
    if (method !== 'GET') {
      options.headers['Content-Type'] = 'application/json';
      options.payload = JSON.stringify(params);
    }

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    let body = {};
    try {
      body = JSON.parse(response.getContentText());
    } catch (e) {
      // JSONパースエラーは無視
    }

    if (responseCode !== 200) {
      const message = body?.message || body?.code || 'Unknown error';
      throw new Error(`kintone API error [${responseCode}] ${message} @ ${path}`);
    }

    return { code: responseCode, body };
  }

  /**
   * 待機処理
   * @private
   */
  async _sleep(ms) {
    return new Promise(resolve => {
      Utilities.sleep(ms);
      resolve();
    });
  }

  /**
   * ログ出力
   * @private
   */
  _log(message) {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] KintoneSpreadsheetExporter: ${message}`);
  }
}