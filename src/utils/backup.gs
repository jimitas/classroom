/**
 * バックアップ・復元モジュール
 * @module Backup
 */

/**
 * クラスマスタをバックアップ
 * @returns {string} バックアップシート名
 */
function backupClassMaster() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const classMasterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CLASS_MASTER);

    if (!classMasterSheet) {
      throw new Error(`シート「${CONFIG.SHEET_NAMES.CLASS_MASTER}」が見つかりません`);
    }

    // タイムスタンプ付きバックアップ名
    const timestamp = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmmss");
    const backupSheetName = `バックアップ_${timestamp}`;

    // クラスマスタを複製
    const backupSheet = classMasterSheet.copyTo(ss);
    backupSheet.setName(backupSheetName);

    // バックアップシートを最後に移動
    ss.moveActiveSheet(ss.getNumSheets());

    console.log(`✓ バックアップ作成完了: ${backupSheetName}`);
    debugLog(`バックアップシート: ${backupSheetName}`);

    return backupSheetName;
  } catch (error) {
    console.error("バックアップ作成エラー:", error);
    throw error;
  }
}

/**
 * 指定されたバックアップからクラスマスタを復元
 * @param {string} backupSheetName - バックアップシート名
 */
function restoreClassMaster(backupSheetName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const backupSheet = ss.getSheetByName(backupSheetName);

    if (!backupSheet) {
      throw new Error(`バックアップシート「${backupSheetName}」が見つかりません`);
    }

    const classMasterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CLASS_MASTER);

    if (!classMasterSheet) {
      throw new Error(`シート「${CONFIG.SHEET_NAMES.CLASS_MASTER}」が見つかりません`);
    }

    // クラスマスタをクリア
    classMasterSheet.clear();

    // バックアップからデータをコピー
    const backupData = backupSheet.getDataRange().getValues();
    if (backupData.length > 0) {
      classMasterSheet.getRange(1, 1, backupData.length, backupData[0].length).setValues(backupData);
    }

    // フォーマットもコピー
    const sourceFormat = backupSheet.getDataRange();
    const targetFormat = classMasterSheet.getRange(1, 1, backupData.length, backupData[0].length);
    sourceFormat.copyFormatToRange(classMasterSheet, 1, backupData[0].length, 1, backupData.length);

    console.log(`✓ 復元完了: ${backupSheetName} → ${CONFIG.SHEET_NAMES.CLASS_MASTER}`);
    debugLog(`復元元: ${backupSheetName}`);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `バックアップ「${backupSheetName}」から復元しました`,
      '復元完了',
      5
    );
  } catch (error) {
    console.error("復元エラー:", error);
    throw error;
  }
}

/**
 * 全バックアップシートを一覧表示
 * @returns {Array<Object>} バックアップ情報の配列 [{name, created}]
 */
function listBackups() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const backups = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith("バックアップ_")) {
        // タイムスタンプを抽出
        const timestampStr = sheetName.replace("バックアップ_", "");
        backups.push({
          name: sheetName,
          timestamp: timestampStr
        });
      }
    });

    // タイムスタンプで降順ソート（新しい順）
    backups.sort((a, b) => b.timestamp.localeCompare(a.timestamp));

    console.log("=".repeat(60));
    console.log("バックアップ一覧");
    console.log("=".repeat(60));
    if (backups.length === 0) {
      console.log("バックアップはありません");
    } else {
      backups.forEach((backup, index) => {
        console.log(`${index + 1}. ${backup.name}`);
      });
    }
    console.log("=".repeat(60));

    return backups;
  } catch (error) {
    console.error("バックアップ一覧取得エラー:", error);
    throw error;
  }
}

/**
 * 古いバックアップを削除（指定件数以上のバックアップを削除）
 * @param {number} keepCount - 保持するバックアップ数（デフォルト: 5）
 */
function cleanupOldBackups(keepCount = 5) {
  try {
    const backups = listBackups();

    if (backups.length <= keepCount) {
      console.log(`バックアップ数: ${backups.length}件（削除不要）`);
      return;
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const toDelete = backups.slice(keepCount);

    console.log(`古いバックアップを削除: ${toDelete.length}件`);

    toDelete.forEach(backup => {
      const sheet = ss.getSheetByName(backup.name);
      if (sheet) {
        ss.deleteSheet(sheet);
        console.log(`✓ 削除: ${backup.name}`);
      }
    });

    console.log(`✓ クリーンアップ完了: ${backups.length - keepCount}件削除、${keepCount}件保持`);
  } catch (error) {
    console.error("バックアップクリーンアップエラー:", error);
    throw error;
  }
}

/**
 * 最新のバックアップから復元
 */
function restoreFromLatestBackup() {
  const backups = listBackups();

  if (backups.length === 0) {
    throw new Error("復元可能なバックアップがありません");
  }

  const latestBackup = backups[0];
  console.log(`最新のバックアップから復元: ${latestBackup.name}`);

  restoreClassMaster(latestBackup.name);
}
