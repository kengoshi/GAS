function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🛠️ CSV操作")
    .addItem("フォルダから一括インポート", "importCsvsAndCleanup")
    .addSeparator() // 仕切り線
    .addItem("シート1以外をすべて削除", "deleteAllSheetsExceptFirst")
    .addToUi();
}

function importCsvsAndCleanup() {
  const folderId = "1OudlyBuyk1xm_PHsQIIYw3K5cZAFmDIc";
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.CSV);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    try {
      const blob = file.getBlob();
      const bytes = blob.getBytes();
      let csvText = "";

      // 1. UTF-8 BOMチェック (EF BB BF)
      if (
        bytes.length >= 3 &&
        (bytes[0] & 0xff) === 0xef &&
        (bytes[1] & 0xff) === 0xbb &&
        (bytes[2] & 0xff) === 0xbf
      ) {
        csvText = blob.getDataAsString("UTF-8").replace(/^\uFEFF/, "");
      } else {
        // 2. Shift-JISかUTF-8かの判定（バイナリレベルでチェック）
        let isSjis = false;
        for (let i = 0; i < bytes.length; i++) {
          const b = bytes[i] & 0xff;
          // Shift-JIS特有の範囲（半角カナや漢字の1バイト目）が含まれているか
          if (
            (b >= 0xa1 && b <= 0xdf) ||
            (b >= 0x81 && b <= 0x9f) ||
            (b >= 0xe0 && b <= 0xfc)
          ) {
            isSjis = true;
            break;
          }
        }

        if (isSjis) {
          csvText = blob.getDataAsString("Shift_JIS");
        } else {
          csvText = blob.getDataAsString("UTF-8");
        }
      }

      const csvData = Utilities.parseCsv(csvText);

      let sheet = ss.getSheetByName(fileName);
      if (sheet) {
        ss.deleteSheet(sheet);
      }
      sheet = ss.insertSheet(fileName);

      if (csvData.length > 0) {
        sheet
          .getRange(1, 1, csvData.length, csvData[0].length)
          .setValues(csvData);
      }

      file.setTrashed(true);
      count++;
    } catch (e) {
      Logger.log(fileName + " の処理中にエラーが発生しました: " + e.toString());
    }
  }

  if (count > 0) {
    SpreadsheetApp.getUi().alert(count + " 個のCSVをインポートしました。");
  } else {
    SpreadsheetApp.getUi().alert("処理するCSVファイルが見つかりませんでした。");
  }
}

function deleteAllSheetsExceptFirst() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const keepSheetName = "シート1"; // ここを、残したいシートの名前に書き換えてください

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "確認",
    "「" + keepSheetName + "」以外のすべてのシートを削除しますか？",
    ui.ButtonSet.YES_NO,
  );

  // 「はい」が押された場合のみ実行
  if (response == ui.Button.YES) {
    let deletedCount = 0;
    for (let i = 0; i < sheets.length; i++) {
      let sheetName = sheets[i].getName();
      if (sheetName !== keepSheetName) {
        ss.deleteSheet(sheets[i]);
        deletedCount++;
      }
    }
    ui.alert(deletedCount + " 個のシートを削除しました。");
  }
}
