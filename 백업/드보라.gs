function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("금전출납부")
    .addItem("다음 달 만들기", "createNextMonth")
    .addToUi();
}

function createNextMonth() {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = activeSpreadsheet.getSheets();
  let previousSheet = sheets[0];
  let previousSheetName = previousSheet.getSheetName();
  let [year, month] = previousSheetName.split(" ");
  year = Number(year.slice(0, -1));
  month = Number(month.slice(0, -1)) + 1;

  if (month === 13) {
    let ui = SpreadsheetApp.getUi();

    if (
      ui.alert(
        `'${year + 1}년 1월'을 만들까요? '아니요'를 선택하면 '다음 달 만들기'가 취소됩니다.`,
        ui.ButtonSet.YES_NO
      ) === ui.Button.YES
    ) {
      year++;
      month = 1;
    } else {
      return;
    }
  }

  let newSheet = sheets[sheets.length - 1].copyTo(activeSpreadsheet);
  let newName = `${year}년 ${month}월`;
  newSheet.setName(newName);
  newSheet
    .getRange("A1")
    .setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText(`금전 출납부 (${newName})`)
        .setTextStyle(0, 7, SpreadsheetApp.newTextStyle().setForegroundColor("#000").build())
        .build()
    );
  let maxRows = previousSheet.getMaxRows();
  newSheet.getRange("D3").setValue(`='${previousSheetName}'!D${maxRows}`);
  newSheet.getRange("E3").setValue(`='${previousSheetName}'!E${maxRows}`);
  newSheet.getRange("F3").setValue(`='${previousSheetName}'!F${maxRows}`);
  activeSpreadsheet.setActiveSheet(newSheet);
  activeSpreadsheet.moveActiveSheet(0);
}

function onEdit(e) {
  let range = e.range;
  let sheet = range.getSheet();
  let rowIndex = range.getRowIndex();

  if (sheet.getRange(rowIndex + 2, 3).getValue() === "장      계") {
    sheet.insertRowAfter(rowIndex + 1);
    sheet.getRange(rowIndex, 6).setValue(`=F${rowIndex - 1}+D${rowIndex}-E${rowIndex}`);
  }

  let maxRows = sheet.getMaxRows();
  sheet.getRange(maxRows - 1, 4).setValue(`=SUM(D4:D${maxRows - 4})`);
  sheet.getRange(maxRows - 1, 5).setValue(`=SUM(E4:E${maxRows - 4})`);
  sheet.getRange(maxRows, 4).setValue(`=D3+D${maxRows - 1}`);
  sheet.getRange(maxRows, 5).setValue(`=E3+E${maxRows - 1}`);
  sheet.getRange(maxRows, 6).setValue(`=F${maxRows - 4}`);
}
