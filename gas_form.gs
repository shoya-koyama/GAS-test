function addingUp() {
  const existingForm = FormApp.openById('***'); // 取得したいフォームの ID（隠します。）
  const formResponses = existingForm.getResponses();

  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(formResponses.length + 2, 1, 1, 1)
    .setValue('集計')
    .setBackground('#000')
    .setFontColor('#fff')
    .setFontWeight('bold');
  sheet.getRange(formResponses.length + 3, 1, 1, 2).setValues([
    ['件数', formResponses.length]
  ]);

  const dataMap = new Map();
  const range = sheet.getRange(`D2:D${formResponses.length + 1}`);
  const values = range.getValues();
  for (const gyou of values) {
    const data = gyou[0];

    if (dataMap.has(data)) {
      dataMap.set(data, dataMap.get(data) + 1);
    } else {
      dataMap.set(data, 1);
    }
  }

  sheet.getRange(formResponses.length + 4, 1, 1, 2).setValues([
    ['評価', '合計値']
  ]);
  const dataArray = Array.from(dataMap, ([key, value]) => [key, value]);
  sheet.getRange(formResponses.length + 5, 1, dataArray.length, 2).setValues(dataArray);
  
  pieChart(sheet, formResponses.length + 2,3, formResponses.length + 4,1, dataArray.length,2);
}

function pieChart(sheet, writeX, writeY, dataX, dataY, rowsMany, columnsMany) {
  const range = sheet.getRange(dataX, dataY, rowsMany, columnsMany);
  const chart = sheet.newChart()
    .asPieChart()
    .addRange(range)
    .setNumHeaders(1)
    .setPosition(writeX,writeY, 0,0)
    .setOption('title','アンケート');

  sheet.insertChart(chart.build());
}
