function sendMultiEmails() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const startRow = 2;
  const numRows = 2;
  const dataRange = sheet.getRange(startRow, 1, numRows, 4);
  const values = dataRange.getValues();

  for (const i in values) {
    const userName = values[i][0];
    const emailAddress = values[i][1];
    const subject = values[i][2];
    const award = values[i][3];

    const message = `${userName} さん、こんにちは。
怪しいものです。

コンテストの結果は見事 ${award} に選ばれました。`;

    MailApp.sendEmail(
      emailAddress, subject, message,
      {cc: 'shoyakoyama06@gmail.com'}
    );
  }
}

function columnChart() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("A6:D14");
  const chart = sheet.newChart();
  const columnChart = chart.asColumnChart()
    .addRange(range)
    .setNumHeaders(1)
    .setPosition(15,1,0,0)
    .setOption('title', '売りあげ')
    .build();
  sheet.insertChart(columnChart);
}

function lineChart() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("A6:D14");
  const chart = sheet.newChart();
  const lineChart = chart.asLineChart()
    .addRange(range)
    .setNumHeaders(1)
    .setPosition(33,1,0,0)
    .setOption('title', '売りあげ')
    .build();
  sheet.insertChart(lineChart);
}

function pieChart() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("A6:B14");
  const chart = sheet.newChart();
  const pieChart = chart.asPieChart()
    .addRange(range)
    .setNumHeaders(1)
    .setPosition(51,1,0,0)
    .setOption('title', '売りあげ')
    .build();
  sheet.insertChart(pieChart);
}

function addup() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("B7:D14");
  const values = range.getValues();
  const totals = [['合計']];
  for (const rowNumber in values) {
    let total = 0;

    for (const colNumber in values[rowNumber]) {
      total += values[rowNumber][colNumber];
    }
    totals.push([total]);
  }
  const outPutRange = sheet.getRange("E6:E14");
  outPutRange.setValues(totals);
}
