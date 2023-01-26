/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = show;
    document.getElementById("go").onclick = run;
  }
});


function show(){
  create();
  let info = document.getElementById("data");
  info.classList.remove("ms-input");
  info.classList.add("show");
  let header = document.getElementById("app-body");
  header.style.display = "none";
}

export async function run() {
  var myHeaders = new Headers();
myHeaders.append("apikey", "dsNHdl0pG4ZJMDYqtw95r02ZiXY1zafB");

var requestOptions = {
  method: 'GET',
  redirect: 'follow',
  headers: myHeaders
};

fetch("https://api.apilayer.com/exchangerates_data/convert?to=BRL&from=USD&amount=1", requestOptions)
  .then(response => response.text())
  .then(result => get_resp(result))
  .catch(error => console.log('error', error));
}

async function create(result) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add("COTACAO");
    const Current_table = sheet.tables.add("A1:C1", true);

    Current_table.getHeaderRowRange().format.fill.color="7B828A";
    Current_table.getDataBodyRange().format.fill.color="7B828A";
    Current_table.getHeaderRowRange().format.font.color="black";
    Current_table.getHeaderRowRange().format.font.size=14;
    Current_table.getHeaderRowRange().format.autofitRows;
    Current_table.getHeaderRowRange().format.autofitColumns;
    Current_table.getHeaderRowRange().values = [["DATA", "REAL", "DÓLAR"]];
    Current_table.name = "DATA_NEW";
    await context.sync();
  })};

  async function get_resp(result) {
    await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject("COTACAO");
    let teste = JSON.parse(result);
    let algo = [];
    for (var i in teste)
      algo.push(i, teste[i]);
    const expensesTable = sheet.tables.getItemOrNullObject("DATA_NEW");
    
    const newData = algo[7];
    const amount = "1";
    // const newValue = amount;
    const newChange = algo[9];
    expensesTable.rows.add(
      null, [[newData, amount, newChange]]
    );

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();

    await context.sync();
  });
}

// export async function getSize() {
// await Excel.run(async (context) => {
//   let sheet = context.workbook.worksheets.getActiveWorksheet();

//   let range = sheet.getUsedRange();
//   range.load("address");
//   await context.sync();
//   const teste = JSON.stringify(range.address, null, 4).split('!');
//   console.log(teste[1]);
//   }).catch(function(error) {
//   console.log("Error: " + error);
//   if (error instanceof OfficeExtension.Error) {
//     console.log("Debug info: " + JSON.stringify(error.debugInfo));
//   }
// });
//}


