/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;

    document.getElementById("addTableItem").onclick = addTableItem;
    document.getElementById("addRangeItem").onclick = addRangeItem;
    document.getElementById("chkAddToSetting").onclick = togglePropertyBag;
    document.getElementById("loadTableData").onclick = loadTableData;
    document.getElementById("loadRangeData").onclick = loadRangeData;

    createTable();
    createRange();
  }
});

var enablePropertyBag = false;
var tableRowCount = 0;
var rangeRow = 1;
var tableData = [];
var rangeData = [];

export async function togglePropertyBag() {
  enablePropertyBag = !enablePropertyBag;
}

async function createTable() {
  try {
    await Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var table = sheet.tables.getItemOrNullObject("ItemTable");
      table.load(["rows/count"]);
      await context.sync();

      if (table.isNullObject) {
        table = sheet.tables.add("A1:J1", true /*hasHeaders*/);
        table.name = "ItemTable";
        table.getHeaderRowRange().values = [
          ["ID", "Row Id", "Order", "Item Name", "Item Type", "Start Date", "End Date", "Duration", "Progress", "Work"]
        ];
        tableRowCount = 0;
      } else {
        tableRowCount = table.rows.count;
      }

      await context.sync();

      console.log("table row count", tableRowCount);
    });
  } catch (error) {
    console.error(error);
  }
}

async function createRange() {
  try {
    await Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getRange("L1:U1");
      range.values = [
        ["ID", "Row Id", "Order", "Item Name", "Item Type", "Start Date", "End Date", "Duration", "Progress", "Work"]
      ];

      await context.sync();

      rangeRow = 2;

      console.log("range created");
    });
  } catch (error) {
    console.error(error);
  }
}

export async function addTableItem() {
  try {
    Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var table = sheet.tables.getItemOrNullObject("ItemTable");

      let range: Excel.Range;

      if (tableRowCount == 0) {
        tableRowCount = 1;
        range = table.getDataBodyRange();
      } else {
        tableRowCount = tableRowCount + 1;
        range = table.getDataBodyRange().getRowsBelow(1);
      }

      range.load("address");
      await context.sync();

      var item = [
        tableRowCount,
        "=Row()",
        tableRowCount,
        "Item " + tableRowCount,
        "Task",
        "09/25/2020",
        "09/26/2020",
        1,
        0,
        0
      ];

      range.values = [item];

      if (enablePropertyBag) {
        tableData.push(item);
        Office.context.document.settings.set("GanttData1", tableData);
        Office.context.document.settings.saveAsync();
      }

      await context.sync();

      console.log("new row range address", range.address);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function addRangeItem() {
  try {
    Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getRange(`L${rangeRow}:U${rangeRow}`);

      var item = [rangeRow, "=Row()", rangeRow, "Item " + rangeRow, "Task", "09/25/2020", "09/26/2020", 1, 0, 0];
      range.values = [item];

      if (enablePropertyBag) {
        rangeData.push(item);
        Office.context.document.settings.set("GanttData2", rangeData);
        Office.context.document.settings.saveAsync();
      }

      await context.sync();

      console.log("new row range done");
      this.rangeRow = this.rangeRow + 1;
    });
  } catch (error) {
    console.error(error);
  }
}

export async function loadTableData() {
  var data = Office.context.document.settings.get("GanttData1");
  document.getElementById("dataString").innerText = JSON.stringify(data);
}

export async function loadRangeData() {
  var data = Office.context.document.settings.get("GanttData2");
  document.getElementById("dataString").innerText = JSON.stringify(data);
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

