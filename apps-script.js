/* 
    Transform moneydj financial statements to Google Spreadsheets
    Copyright (C) 2022 Yang Yuanzhi <yangyuanzhi@gmail.com>

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

// var moneydj = "https://dj.mybank.com.tw";
var moneydj = "https://stockchannelnew.sinotrade.com.tw";
// var moneydj = "http://pscnetinvest.moneydj.com.tw";
// var moneydj = "http://jsjustweb.jihsun.com.tw";
// var moneydj = "https://moneydj.emega.com.tw";

function getContentTextFromCache(url) {
  var cache = CacheService.getDocumentCache();
  var xml = cache.get(url);
  if (xml != null && xml != []) {
    return xml;
  }

  xml = UrlFetchApp.fetch(url).getContentText("big5");
  if (xml != null && xml != []) {
    cache.put(url, xml, 21600);
  }
  return xml;
}

function TWPRICE(code) { 
  var url = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=otc_"+code+".tw";
  var json = getContentTextFromCache(url);
  var data = JSON.parse(json);
  return data.msgArray[0].z;
}

function setCurrentCell(spreadSheet) {
    var sheet = spreadSheet.getSheetByName("價值評估");
    var cell = sheet.getRange("B1");
    sheet.setCurrentCell(cell);
}

function fixup(xml, remove_redundant_form_end_tag) {
  xml = '<!DOCTYPE html>' + xml.substring(11);
  xml = xml.replace('content="IE=edge">', 'content="IE=edge"/>');
  xml = xml.replace('CONTENT="text/html; charset=big5">','CONTENT="text/html; charset=big5"/>');
  xml = xml.replace("stylesheet", '"stylesheet"');
  xml = xml.replace('type="text/css">', 'type="text/css"/>');
  xml = xml.replace('class="t01" border="0"', 'class="t01"');
  xml = xml.replace(/&/g,"&amp;");
  xml = xml.replace("selected>", 'selected="1">');
  if (remove_redundant_form_end_tag) {
    xml = xml.replace("</FORM>", "");
  }
  xml = xml.replace("</form>", "</FORM>");
  xml = xml.replace("<BR>", "<BR/>");
  return xml;
}

function financialReport(spreadSheet, url, sheetName, remove_redundant_form_end_tag) {
  var xml = fixup(getContentTextFromCache(url), remove_redundant_form_end_tag);
  var document = XmlService.parse(xml);
  var root = document.getRootElement();
  var ns = root.getNamespace();
  var table = root.getChild('body', ns).getChild("div", ns).getChild('table', ns).getChildren('tr', ns)[1].getChildren('td', ns)[1].getChild('FORM', ns).getChild('table', ns);
  var rows = table.getChild('tr', ns).getChild('td', ns).getChildren('div', ns);
  var table_title = rows[0].getChild('div', ns).getChild('div', ns).getText() +
                    rows[0].getChild('div', ns).getChild('div', ns).getChild('div', ns).getText();
  var currentSheet = spreadSheet.getSheetByName(sheetName);
  currentSheet.getRange("A1:I200").clearContent();
  var cell = currentSheet.getRange(1,1);
  cell.setValue(table_title);
  // Logger.log(table_title);
  for (var row = 1; row < rows.length; row++)
  {
    var columns = rows[row].getChildren('span', ns);
    for (var col = 0; col < columns.length; col++)
    {
      cell = currentSheet.getRange(row + 1, col + 1);
      cell.setValue(columns[col].getText());
      // Logger.log(columns[col].getText());
    }
  }
}

function getStock(spreadSheet) {
  return spreadSheet.getSheetByName("價值評估").getRange("B1").getValue();
}

function quarterIncomeStatement(spreadSheet, stock) {
  var url = moneydj + '/z/zc/zcq/zcq0.djhtm?b=Q&a=' + stock;
  financialReport(spreadSheet, url, "季損益表", true);
}

function yearIncomeStatement(spreadSheet, stock) {
  var url = moneydj + '/z/zc/zcq/zcq0.djhtm?b=Y&a=' + stock;
  financialReport(spreadSheet, url, "年損益表", true);
}

function quarterBalanceSheet(spreadSheet, stock) {
  var url = moneydj + '/z/zc/zcp/zcpa/zcpa0.djhtm?b=Q&a=' + stock;
  financialReport(spreadSheet, url, "季資產負債表", false);
}

function yearBalanceSheet(spreadSheet, stock) {
  var url = moneydj + '/z/zc/zcp/zcpa/zcpa0.djhtm?b=Y&a=' + stock;
  financialReport(spreadSheet, url, "年資產負債表", false);
}

function quarterCashFlow(spreadSheet, stock) {
  var url = moneydj + '/z/zc/zc30.djhtm?b=Q&a=' + stock;
  financialReport(spreadSheet, url, "季現金流量表", false);
}

function yearCashFlow(spreadSheet, stock) {
  var url = moneydj + '/z/zc/zc30.djhtm?b=Y&a=' + stock;
  financialReport(spreadSheet, url, "年現金流量表", false);
}

function onMyOpen(e) {
  var spreadSheet = e.source;
  var stock = getStock(spreadSheet);
  quarterIncomeStatement(spreadSheet, stock);
  yearIncomeStatement(spreadSheet, stock);
  quarterBalanceSheet(spreadSheet, stock);
  yearBalanceSheet(spreadSheet, stock);
  quarterCashFlow(spreadSheet, stock);
  yearCashFlow(spreadSheet, stock);
  setCurrentCell(spreadSheet);
}

function onMyEdit(e) {
  var spreadSheet = e.source;
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  if (spreadSheet.getActiveSheet().getSheetName() == "價值評估" && row == 1 && col == 2) {
    quarterIncomeStatement(spreadSheet, e.value);
    yearIncomeStatement(spreadSheet, e.value);
    quarterBalanceSheet(spreadSheet, e.value);
    yearBalanceSheet(spreadSheet, e.value);
    quarterCashFlow(spreadSheet, e.value);
    yearCashFlow(spreadSheet, e.value);
    setCurrentCell(spreadSheet);
    var ui = SpreadsheetApp.getUi();
    // ui.alert("active sheet="+spreadSheet.getActiveSheet().getName());
    // ui.alert("row=" + row + " col=" + col + " value=" + e.value);
    ui.alert("財務報表更新完畢！");
  }
}
