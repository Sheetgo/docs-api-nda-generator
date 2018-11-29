/**
 * Copyright (c) 2018 Sheetgo Europe, S.L.
 *
 * This source code is licensed under the MIT License (MIT) found in the LICENSE file in the
 * root directory of this source tree or on: https://opensource.org/licenses/MIT
 *
 *
 * @link https://github.com/Sheetgo/docs-api-nda-generator
 * @version 1.0.0
 * @licence MIT
 *
 */

/**
 * Get a cell value from a sheet based on a range
 * @param {string} sheetName - Name of the sheet that has the data
 * @param {string} cell - Name of the sheet range. Eg.: 'C2'
 * @returns {object}
 */
function getValue(sheetName, cell) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    return sheet.getRange(cell).getValue()
}

/**
 * Set values into a sheet base on a range
 * @param {string} sheetName - Name of the sheet that will receive the data
 * @param {array[]} data - The data to save in matrix format
 * @param {int} firstRow - The starting range row cell
 * @param {int} firstColumn - The starting range column cell
 * @param {boolean} [clearData] - If has to clear the old data
 * @param {boolean} [keepData] - If true the new data will be appended below the previous data
 */
function setValues(sheetName, data, firstRow, firstColumn, clearData, keepData) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    var lastRow = sheet.getLastRow()
    if (clearData && lastRow > firstRow) {
        var range = sheet.getRange(firstRow, firstColumn, lastRow - 1, data[0].length)
        range.clearContent()
    }
    if (keepData) {
        firstRow = lastRow + 1
    }
    var newRange = sheet.getRange(firstRow, firstColumn, data.length, data[0].length)
    newRange.setValues(data)
}

/**
 * Show an UI dialog
 * @param {string} title - Dialog title
 * @param {string} message - Dialog message
 */
function showUiDialog(title, message) {
    try {
        var ui = SpreadsheetApp.getUi()
        ui.alert(title, message, ui.ButtonSet.OK)
    } catch (e) {
        // pass
    }
}
