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

var NDA_DOCS_TEMPLATE_ID = '1InRRZT3XumIZCfN7k1fWl7dke0S1NPVnUqV9OdhGku4'
var NDA_DOCS_TEMPLATE_TITLE = 'Non Disclosure Agreement Template'

/**
 * Creates the Topbar menu in the spreadsheet. This function is fired every time a spreadsheet is open
 * @param {object} e - User and spreadsheet basic parameters
 */
function onOpen(e) {

    // Get the context
    var ui = SpreadsheetApp.getUi()

    // Create the menu
    var menu = ui.createMenu('NDA Generator')

    // Check if user is already authorized
    if (e && e.authMode === ScriptApp.AuthMode.LIMITED) {
        menu.addItem('Install', 'install')
    } else {
        menu.addItem('Run', 'run')
    }

    // Add the menu to the spreadsheet
    menu.addToUi()
}

/**
 * Makes a copy of the main NDA Docs template, creates the template folder,
 * moves the active spreadsheet into it, and creates a folder for the PDFs
 */
function install() {
    try {

        // Create the Solution folder on users Drive
        var folder = DriveApp.createFolder('NDA Generator')

        // Move the active spreadsheet
        var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId())
        this.moveFile(spreadsheetFile, folder)

        // Move the attached Google Form
        var form = FormApp.openByUrl(SpreadsheetApp.getActive().getFormUrl())
        this.moveFile(DriveApp.getFileById(form.getId()), folder)

        // Make a copy of NDA document
        var ndaTemplateFileId = DriveApp
            .getFileById(NDA_DOCS_TEMPLATE_ID)
            .makeCopy(NDA_DOCS_TEMPLATE_TITLE, folder)
            .getId()

        // Set NDA doc id on settings tab
        SpreadsheetApp.getActive()
            .getSheetByName('Settings')
            .getRange('C2')
            .setValue(ndaTemplateFileId)

        // Activate the trigger
        toggleTrigger()

    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)

    }
}

/**
 * Runs the main script function
 */
function run() {
    try {

        // Get NDA Doc template from settings sheet
        var templateId = SpreadsheetApp.getActive()
            .getSheetByName('Settings')
            .getRange('C2')
            .getValue()

        // Get spreadsheet timezone
        var timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone()

        // Get the form data to generate the NDAs
        var NDAFormData = this.getNDAFormData()

        // Get the repository folder
        var folder = getRepositoryFolder()

        // Get the form sheet
        var formSheet = SpreadsheetApp.getActive().getSheetByName('Form Responses')

        // Get link and status on sheet
        var statusColumn = getColumnIndexByName('NDA Status')
        var linkColumn = getColumnIndexByName('NDA Link PDF')

        NDAFormData.forEach(function(item) {

            // Get the current date formatted
            var dateFormatted = Utilities.formatDate(new Date(), timezone, 'MMM dd, yyyy HH:mm')

            // Generate the file name
            var fileName = 'NDA - ' + item.data['Full Name'] + ' - ' + dateFormatted

            // Copy the template file and save into the folder
            var driveFile = copyFile(templateId, fileName, folder)

            // Merge the texts
            mergeTexts(driveFile.getId(), item.data)

            // Save the PDF file on the repository folder
            var pdfFile = saveAsPDF(driveFile.getId(), fileName, folder)

            formSheet.getRange(item.rowIndex, statusColumn).setValue('sent')
            // formSheet.getRange(item.rowIndex, linkColumn).setValue(pdfFile.getURL())

            sendMail(item.data['Email Address'], item.data['Full Name'], pdfFile)

        })

    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)

    }
}

/**
 * Returns the Column index number based on name
 */
function getColumnIndexByName(ColumnName) {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Form Responses')
    var data = sheet.getDataRange().getValues()
    return data[0].indexOf(ColumnName) + 1
}

/**
 * Switch on the trigger that will run on form submit to generate the NDA
 */
function toggleTrigger() {
    try {

        var sheet = SpreadsheetApp.getActive()
        ScriptApp.newTrigger('run')
            .forSpreadsheet(sheet)
            .onFormSubmit()
            .create()

    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)
    }
}

/**
 * Gets the user settings
 * @returns {object}
 */
function getSettings() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings')
    return {
        template: sheet.getRange('C2').getValue(),
        email_subject: sheet.getRange('C4').getValue(),
        email_body: sheet.getRange('C6').getValue(),
        repository: sheet.getRange('C8').getValue()
    }
}

/**
 * Get the NDA From data formatted from range
 */
function getNDAFormData() {

    // Get all data
    var data = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('Form Responses')
        .getDataRange()
        .getValues()

    var jsonFormatted = this.parseDataToJsonArray(data)

    // Filter sent NDAs
    jsonFormatted = jsonFormatted
        .map(function(item, index) {
            return {
                rowIndex: index + 2,
                data: item
            }
        })
        .filter(function(item) {
        if (item.data['NDA Status'] !== 'sent') {
            delete item.data['Timestamp']
            delete item.data['NDA Status']
            delete item.data['NDA Link PDF']
            return item
        }
    })

    return jsonFormatted
}

/**
 * Get json formatted from data
 */
function parseDataToJsonArray(data) {

    var obj = {}
    var result = []
    var headers = data[0]
    var cols = headers.length
    var row = []

    for (var i = 1, l = data.length; i < l; i++) {
        // get a row to fill the object
        row = data[i]
        // clear object
        obj = {}
        for (var col = 0; col < cols; col++) {
            // fill object with new values
            obj[headers[col]] = row[col]
        }
        // add object in a final result
        result.push(obj)
    }

    return result
}

/**
 * Sends an email
 * @param {string} emailAddress A destination email address
 * @param {string} fullName A full name of email address owner
 * @param {File} attachment A DriveApp File instance
 */
function sendMail(emailAddress, fullName, attachment) {
    var emailSubject = SpreadsheetApp.getActive().getSheetByName('Settings').getRange('C4').getValue()
    var emailContent = SpreadsheetApp.getActive().getSheetByName('Settings').getRange('C6').getValue()

    // Parse line breaks
    emailContent = emailContent.replace(/(\r\n|\n|\r)/gm, '<br>')
    var body = '<p>Hello ' + fullName + ',</p><p>' + emailContent + '</p>'

    MailApp.sendEmail(emailAddress, emailSubject, body, {htmlBody: body, attachments: [attachment]})
}
