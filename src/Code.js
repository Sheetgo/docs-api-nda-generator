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

var NDA_SETTINGS_SPREADSHEET = '15cA7P0Tna5bN0smgNu16XDatoFf5_YogJl6PlM1PdV8'
var NDA_DOCS_TEMPLATE_ID = '1InRRZT3XumIZCfN7k1fWl7dke0S1NPVnUqV9OdhGku4'
var NDA_DOCS_TEMPLATE_TITLE = 'Non Disclosure Agreement Template'
var NDA_SETTINGS = {}

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
    if (isInstalled()) {
        menu.addItem('Run', 'run')
    } else {
        menu.addItem('Install', 'showInstallDialog')
    }

    // Add the menu to the spreadsheet
    menu.addToUi()
}

/**
 * Checks if the solution is installed by getting the Settings sheet
 * @returns {boolean}
 */
function isInstalled() {

    // Get the current spreadsheet
    var activeSpreadsheet = SpreadsheetApp.getActive()

    // Get settings sheet
    var sheet = activeSpreadsheet.getSheetByName('Settings')

    return sheet !== null
}

/**
 * Loads the showDialog
 */
function showInstallDialog() {
    var html = HtmlService.createTemplateFromFile('Dialog.html')
    html = html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(220).setHeight(120)
    SpreadsheetApp.getUi().showModalDialog(html, "Automated NDA Generator")
}

/**
 * Makes a copy of the main NDA Docs template, creates the template folder,
 * moves the active spreadsheet into it, and creates a folder for the PDFs
 */
function install() {

    // Get the current spreadsheet
    var activeSpreadsheet = SpreadsheetApp.getActive()

    // Copy the template settings sheet
    SpreadsheetApp
        .openById(NDA_SETTINGS_SPREADSHEET)
        .getSheetByName('Settings')
        .copyTo(activeSpreadsheet)
        .setName('Settings')

    // Create the Solution folder on users Drive
    var folder = createFolder('Automated NDA Generator')

    // Move the active spreadsheet
    moveFile(activeSpreadsheet.getId(), folder)

    // Move the attached Google Form
    moveFile(FormApp.openByUrl(activeSpreadsheet.getFormUrl()).getId(), folder)

    // Make a copy of NDA document
    var ndaTemplateFile = copyFile(NDA_DOCS_TEMPLATE_ID, NDA_DOCS_TEMPLATE_TITLE, folder)

    // Set NDA doc id on settings tab
    getNamedRange('template_id').setValue(ndaTemplateFile.getId())

    // Activate the trigger
    toggleTrigger()

    // Update the Topbar menu
    onOpen()

    // Return NDA folder URL
    return folder.getUrl()

}

/**
 * Generate the NDA documents
 */
function run() {
    try {

        NDA_SETTINGS = getSettings()

        // Get the current spreadsheet
        var activeSpreadsheet = SpreadsheetApp.getActive()

        // Get NDA Doc template from settings
        var templateId = NDA_SETTINGS['template_id']

        // Get spreadsheet timezone
        var timezone = activeSpreadsheet.getSpreadsheetTimeZone()

        // Get the form data to generate the NDAs
        var NDAFormData = this.getNDAFormData()

        // Get the repository folder
        var folder = getRepositoryFolder(activeSpreadsheet.getId())

        // Get the form sheet
        var formSheet = activeSpreadsheet.getSheetByName('Form Responses')

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

            // Save the PDF URL in the NDA spreadsheet
            formSheet.getRange(item.rowIndex, linkColumn).setValue(getFileURL(pdfFile.getId()))

            // Send an email to the NDA recipient
            sendMail(item.data['Email Address'], item.data['Full Name'], pdfFile)

            // Save the status as sent in the NDA spreadsheet
            formSheet.getRange(item.rowIndex, statusColumn).setValue('Sent')

            // Remove the Docs file
            removeFile(driveFile.getId())

        })

    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)

    }
}

/**
 * Returns the Column index number based on name
 * @returns {int}
 */
function getColumnIndexByName(columnName) {

    // Get the current spreadsheet
    var activeSpreadsheet = SpreadsheetApp.getActive()

    // Get the responses sheet
    var sheet = activeSpreadsheet.getSheetByName('Form Responses')

    // Get the data
    var data = sheet.getDataRange().getValues()

    // Get the index of the column
    return data[0].indexOf(columnName) + 1
}

/**
 * Switch on the trigger that will run on form submit to generate the NDA
 */
function toggleTrigger() {
    try {

        // Get the current spreadsheet
        var activeSpreadsheet = SpreadsheetApp.getActive()

        ScriptApp.newTrigger('run')
            .forSpreadsheet(activeSpreadsheet)
            .onFormSubmit()
            .create()

    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)
    }
}

/**
 * Gets the spreadsheet namedRanged
 * @returns {object}
 */
function getNamedRanges() {

    // Get the current spreadsheet
    var activeSpreadsheet = SpreadsheetApp.getActive()

    // Get the spreadsheet named ranges
    var namedRanges = activeSpreadsheet.getNamedRanges()

    // Map each named range with the name and the range information
    return namedRanges.map(function(namedRange) {
        return {
            name: namedRange.getName(),
            range: namedRange.getRange()
        }
    })
}

/**
 * Gets the spreadsheet namedRanged
 * @param {string} name - The name of the range
 * @returns {Range}
 */
function getNamedRange(name) {

    // Get the spreadsheet named ranges
    var namedRanges = getNamedRanges()

    // Get the range by name
    var range = null
    namedRanges.forEach(function(namedRange) {
        if (namedRange.name === name) {
            range = namedRange.range
        }
    })

    return range
}

/**
 * Gets the user settings
 * @returns {object}
 */
function getSettings() {

    // Get the spreadsheet named ranges
    var namedRanges = getNamedRanges()

    // Init the object
    var settings = {}

    // Get each range name and value and mock into the settings object
    namedRanges.forEach(function(namedRange) {
        settings[namedRange.name] = namedRange.range.getValue()
    })

    return settings
}

/**
 * Get the NDA From data formatted from range
 * @returns {object}
 */
function getNDAFormData() {

    // Get the current spreadsheet
    var activeSpreadsheet = SpreadsheetApp.getActive()

    // Get all data
    var data = activeSpreadsheet
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
        if (item.data['NDA Status'] !== 'Sent') {
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
 * @returns {object}
 */
function parseDataToJsonArray(data) {

    var obj = {}
    var result = []
    var headers = data[0]
    var cols = headers.length
    var row = []

    for (var i = 1, l = data.length; i < l; i++) {

        // Get a row to fill the object
        row = data[i]

        // Clear object
        obj = {}

        for (var col = 0; col < cols; col++) {
            // Fill object with new values
            obj[headers[col]] = row[col]
        }

        // Add object in a final result
        result.push(obj)
    }

    return result
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
        Logger.log('Error on show dialog' + e)
    }
}


/**
 * Send an email with an attachment
 * @param {string} emailAddress - Recipient email address
 * @param {string} fullName - Recipient full name
 * @param {File} attachment - DriveApp File instance
 */
function sendMail(emailAddress, fullName, attachment) {

    // Get the email subject
    var emailSubject = NDA_SETTINGS['email_subject']

    // Get the email content
    var emailContent = NDA_SETTINGS['email_body']

    // Convert line breaks to html format
    emailContent = emailContent.replace(/(\r\n|\n|\r)/gm, '<br>')
    var body = '<p>Hello ' + fullName + ',</p><p>' + emailContent + '</p>'

    // Send the email
    MailApp.sendEmail(emailAddress, emailSubject, body, {htmlBody: body, attachments: [attachment]})
}
