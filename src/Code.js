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
 * Runs the install script function, create a solution folder with this spreadsheet, copy of NDA doc 
 * and another folder for PDFs.
 */
function install() {
    try {
        // Create the Solution folder on users Drive 
        var folder = DriveApp.createFolder('NDA Generator')
        var solutionFolderId = folder.getId()

        // Move host spreadsheet
        var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId())
        this.moveFile(spreadsheetFile, folder)

        // Move linked form
        var form = FormApp.openByUrl(SpreadsheetApp.getActive().getFormUrl())
        this.moveFile(DriveApp.getFileById(form.getId()), folder)

        // Make a copy of NDA document 
        var ndaTemplateFileId = DriveApp.getFileById('1InRRZT3XumIZCfN7k1fWl7dke0S1NPVnUqV9OdhGku4').makeCopy("Non Disclosure Agreement Template", folder).getId()

        // Set NDA doc id on settings tab
        SpreadsheetApp.getActive()
            .getSheetByName('Settings')
            .getRange('c2')
            .setValue(ndaTemplateFileId)

        // Create PDFs folder
        var actualFolderId = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents().next().getId()
        DriveApp.getFolderById(actualFolderId).createFolder('PDF Folder')

    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)

    }
}

/**
 * Runs the main script function, gets Youtube data and writes on the spreadsheet
 */
function run() {
    try {
        // Parse to object
        var data = this.getFormattedData()

        Logger.log(data)


    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)

    }
}

/**
 * Switch on/off the trigger that will run daily at midnight to update the data
 */
function toggleTrigger() {
    try {

        // Get all the user's trigger
        var triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet())

        // If has triggers
        if (triggers && triggers.length > 0) {

            // Disable all the triggers
            for (var i in triggers) {
                if (triggers.hasOwnProperty(i)) {
                    ScriptApp.deleteTrigger(triggers[i])
                }
            }

            // Show the message
            showUiDialog(
                'Automatic update',
                'The trigger was DISABLED'
            )

        } else {

            // Create the trigger
            ScriptApp.newTrigger('run').timeBased().atHour(0).everyDays(1).create()

            // Show the message
            showUiDialog(
                'Automatic update',
                'The trigger was ENABLED, and will run every day at midnight to update the data'
            )
        }

    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)
    }
}

/**
 * Gets the template settings
 * @returns {object}
 */
function getSettings() {
    return getValue('Settings', 'C2')
}

/**
 * Move a file from one folder into another
 * @param {Object} file A file object in Google Drive
 * @param {Object} dest_folder A folder object in Google Drive 
 */
function moveFile(file, dest_folder) {
    dest_folder.addFile(file)
    var parents = file.getParents()
    while (parents.hasNext()) {
        var folder = parents.next()
        if (folder.getId() != dest_folder.getId()) {
            folder.removeFile(file)
        }
    }
}

/**
 * Returns repository folder and create if not exists
 */
function getRepositoryFolder() {
    // Get NDA doc name from settings sheet
    var repositoryFolderName = SpreadsheetApp.getActive()
        .getSheetByName('Settings')
        .getRange('C8')
        .getValue()

    var repositoryFolder = null
    var parentFolderId = DriveApp.getFolderById(SpreadsheetApp.getActive().getId()).getParents().next().getId()
    var childrenFolders = DriveApp.getFolderById(parentFolderId).getFolders()

    // Checks if nda folder exists
    while (childrenFolders.hasNext()) {
        var childfolder = childrenFolders.next()
        if (childfolder.getName() == repositoryFolderName) {
            repositoryFolder = childfolder
        }
    }

    // Create nda folder if not exists
    if (!repositoryFolder) {
        repositoryFolder = DriveApp.getFolderById(parentFolderId).createFolder(repositoryFolderName)
    }

    return repositoryFolder
}

/**
 * Get formatted data from range
 */
function getFormattedData() {
    // Get all data 
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses')
        .getDataRange()
        .getValues()

    var jsonFormatted = this.parseDataToJsonArray(data)

    // Filter sent NDAs
    jsonFormatted = jsonFormatted.filter(function (item) {
        if (item['NDA Status'] != 'sent') {
            delete item['NDA Status']
            delete item['NDA Link PDF']
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
    emailContent = emailContent.replace(/(\r\n|\n|\r)/gm, "<br>")
    var body = '<p>Hello ' + fullName + ',</p><br><p>' + emailContent + '</p>'

    MailApp.sendEmail(emailAddress, emailSubject, body, { htmlBody: body, attachments: [attachment] })
}
