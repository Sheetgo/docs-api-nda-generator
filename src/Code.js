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
 * Runs the main script function, gets Youtube data and writes on the spreadsheet
 */
function run() {
    try {

        // TODO: The script

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

