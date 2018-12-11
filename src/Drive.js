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
 * Creates a folder
 * @param {string} name - The folder name
 */
function createFolder(name) {
    return DriveApp.createFolder(name)
}

/**
 * Copies a file, rename it and move it into the solution folder
 * @param {string} id - The file id to be copied
 * @param {string} [fileName] - New file name
 * @param {object} [folder] - Folder object in Google Drive
 */
function copyFile(id, fileName, folder) {

    // Get the file
    var file = DriveApp.getFileById(id)

    // Create the new file
    var newFile = file.makeCopy(fileName || file.getName())

    // If has a folder, move the new file there
    if (folder) {
        moveFile(newFile.getId(), folder)
    }

    // Return the new file
    return newFile
}

/**
 * Moves a file from one folder into another
 * @param {string} id - The file id to be moved
 * @param {object} dest_folder - A folder object in Google Drive
 */
function moveFile(id, dest_folder) {

    // Open the file
    var file = DriveApp.getFileById(id)

    // Add file to the folder
    dest_folder.addFile(file)

    // Get file parent folders
    var parents = file.getParents()

    // Remove the file from the other parents
    while (parents.hasNext()) {
        var folder = parents.next()
        if (folder.getId() !== dest_folder.getId()) {
            folder.removeFile(file)
        }
    }
}

/**
 * Converts a file to PDF and save it on a folder
 * @param {object} id - The file id
 * @param {string} name - The file name
 * @param {object} folder - Drive Folder object
 */
function saveAsPDF(id, name, folder) {

    // Open the file
    var file = DriveApp.getFileById(id)

    // Get the file binary data as pdf
    var fileBlob = file.getBlob().getAs('application/pdf')

    // Add the PDF extension
    fileBlob.setName(file.getName() + ".pdf")

    // Create the file on the folder and return
    return folder.createFile(fileBlob)

}

/**
 * Removes a file - sent it the trash folder
 * @param {string} id - The file id
 */
function removeFile(id) {
    Drive.Files.remove(id)
}

/**
 * Returns the Drive file link
 * @param {string} id - The file id
 * @returns {string}
 */
function getFileURL(id) {
    var file = Drive.Files.get(id)
    return file['alternateLink']
}


/**
 * Returns the repository folder and create if not exists
 * @param {string} spreadsheetId - The spreadsheet id for reference to get the repository folder
 */
function getRepositoryFolder(spreadsheetId) {

    // Get NDA doc name from settings
    var repositoryFolderName = NDA_SETTINGS['repository_name']

    var repositoryFolder = null
    var parentFolderId = DriveApp.getFolderById(spreadsheetId).getParents().next().getId()
    var childrenFolders = DriveApp.getFolderById(parentFolderId).getFolders()

    // Check if nda folder exists
    while (childrenFolders.hasNext()) {
        var childfolder = childrenFolders.next()
        if (childfolder.getName() === repositoryFolderName) {
            repositoryFolder = childfolder
        }
    }

    // Create nda folder if not exists
    if (!repositoryFolder) {
        repositoryFolder = DriveApp.getFolderById(parentFolderId).createFolder(repositoryFolderName)
    }

    return repositoryFolder
}
