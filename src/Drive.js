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
 * Copy a file, rename it and move it into the solution folder
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
        moveFile(newFile, folder)
    }

    // Return the new file
    return newFile
}

/**
 * Move a file from one folder into another
 * @param {object} file - A file object in Google Drive
 * @param {object} dest_folder - A folder object in Google Drive
 */
function moveFile(file, dest_folder) {
    dest_folder.addFile(file)
    var parents = file.getParents()
    while (parents.hasNext()) {
        var folder = parents.next()
        if (folder.getId() !== dest_folder.getId()) {
            folder.removeFile(file)
        }
    }
}

/**
 * Save a blob into a Google Drive file
 * @param {object} blob - The binary information of the file
 * @param {string} name - The file name
 * @param {object} folder - Folder object in Google Drive
 */
function saveBlobAsFile(blob, name, folder) {
    // var doc = DocumentApp.openById(id)
    // var docBlob = DocumentApp.openById(id).getAs('application/pdf')
    // docBlob.setName(doc.getName() + ".pdf") // Add the PDF extension
    // var invFolder = SETTINGS.folderId
    var file = DriveApp.getFolderById(folder).createFile(docBlob)
    var url = file.getUrl()
    var id = file.getId()
    return [url, id]
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

