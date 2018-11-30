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
 * Merge data into a document from an external source
 * @param {string} docId - The doc file id
 * @param {object} substitutions - The dictionary with keys and values to replace on the document
 */
function mergeTexts(docId, substitutions) {
    var requests = Object
        .keys(substitutions)
        .map(function(key) {
            return {
                replaceAllText: {
                    containsText: {
                        text: '{{' + key + '}}',
                        matchCase: 'true'
                    },
                    replaceText: String(substitutions[key])
                }
            }
        })
    var batchRequests = { requests: requests }
    var response = Docs.Documents.batchUpdate(
        batchRequests,
        'documents/' + docId
    )
}
