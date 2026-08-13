'use strict';

const USER_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be';
const PLATFORM_TABLE_PATH = `${USER_PATH}/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const CLIENTS_TABLE_PATH = `${USER_PATH}/drive/items/01OSXVECQNNRY4S7VCKBF2SOETFSLESSLH/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const RECOMMENDATIONS_TABLE_PATH = `${USER_PATH}/drive/items/01OSXVECRAQXJDB7TBYFGKA5YQJXO3YAOS/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const FEEDBACK_ROWS_PATH = `${USER_PATH}/drive/items/01OSXVECXO7I5R6LKLXJD3VWXORUAF7J37/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/add`;
const SEND_MAIL_PATH = `${USER_PATH}/sendMail`;

function referencePhotoPath(platformRowIndex) {
    return `${USER_PATH}/drive/root:/2. ENTREGA/1. CONTROLAR PLATAFORMA/PG - FOTOS DE REFERÊNCIA/${platformRowIndex}.jpg:/content`;
}

function createMicrosoftGraphAdapter({ graphClient }) {
    function readRows(path) {
        return graphClient.api(path).get();
    }

    return {
        extractRows(response) {
            return response.value;
        },

        extractRowCells(row) {
            return row.values[0];
        },

        readPlatformRows() {
            return readRows(`${PLATFORM_TABLE_PATH}/rows`);
        },

        readClientRows() {
            return readRows(`${CLIENTS_TABLE_PATH}/rows`);
        },

        readRecommendationRows() {
            return readRows(`${RECOMMENDATIONS_TABLE_PATH}/rows`);
        },

        appendPlatformRows(rows) {
            return graphClient.api(`${PLATFORM_TABLE_PATH}/rows/add`).post({ values: rows });
        },

        appendClientRows(rows) {
            return graphClient.api(`${CLIENTS_TABLE_PATH}/rows/add`).post({ values: rows });
        },

        updatePlatformRow(platformRowIndex, cells) {
            return graphClient.api(`${PLATFORM_TABLE_PATH}/rows/itemAt(index=${platformRowIndex})`).update({ values: [cells] });
        },

        updateRecommendationRow(rowIndex, cells) {
            return graphClient.api(`${RECOMMENDATIONS_TABLE_PATH}/rows/itemAt(index=${rowIndex})`).update({ values: [cells] });
        },

        appendRecommendationRow(cells) {
            return graphClient.api(`${RECOMMENDATIONS_TABLE_PATH}/rows/add`).post({ values: [cells] });
        },

        appendFeedbackRow(cells) {
            return graphClient.api(FEEDBACK_ROWS_PATH).post({ values: [cells] });
        },

        uploadReferencePhoto(platformRowIndex, photo) {
            return graphClient.api(referencePhotoPath(platformRowIndex)).put(photo);
        },

        downloadReferencePhoto(platformRowIndex) {
            return graphClient.api(referencePhotoPath(platformRowIndex)).get();
        },

        sendMail(message) {
            return graphClient.api(SEND_MAIL_PATH).post({ message });
        },
    };
}

module.exports = { createMicrosoftGraphAdapter };
