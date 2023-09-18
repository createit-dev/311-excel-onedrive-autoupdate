require('dotenv').config();
const axios = require('axios');
const ExcelJS = require('exceljs');

const EXCEL_FILE_PATH = process.env.EXCEL_FILE_PATH;
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const SITE_ID = process.env.AZURE_SHAREPOINT_SITE_ID;
const FILE_NAME = process.env.AZURE_FILE_NAME;

async function getAccessToken() {
    const tokenEndpoint = "https://login.microsoftonline.com/" + AZURE_TENANT_ID  + "/oauth2/v2.0/token";
    const requestData = {
        client_id: process.env.AZURE_CLIENT_ID,
        scope: "https://graph.microsoft.com/.default",
        client_secret: process.env.AZURE_CLIENT_SECRET,
        grant_type: "client_credentials"
    };

    const response = await axios.post(tokenEndpoint, new URLSearchParams(requestData));

    return response.data.access_token;
}

async function downloadFileFromOneDrive(accessToken) {
    const endpoint = `${GRAPH_BASE_URL}/sites/${SITE_ID}/drive/root:/${FILE_NAME}:?$select=@microsoft.graph.downloadUrl`;
    const headers = {
        'Authorization': `Bearer ${accessToken}`
    };

    try {
        const response = await axios.get(endpoint, { headers: headers});

        const downloadUrl = response.data['@microsoft.graph.downloadUrl'];
        const fileResponse = await axios.get(downloadUrl, { responseType: 'arraybuffer' });
        return fileResponse.data;

    } catch (error) {
        console.error('Error:', error.response.data);
        console.log(endpoint);
        console.log(accessToken);
        return [];
    }
}

async function uploadFileToOneDrive(accessToken, fileData) {
    const headers = {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Prefer': 'bypass-shared-lock'
    };

    const originalEndpoint = `${GRAPH_BASE_URL}/sites/${SITE_ID}/drive/root:/${FILE_NAME}:/content`;

    try {
        await axios.put(originalEndpoint, fileData, { headers: headers });
    } catch (error) {
        if (error.response && error.response.data && error.response.data.error && error.response.data.error.message.includes("locked")) {
            console.log("Current file is locked. Deleting and creating a new one.");

            const deleteEndpoint = `${GRAPH_BASE_URL}/sites/${SITE_ID}/drive/root:/${FILE_NAME}`;

            try {
                await axios.delete(deleteEndpoint, { headers: headers });
            } catch (deleteError) {
                console.error('Error deleting the locked file:', deleteError.response.data);
                return;
            }

            try {
                await axios.put(originalEndpoint, fileData, { headers: headers });
            } catch (uploadError) {
                console.error('Error creating new file:', uploadError.response.data);
                return;
            }
        }
    }
}

function findRowByID(worksheet, candidateID) {
    let targetRow = null;

    worksheet.eachRow((row, rowNumber) => {
        if (row.getCell(1).value === candidateID) {
            targetRow = row;
            return;
        }
    });

    return targetRow;
}

const getFieldByName = (name, candidate) => {
    const field = candidate.fields.find(field => field.name === name);

    if (field && field.values && field.values.length) {
        return field.values.map(val => {
            if (val.hasOwnProperty('text')) {
                return val.text;
            } else if (val.hasOwnProperty('value')) {
                return val.value;
            } else if (val.hasOwnProperty('flag')) {
                return val.flag ? "Yes" : "No";
            }
        }).join(', ');
    }

    return null;
};

function setCellValueAndColor(row, cellNumber, value) {
    const cell = row.getCell(cellNumber);
    cell.value = value;
    cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' }
    };
}

async function update_excel(items) {

    const accessToken = await getAccessToken();
    const fileData = await downloadFileFromOneDrive(accessToken);

    if (!fileData || fileData.length === 0) {
        console.log("Error. Something went wrong in node script...");
        return;
    }

    console.log("OpenDrive data fetched!");

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(fileData);
    const worksheet = workbook.getWorksheet(1);

    for (const elem of items) {
        const row = findRowByID(worksheet, elem.id);

        if (row) {
            console.log("..updating row");

            setCellValueAndColor(row, 1, elem.id);
            setCellValueAndColor(row, 3, elem.name);
            setCellValueAndColor(row, 4, elem.email);
            setCellValueAndColor(row, 5, elem.phone);

            setCellValueAndColor(row, 6, getFieldByName('Position Applied For', elem));
            setCellValueAndColor(row, 7, getFieldByName('Company', elem));
            setCellValueAndColor(row, 8, getFieldByName('City', elem));

        } else {
            console.log("..adding row");

            const newRow = worksheet.addRow([
                elem.id,
                '',
                elem.name,
                elem.email,
                elem.phone,
                getFieldByName('Position Applied For', elem),
                getFieldByName('Company', elem),
                getFieldByName('City', elem),
            ]);
        }
    }

    try {
        const buffer = await workbook.xlsx.writeBuffer();
        await uploadFileToOneDrive(accessToken, buffer);
        console.log(`Successfully saved to OneDrive`);

    } catch (error) {
        console.error(`Error writing to file`);
        console.error(`Error writing to ${EXCEL_FILE_PATH}:`, error);
    }
}

async function getDataFromAPI() {
    try {
        const response = await axios.get('https://jsonplaceholder.typicode.com/users');
        return response.data.map(user => ({
            id: user.id,
            name: user.name,
            email: user.email,
            phone: user.phone,
            fields: [
                { name: 'Position Applied For', values: [{ text: user.id % 2 === 0 ? 'Web Developer' : 'System Analyst' }] },
                { name: 'Company', values: [{ text: user.company.name }] },
                { name: 'City', values: [{ text: user.address.city }] }
            ]
        }));
    } catch (error) {
        console.error('Error fetching data from API:', error);
        return [];
    }
}

(async () => {

    const items = await getDataFromAPI();
    await update_excel(items);

})();
