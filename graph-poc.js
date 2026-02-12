require('dotenv').config();
const { ConfidentialClientApplication } = require('@azure/msal-node');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const pdf = require('pdf-parse');

// Azure Configuration
const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};

const cca = new ConfidentialClientApplication(config);

async function getAccessToken() {
    const authResponse = await cca.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default'],
    });
    return authResponse.accessToken;
}

/**
 * Uploads a file to the user's root drive (or a specified drive)
 */
async function uploadToGraph(accessToken, filePath) {
    const fileName = path.basename(filePath);
    const fileStream = fs.readFileSync(filePath);

    console.log(`Uploading ${fileName} to MS Graph...`);

    // For application permissions, we must target a specific user or site.
    // Replace ${process.env.USER_EMAIL} with the email of your developer account.
    const url = `https://graph.microsoft.com/v1.0/users/${process.env.USER_EMAIL}/drive/root:/temp_pocs/${fileName}:/content`;

    try {
        const response = await axios.put(url, fileStream, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }
        });
        return response.data;
    } catch (error) {
        console.error('Upload failed:', error.response ? error.response.data : error.message);
        throw error;
    }
}

/**
 * Converts a drive item to PDF and returns the buffer
 */
async function convertToPdf(accessToken, itemId) {
    console.log(`Converting item ${itemId} to PDF...`);

    // The format=pdf query parameter triggers the conversion
    const url = `https://graph.microsoft.com/v1.0/users/${process.env.USER_EMAIL}/drive/items/${itemId}/content?format=pdf`;

    try {
        const response = await axios.get(url, {
            headers: { 'Authorization': `Bearer ${accessToken}` },
            responseType: 'arraybuffer'
        });
        return response.data;
    } catch (error) {
        console.error('Conversion failed:', error.response ? error.response.data : error.message);
        throw error;
    }
}

/**
 * Searches for text in a PDF buffer and returns page numbers
 */
async function findTextInPdf(pdfBuffer, searchText) {
    console.log(`Searching for "${searchText}" in PDF...`);

    const data = await pdf(pdfBuffer);
    const pages = data.text.split('\f'); // \f is the form feed character (page break)

    const occurrences = [];
    pages.forEach((pageText, index) => {
        if (pageText.toLowerCase().includes(searchText.toLowerCase())) {
            occurrences.push(index + 1);
        }
    });

    return occurrences;
}

/**
 * Cleanup: Deletes the temporary file from Graph
 */
async function deleteFromGraph(accessToken, itemId) {
    console.log(`Cleaning up: Deleting item ${itemId}...`);
    const url = `https://graph.microsoft.com/v1.0/users/${process.env.USER_EMAIL}/drive/items/${itemId}`;
    await axios.delete(url, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });
}

async function main() {
    const docxPath = process.argv[2];
    const searchText = process.argv[3];

    if (!docxPath || !searchText) {
        console.log('Usage: node graph-poc.js <path-to-docx> <search-text>');
        return;
    }

    try {
        const token = await getAccessToken();
        const driveItem = await uploadToGraph(token, docxPath);
        const pdfBuffer = await convertToPdf(token, driveItem.id);

        const pages = await findTextInPdf(pdfBuffer, searchText);
        console.log('\n--- RESULTS ---');
        if (pages.length > 0) {
            console.log(`Found "${searchText}" on pages: ${pages.join(', ')}`);
        } else {
            console.log(`"${searchText}" not found in the document.`);
        }

        await deleteFromGraph(token, driveItem.id);
    } catch (error) {
        console.error('Process failed:', error.message);
    }
}

if (require.main === module) {
    main();
}
