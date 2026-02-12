require('dotenv').config();
const { PublicClientApplication, LogLevel } = require('@azure/msal-node');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const pdf = require('pdf-parse');
const pdfjsLib = require('pdfjs-dist/build/pdf.js');

// Token Cache Persistence
const cachePath = path.join(__dirname, 'msal-cache.json');

const beforeCacheAccess = async (cacheContext) => {
    if (fs.existsSync(cachePath)) {
        cacheContext.tokenCache.deserialize(fs.readFileSync(cachePath, "utf-8"));
    }
};

const afterCacheAccess = async (cacheContext) => {
    if (cacheContext.cacheHasChanged) {
        fs.writeFileSync(cachePath, cacheContext.tokenCache.serialize());
    }
};

const cachePlugin = {
    beforeCacheAccess,
    afterCacheAccess
};

// Configuration for Personal/Public Account
const config = {
    auth: {
        clientId: process.env.CLIENT_ID, // User must provide this from a "Public" app registration
        authority: 'https://login.microsoftonline.com/common', // "common" supports both Work and Personal accounts
    },
    cache: {
        cachePlugin
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                // console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: LogLevel.Verbose,
        }
    }
};

const pca = new PublicClientApplication(config);

async function getAccessToken() {
    // Try to get account from cache
    const msalTokenCache = pca.getTokenCache();
    const accounts = await msalTokenCache.getAllAccounts();

    const scopes = ['Files.ReadWrite', 'Files.ReadWrite.All', 'User.Read'];

    if (accounts.length > 0) {
        console.log("Found account in cache, attempting silent login...");
        try {
            const silentRequest = {
                account: accounts[0],
                scopes: scopes,
            };
            const response = await pca.acquireTokenSilent(silentRequest);
            console.log("Silent login successful! Logged in as:", response.account.username);
            return response.accessToken;
        } catch (error) {
            console.log("Silent login failed, falling back to device code flow.");
        }
    }

    console.log("Starting Device Code Flow...");

    const deviceCodeRequest = {
        deviceCodeCallback: (response) => {
            if (response && response.message) {
                console.log("\n" + response.message + "\n");
            } else {
                console.log("Device code response received but no message found:", response);
            }
        },
        scopes: scopes,
    };

    try {
        const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
        console.log("Login successful! Logged in as:", response.account.username);
        return response.accessToken;
    } catch (error) {
        console.error("\nAuth Error Details:");
        console.error("- Error Code:", error.errorCode);
        console.error("- Message:", error.errorMessage);
        if (error.subError) console.error("- SubError:", error.subError);

        if (error.errorCode === 'post_request_failed') {
            console.log("\nTIP: Ensure 'Allow public client flows' is set to 'Yes' in Authentication -> Advanced Settings in Azure Portal.");
        }

        throw error;
    }
}

async function uploadToMetric(accessToken, filePath) {
    const fileName = path.basename(filePath);
    const fileStream = fs.readFileSync(filePath);

    console.log(`Uploading ${fileName} to OneDrive...`);

    // Upload to App Root folder (special folder for apps) to avoid clutter, 
    // OR just root if App Folder not supported for the scope. 
    // Using root:/temp_pocs/ is cleaner.
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/temp_pocs/${fileName}:/content`;

    const response = await axios.put(url, fileStream, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        }
    });
    return response.data;
}

async function convertToPdf(accessToken, itemId) {
    console.log(`Converting item ${itemId} to PDF...`);
    // Personal accounts definitely support /me/drive
    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content?format=pdf`;

    try {
        const response = await axios.get(url, {
            headers: { 'Authorization': `Bearer ${accessToken}` },
            responseType: 'arraybuffer'
        });
        return response.data;
    } catch (error) {
        if (error.response && error.response.status === 400) {
            console.error("Conversion failed. Note: Personal OneDrive sometimes has stricter limits or different endpoints for conversion.");
        }
        throw error;
    }
}

async function findTextInPdf(pdfBuffer, searchText) {
    console.log(`Searching for "${searchText}" in PDF (page by page)...`);

    const loadingTask = pdfjsLib.getDocument({ data: pdfBuffer });
    const pdfDoc = await loadingTask.promise;
    const numPages = pdfDoc.numPages;
    console.log(`PDF loaded. Total pages: ${numPages}`);

    const occurrences = [];
    for (let i = 1; i <= numPages; i++) {
        const page = await pdfDoc.getPage(i);
        const content = await page.getTextContent();
        const strings = content.items.map(item => item.str);
        const text = strings.join(' ');

        if (text.toLowerCase().includes(searchText.toLowerCase())) {
            occurrences.push(i);
        }
    }

    return occurrences;
}

async function main() {
    const docxPath = process.argv[2];
    const searchText = process.argv[3];

    if (!docxPath) {
        console.log("Usage: node graph-interactive.js <file.docx> <search_text>");
        return;
    }

    try {
        const token = await getAccessToken();
        const driveItem = await uploadToMetric(token, docxPath);
        const pdfBuffer = await convertToPdf(token, driveItem.id);

        // Save PDF for debugging
        fs.writeFileSync('converted_check.pdf', Buffer.from(pdfBuffer));
        console.log("Saved converted PDF to converted_check.pdf for inspection.");

        const pages = await findTextInPdf(pdfBuffer, searchText || "invention");

        console.log('\n--- RESULTS ---');
        console.log(`Found on pages: ${pages.join(', ')}`);

        // Cleanup
        await axios.delete(`https://graph.microsoft.com/v1.0/me/drive/items/${driveItem.id}`, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

    } catch (err) {
        console.error("Error:", err.message);
        if (err.response) console.error("API Response:", err.response.data);
    }
}

main();
