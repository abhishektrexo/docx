const express = require('express');
const multer = require('multer');
const { PublicClientApplication, LogLevel } = require('@azure/msal-node');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const pdfjsLib = require('pdfjs-dist/build/pdf.js');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// --- Configuration ---
const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: 'https://login.microsoftonline.com/common',
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

// Initialize MSAL with Cache from Env Var
const pca = new PublicClientApplication(config);

const getCacheFromEnv = () => {
    if (process.env.MSAL_CACHE) {
        return process.env.MSAL_CACHE;
    }
    return null;
};

// Initialize Cache
const cacheData = getCacheFromEnv();
if (cacheData) {
    pca.getTokenCache().deserialize(cacheData);
    console.log("Cache loaded from environment variable.");
} else {
    console.warn("WARNING: MSAL_CACHE environment variable is not set.");
}


// --- Helper Functions ---

async function getAccessToken() {
    const msalTokenCache = pca.getTokenCache();
    const accounts = await msalTokenCache.getAllAccounts();
    const scopes = ['Files.ReadWrite', 'Files.ReadWrite.All', 'User.Read'];

    if (accounts.length > 0) {
        try {
            const silentRequest = {
                account: accounts[0],
                scopes: scopes,
            };
            const response = await pca.acquireTokenSilent(silentRequest);
            console.log("Silent login successful via cache.");
            return response.accessToken;
        } catch (error) {
            console.error("Silent login failed:", error);
            throw new Error("Silent login failed. Cache might be expired.");
        }
    } else {
        throw new Error("No accounts found in cache. Please authenticate locally and set MSAL_CACHE env var.");
    }
}

async function uploadToMetric(accessToken, fileBuffer, originalName) {
    const fileName = `verzel_${Date.now()}_${originalName}`;
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/temp_pocs/${fileName}:/content`;

    console.log(`Uploading ${fileName}...`);
    const response = await axios.put(url, fileBuffer, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        }
    });
    return response.data;
}

async function convertToPdf(accessToken, itemId) {
    console.log(`Converting item ${itemId} to PDF...`);
    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content?format=pdf`;

    const response = await axios.get(url, {
        headers: { 'Authorization': `Bearer ${accessToken}` },
        responseType: 'arraybuffer'
    });
    return response.data;
}

async function findTextInPdf(pdfBuffer, searchText) {
    console.log(`Searching for "${searchText}"...`);
    const loadingTask = pdfjsLib.getDocument({ data: pdfBuffer });
    const pdfDoc = await loadingTask.promise;
    const numPages = pdfDoc.numPages;

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

async function deleteFile(accessToken, itemId) {
    try {
        await axios.delete(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        console.log(`Deleted temp file ${itemId}.`);
    } catch (e) {
        console.error("Cleanup failed:", e.message);
    }
}

// --- Routes ---

// Serve static frontend
app.use(express.static(path.join(__dirname, '../public')));

app.post('/api/search', upload.single('file'), async (req, res) => {
    const searchText = req.body.search_text;
    const file = req.file;

    if (!file || !searchText) {
        return res.status(400).json({ error: "Missing file or search_text" });
    }

    try {
        const token = await getAccessToken();
        const driveItem = await uploadToMetric(token, file.buffer, file.originalname);
        const pdfBuffer = await convertToPdf(token, driveItem.id);
        const pages = await findTextInPdf(pdfBuffer, searchText);

        // Cleanup asynchronously
        deleteFile(token, driveItem.id);

        res.json({
            success: true,
            pages: pages,
            count: pages.length
        });

    } catch (error) {
        console.error("Error processing request:", error);
        res.status(500).json({
            success: false,
            error: error.message,
            details: error.response?.data
        });
    }
});

// For local testing
if (process.env.NODE_ENV !== 'production') {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
        console.log(`Server running on port ${PORT}`);
    });
}

module.exports = app;
