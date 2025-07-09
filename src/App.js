import React, { useState, useCallback, useMemo, useEffect, useRef } from 'react';
import { useDropzone } from 'react-dropzone';
import { Loader, UploadCloud, Copy, Check, XCircle, Trash2, UserPlus, Search, Users, LogIn, LogOut, Save, PlusCircle, BrainCircuit, ChevronDown, Send } from 'lucide-react';

// --- Google API Configuration ---
// IMPORTANT: Replace these with your own keys from Google Cloud Console.
const API_KEY = "YOUR_GOOGLE_API_KEY"; // Your Google Cloud API Key
const CLIENT_ID = "YOUR_GOOGLE_CLIENT_ID"; // Your Google Cloud OAuth 2.0 Client ID
const DISCOVERY_DOCS = [
    "https://sheets.googleapis.com/$discovery/rest?version=v4",
    "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"
];
const SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file";
const SPREADSHEET_NAME = "LabValueExtractor_Data";

// --- GoHighLevel/Make.com Webhook Configuration ---
// IMPORTANT: Replace this with the webhook URL you get from Make.com in Part 2.
const MAKE_WEBHOOK_URL = "https://hook.us2.make.com/tqphg08ye5enlrwuj1kkjyo9iea5dtej";

// --- Reference Ranges ---
const REFERENCE_RANGES = {
    'LH': { low: 1.7, high: 8.6, units: 'mIU/mL' }, 'FSH': { low: 1.5, high: 12.4, units: 'mIU/mL' }, 'Estradiol (E2)': { low: 10, high: 40, units: 'pg/mL' }, 'Progesterone': { low: 0.1, high: 0.8, units: 'ng/mL' }, 'Prolactin': { low: 4, high: 23, units: 'ng/mL' }, 'Testosterone': { low: 280, high: 1100, units: 'ng/dL' }, 'DHEA-S': { low: 102.6, high: 416.3, units: 'ug/dL' }, 'AMH': { low: 1.5, high: 9.5, units: 'ng/mL' }, 'TSH': { low: 0.45, high: 4.5, units: 'uIU/mL' }, 'Free T3': { low: 2.0, high: 4.4, units: 'pg/mL' }, 'Free T4': { low: 0.82, high: 1.77, units: 'ng/dL' }, 'Total T3': { low: 71, high: 180, units: 'ng/dL' }, 'Total T4': { low: 4.5, high: 12.0, units: 'ug/dL' }, 'Vitamin D': { low: 30, high: 100, units: 'ng/mL' }, 'B12': { low: 232, high: 1245, units: 'pg/mL' }, 'Ferritin': { low: 30, high: 400, units: 'ng/mL' }, 'Iron': { low: 65, high: 175, units: 'mcg/dL' }, 'Iron Saturation': { low: 20, high: 50, units: '%' }, 'TIBC': { low: 250, high: 450, units: 'mcg/dL' }, 'Fasting Glucose': { low: 70, high: 99, units: 'mg/dL' }, 'Fasting Insulin': { low: 2.6, high: 24.9, units: 'uIU/mL' }, 'HbA1c': { low: 4.0, high: 5.6, units: '%' }, 'Cholesterol Total': { low: 100, high: 199, units: 'mg/dL' }, 'HDL': { low: 40, high: 60, units: 'mg/dL' }, 'LDL': { low: 0, high: 99, units: 'mg/dL' }, 'Triglycerides': { low: 0, high: 149, units: 'mg/dL' }, 'Hemoglobin': { low: 13.2, high: 16.6, units: 'g/dL' }, 'Hematocrit': { low: 38.3, high: 48.6, units: '%' }, 'WBC': { low: 4.5, high: 11.0, units: 'K/uL' }, 'RBC': { low: 4.5, high: 5.9, units: 'M/uL' }, 'Platelets': { low: 150, high: 450, units: 'K/uL' }, 'MCV': { low: 80, high: 100, units: 'fL' }, 'MCH': { low: 27, high: 34, units: 'pg' }, 'MCHC': { low: 32, high: 36, units: 'g/dL' }, 'Neutrophils': { low: 40, high: 75, units: '%' }, 'Lymphocytes': { low: 20, high: 45, units: '%' }, 'Monocytes': { low: 2, high: 10, units: '%' }, 'Eosinophils': { low: 0, high: 6, units: '%' }, 'Basophils': { low: 0, high: 1, units: '%' }, 'Sodium': { low: 135, high: 145, units: 'mmol/L' }, 'Potassium': { low: 3.5, high: 5.2, units: 'mmol/L' }, 'Alkaline Phosphatase': { low: 44, high: 147, units: 'IU/L' },
};

const toBase64 = file => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = error => reject(error);
});

export default function App() {
    // --- State Management ---
    const [gapiClient, setGapiClient] = useState(null);
    const [tokenClient, setTokenClient] = useState(null);
    const [accessToken, setAccessToken] = useState(null);
    const [spreadsheetId, setSpreadsheetId] = useState(null);
    
    const [clientList, setClientList] = useState([]);
    const [activeClientData, setActiveClientData] = useState(null);
    const [selectedClientEmail, setSelectedClientEmail] = useState(null);
    const [newClientEmail, setNewClientEmail] = useState('');
    const [searchQuery, setSearchQuery] = useState('');
    const [isDirty, setIsDirty] = useState(false);
    
    const [isLoading, setIsLoading] = useState(false);
    const [loadingMessage, setLoadingMessage] = useState('');
    const [error, setError] = useState(null);

    const [isAnalysisModalOpen, setIsAnalysisModalOpen] = useState(false);
    const [analysisResult, setAnalysisResult] = useState('');
    const [isAnalyzing, setIsAnalyzing] = useState(false);
    const [isAnalysisDropdownOpen, setIsAnalysisDropdownOpen] = useState(false);
    const [isSendingToGHL, setIsSendingToGHL] = useState(false);
    const [ghlStatus, setGhlStatus] = useState('');
    
    const processingRef = useRef(false);

    // --- Google API Initialization ---
    useEffect(() => {
        if (API_KEY === "YOUR_GOOGLE_API_KEY" || CLIENT_ID === "YOUR_GOOGLE_CLIENT_ID") {
            setError("Configuration Error: Google API Key and Client ID are missing. Please add them to the code to enable Google Sign-In.");
            return;
        }

        const gapiScript = document.createElement('script');
        gapiScript.src = 'https://apis.google.com/js/api.js';
        gapiScript.onload = () => window.gapi.load('client', initializeGapiClient);
        document.body.appendChild(gapiScript);

        const gisScript = document.createElement('script');
        gisScript.src = 'https://accounts.google.com/gsi/client';
        gisScript.onload = initializeGisClient;
        document.body.appendChild(gisScript);

        if (window.pdfjsLib) {
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.102/pdf.worker.min.js`;
        }
    }, []);

    const initializeGapiClient = async () => {
        try {
            await window.gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: DISCOVERY_DOCS,
            });
            setGapiClient(window.gapi.client);
        } catch (e) {
            setError("Failed to initialize Google API Client. Check API Key and enabled APIs.");
            console.error("GAPI Client Init Error:", e);
        }
    };

    const initializeGisClient = () => {
        try {
            const client = window.google.accounts.oauth2.initTokenClient({
                client_id: CLIENT_ID,
                scope: SCOPES,
                callback: (tokenResponse) => {
                    if (tokenResponse && tokenResponse.access_token) {
                        setAccessToken(tokenResponse.access_token);
                    } else {
                        console.error("Token response was empty or invalid.");
                    }
                },
            });
            setTokenClient(client);
        } catch (e) {
            setError("Failed to initialize Google Identity Services. Check Client ID.");
            console.error("GIS Client Init Error:", e);
        }
    };
    
    useEffect(() => {
        if (accessToken && gapiClient) {
            gapiClient.setToken({ access_token: accessToken });
            loadClientList();
        }
    }, [accessToken, gapiClient]);
    
    const handleAuthClick = () => {
        if (tokenClient) {
            tokenClient.requestAccessToken({ prompt: 'consent' });
        }
    };

    const handleSignoutClick = () => {
        if (accessToken) {
            window.google.accounts.oauth2.revoke(accessToken, () => {
                setAccessToken(null);
                setClientList([]);
                setActiveClientData(null);
                setSelectedClientEmail(null);
            });
        }
    };

    // Refactored core data loading logic
    const loadDataForClient = async (clientEmail) => {
        if (!gapiClient || !accessToken) return;
        setActiveClientData(null);
        setIsLoading(true);
        setLoadingMessage(`Loading data for ${clientEmail}...`);
        setError(null);
        try {
            const response = await gapiClient.sheets.spreadsheets.values.get({
                spreadsheetId: spreadsheetId,
                range: `'${clientEmail}'!A:Z`,
            });
            const values = response.result.values || [];
            setActiveClientData(values);
        } catch (err) {
            console.error("Error loading sheet data:", err);
            setError(`Failed to load data for ${clientEmail}.`);
        } finally {
            setIsLoading(false);
            setLoadingMessage('');
        }
    };

    const handleClientSelection = async (clientEmail) => {
        if (processingRef.current) {
            alert("Please wait for the current operation to complete.");
            return;
        }
        if (isDirty) {
            if (!window.confirm("You have unsaved changes. Are you sure you want to switch clients? Your changes will be lost.")) {
                return;
            }
        }
        setSelectedClientEmail(clientEmail);
        setNewClientEmail(clientEmail);
        setIsDirty(false);
        await loadDataForClient(clientEmail);
    };

    const loadClientList = async () => {
        if (!gapiClient || !accessToken) return;
        setIsLoading(true);
        setLoadingMessage('Finding your data sheet...');
        try {
            const response = await gapiClient.drive.files.list({
                q: `mimeType='application/vnd.google-apps.spreadsheet' and name='${SPREADSHEET_NAME}' and trashed=false`,
                fields: 'files(id, name)',
            });
            let SPREADSHEET_ID = response.result.files.length > 0 ? response.result.files[0].id : null;

            if (!SPREADSHEET_ID) {
                setLoadingMessage('Creating new data sheet in your Google Drive...');
                const createResponse = await gapiClient.sheets.spreadsheets.create({
                    properties: { title: SPREADSHEET_NAME },
                    sheets: [{ properties: { title: 'Welcome' } }]
                });
                SPREADSHEET_ID = createResponse.result.spreadsheetId;
            }
            setSpreadsheetId(SPREADSHEET_ID);
            
            setLoadingMessage('Loading client list...');
            const sheetDataResponse = await gapiClient.sheets.spreadsheets.get({
                spreadsheetId: SPREADSHEET_ID,
                fields: 'sheets.properties.title',
            });
            const sheets = sheetDataResponse.result.sheets.map(s => s.properties.title).filter(t => t !== 'Welcome');
            setClientList(sheets);

        } catch (err) {
            setError('Could not load data from Google Sheets. Ensure permissions are granted and APIs are enabled.');
            console.error("Error loading data:", err);
        } finally {
            setIsLoading(false);
            setLoadingMessage('');
        }
    };
    
    const onDrop = useCallback(async (acceptedFiles) => {
        if (processingRef.current) return;

        const file = acceptedFiles[0];
        if (!file || !spreadsheetId) return;
        if (!newClientEmail.trim()) {
            setError('Please enter a client email.');
            return;
        }

        processingRef.current = true;
        setIsLoading(true);
        setError(null);

        try {
            let fileParts = [];
            
            if (file.type === 'application/pdf') {
                setLoadingMessage('Processing PDF...');
                const arrayBuffer = await file.arrayBuffer();
                const pdf = await window.pdfjsLib.getDocument(arrayBuffer).promise;
                const numPages = pdf.numPages;

                for (let i = 1; i <= numPages; i++) {
                    setLoadingMessage(`Processing Page ${i} of ${numPages}...`);
                    const page = await pdf.getPage(i);
                    const viewport = page.getViewport({ scale: 2.0 });
                    const canvas = document.createElement('canvas');
                    const context = canvas.getContext('2d');
                    canvas.height = viewport.height;
                    canvas.width = viewport.width;
                    await page.render({ canvasContext: context, viewport: viewport }).promise;
                    const base64Data = canvas.toDataURL('image/jpeg').split(',')[1];
                    fileParts.push({ inlineData: { mimeType: 'image/jpeg', data: base64Data } });
                }
            } else {
                setLoadingMessage('Processing Image...');
                const base64Data = await toBase64(file);
                fileParts.push({ inlineData: { mimeType: file.type, data: base64Data } });
            }

            setLoadingMessage('Analyzing document with AI...');
            const prompt = `
                You are an expert lab value extraction tool. Analyze the provided document images. The images represent pages of a single report.
                Extract the report date (collection date) and any of the lab values from the requested list below. Consolidate results from all pages into a single JSON object.
                CRITICAL INSTRUCTION: Only extract the markers explicitly listed. If a marker like 'Cortisol' is present but not in the requested list, you MUST ignore it completely.
                Be flexible with names; for example, "TESTOSTERONE, TOTAL, MS" should be "Testosterone". 
                Pay close attention to layouts where the marker, units, and value might be on separate lines. For example, if you see "ESTRADIOL" on one line, "pg/mL" on the line below, and the value "65" to the right, you must correctly associate these as a single Estradiol result.
                Return a single JSON object with a top-level key "reportDate" and other keys for categories.
                Requested Markers: Hormones, Thyroid Panel, Vitamins & Nutrients, Glucose / Insulin / Metabolic, CBC Panel, Electrolytes / Other.
            `;
            const payload = { contents: [{ role: "user", parts: [{ text: prompt }, ...fileParts] }] };
            const geminiApiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${API_KEY}`;

            const geminiResponse = await fetch(geminiApiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
            if (!geminiResponse.ok) throw new Error(`Gemini API request failed. Status: ${geminiResponse.status}`);
            
            const result = await geminiResponse.json();
            if (!result.candidates?.[0]?.content?.parts?.[0]) throw new Error("AI could not extract data from the document.");

            let text = result.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
            const parsedJson = JSON.parse(text);

            setLoadingMessage('Saving data to Google Sheets...');
            await writeToSheet(newClientEmail.trim(), file.name, parsedJson);
            
            await loadClientList();
            setSelectedClientEmail(newClientEmail.trim());
            await loadDataForClient(newClientEmail.trim());
            setNewClientEmail(newClientEmail.trim());

        } catch (err) {
            console.error("Processing Error:", err);
            setError(err.message || "An unexpected error occurred during processing.");
        } finally {
            setIsLoading(false);
            setLoadingMessage('');
            processingRef.current = false;
        }
    }, [newClientEmail, spreadsheetId, gapiClient, clientList, accessToken]);

    const writeToSheet = async (clientEmail, fileName, extractedData) => {
        const sheetExists = clientList.includes(clientEmail);
        let existingData = [];

        if (sheetExists) {
            const response = await gapiClient.sheets.spreadsheets.values.get({ spreadsheetId, range: `'${clientEmail}'!A:Z` });
            existingData = response.result.values || [];
        }
        
        const reportDate = extractedData.reportDate || new Date().toLocaleDateString();
        const newColumnHeader = `${reportDate}\n${fileName}`;
        let updatedData = [];

        if (existingData.length === 0) {
            updatedData.push(['Marker', 'Reference Range', newColumnHeader]);
            Object.keys(REFERENCE_RANGES).forEach(marker => {
                const range = REFERENCE_RANGES[marker];
                const rangeString = `${range.low} - ${range.high} ${range.units}`;
                const extractedValue = findValue(extractedData, marker);
                updatedData.push([marker, rangeString, extractedValue]);
            });
        } else {
            updatedData = existingData;
            updatedData[0].push(newColumnHeader);
            for (let i = 1; i < updatedData.length; i++) {
                const marker = updatedData[i][0];
                const extractedValue = findValue(extractedData, marker);
                while (updatedData[i].length < updatedData[0].length -1) {
                    updatedData[i].push('—');
                }
                updatedData[i].push(extractedValue);
            }
        }
        
        if (!sheetExists) {
            await gapiClient.sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: { requests: [{ addSheet: { properties: { title: clientEmail } } }] }
            });
        }
        
        await gapiClient.sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `'${clientEmail}'!A1`,
            valueInputOption: 'USER_ENTERED',
            resource: { values: updatedData },
        });
    };
    
    const handleAddManualColumn = () => {
        if (processingRef.current) return;
        const date = prompt("Enter a date or note for the new column (e.g., 'Manual Entry 7/9/2025'):");
        if (!date || !activeClientData) return;

        const newData = activeClientData.map((row, index) => (index === 0) ? [...row, date] : [...row, '—']);
        setActiveClientData(newData);
        setIsDirty(true);
    };

    const handleDeleteColumn = async (colIndex) => {
        if (processingRef.current) return;
        if (!activeClientData || colIndex < 2) return;
        if (window.confirm("Are you sure you want to delete this entire column? This action will be saved immediately.")) {
            const newData = activeClientData.map(row => {
                const newRow = [...row];
                newRow.splice(colIndex, 1);
                return newRow;
            });
            await handleSaveManualChanges(newData);
            setActiveClientData(newData);
        }
    };
    
    const handleSaveManualChanges = async (dataToSave) => {
        if (processingRef.current) return;
        const data = dataToSave || activeClientData;
        if (!selectedClientEmail || !data) return;
        
        processingRef.current = true;
        setIsLoading(true);
        setLoadingMessage('Saving changes to Google Sheets...');
        setError(null);
        try {
            await gapiClient.sheets.spreadsheets.values.clear({ spreadsheetId, range: `'${selectedClientEmail}'` });
            await gapiClient.sheets.spreadsheets.values.update({
                spreadsheetId,
                range: `'${selectedClientEmail}'!A1`,
                valueInputOption: 'USER_ENTERED',
                resource: { values: data },
            });
            setIsDirty(false);
        } catch (err) {
            console.error("Error saving data:", err);
            setError("Failed to save changes. Please try again.");
        } finally {
            setIsLoading(false);
            setLoadingMessage('');
            processingRef.current = false;
        }
    };

    const handleCellChange = (rowIndex, colIndex, value) => {
        const updatedData = activeClientData.map((row, rIdx) => (rIdx === rowIndex) ? row.map((cell, cIdx) => (cIdx === colIndex ? value : cell)) : row);
        setActiveClientData(updatedData);
        setIsDirty(true);
    };

    const handleHeaderChange = (colIndex, value) => {
        const updatedData = activeClientData.map((row, rIdx) => (rIdx === 0) ? row.map((cell, cIdx) => (cIdx === colIndex ? value : cell)) : row);
        setActiveClientData(updatedData);
        setIsDirty(true);
    };

    const handleAnalyzeLabs = async (analysisType) => {
        if (!activeClientData || (activeClientData[0] || []).length <= 2 || processingRef.current) return;
        
        setIsAnalyzing(true);
        setAnalysisResult('');
        setGhlStatus('');
        setIsAnalysisModalOpen(true);
        setError(null);
        setIsAnalysisDropdownOpen(false);

        try {
            const headers = activeClientData[0];
            const lastColIndex = headers.length - 1;
            
            let labDataString = "Here are the latest lab results:\n";
            for (let i = 1; i < activeClientData.length; i++) {
                const row = activeClientData[i];
                const marker = row[0];
                const refRange = row[1];
                const value = row[lastColIndex];
                if (marker && value && value !== '—') {
                    labDataString += `- ${marker}: ${value} (Reference: ${refRange})\n`;
                }
            }
            
            let analysisFocus = '';
            switch (analysisType) {
                case 'Fertility':
                    analysisFocus = 'Focus on key fertility markers like FSH, LH, Estradiol, AMH, TSH, and Prolactin. Explain what each key marker generally indicates in the context of fertility.';
                    break;
                case 'Thyroid':
                    analysisFocus = 'Focus on key thyroid markers like TSH, Free T3, Free T4, Total T3, and Total T4. Explain what each marker generally indicates for thyroid health.';
                    break;
                case 'Metabolic':
                    analysisFocus = 'Focus on key metabolic markers like Fasting Glucose, Fasting Insulin, HbA1c, and the lipid panel (Cholesterol, HDL, LDL, Triglycerides). Explain what each marker generally indicates for metabolic health.';
                    break;
                default:
                    analysisFocus = 'Provide a general overview of the lab results.';
            }

            const prompt = `You are a helpful assistant providing a concise, bullet-pointed analysis of lab results. Analyze the following lab results. ${analysisFocus} Do not provide any medical advice, diagnosis, or treatment recommendations. Here is the data:\n${labDataString}`;

            const payload = { contents: [{ role: "user", parts: [{ text: prompt }] }] };
            const geminiApiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${API_KEY}`;

            const geminiResponse = await fetch(geminiApiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
            if (!geminiResponse.ok) throw new Error(`Gemini API request failed. Status: ${geminiResponse.status}`);
            
            const result = await geminiResponse.json();
            if (!result.candidates?.[0]?.content?.parts?.[0]) throw new Error("The AI could not generate an analysis.");
            
            setAnalysisResult(result.candidates[0].content.parts[0].text);

        } catch (err) {
            console.error("Analysis Error:", err);
            setError(err.message || "An unexpected error occurred during analysis.");
            setAnalysisResult("Sorry, an error occurred while generating the analysis. Please try again.");
        } finally {
            setIsAnalyzing(false);
        }
    };
    
    const handleSendToGHL = async () => {
        if (!selectedClientEmail || !analysisResult || MAKE_WEBHOOK_URL === "YOUR_MAKE_WEBHOOK_URL_HERE") {
            setGhlStatus("Error: Make.com Webhook URL is not configured.");
            return;
        }
        setIsSendingToGHL(true);
        setGhlStatus('Sending...');
        try {
            const response = await fetch(MAKE_WEBHOOK_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: selectedClientEmail, note: analysisResult })
            });

            if (!response.ok) {
                throw new Error(`Webhook request failed with status ${response.status}`);
            }
            
            setGhlStatus('Successfully sent to GoHighLevel!');

        } catch (err) {
            console.error("GHL Send Error:", err);
            setGhlStatus(`Error: ${err.message}`);
        } finally {
            setIsSendingToGHL(false);
        }
    };

    const handleDeleteClient = async () => {
        if (!selectedClientEmail || processingRef.current) return;
        if (window.confirm(`Are you sure you want to permanently delete all data for ${selectedClientEmail}? This action cannot be undone.`)) {
            processingRef.current = true;
            setIsLoading(true);
            setLoadingMessage(`Deleting ${selectedClientEmail}...`);
            setError(null);
            try {
                const spreadsheet = await gapiClient.sheets.spreadsheets.get({ spreadsheetId });
                const sheet = spreadsheet.result.sheets.find(s => s.properties.title === selectedClientEmail);
                if (sheet) {
                    const sheetId = sheet.properties.sheetId;
                    await gapiClient.sheets.spreadsheets.batchUpdate({
                        spreadsheetId,
                        resource: { requests: [{ deleteSheet: { sheetId } }] }
                    });
                }
                setActiveClientData(null);
                setSelectedClientEmail(null);
                setNewClientEmail('');
                await loadClientList();
            } catch (err) {
                 console.error("Error deleting client:", err);
                 setError("Failed to delete client. Please try again.");
            } finally {
                setIsLoading(false);
                setLoadingMessage('');
                processingRef.current = false;
            }
        }
    };
    
    const findValue = (extractedData, markerName) => {
        for (const category in extractedData) {
            if (Array.isArray(extractedData[category])) {
                const found = extractedData[category].find(item => item.marker === markerName || item.marker.includes(markerName));
                if (found) return `${found.value} ${found.units || ''}`.trim();
            }
        }
        return '—';
    };
    
    const getRangeStatus = (resultValue, markerName) => {
        const range = REFERENCE_RANGES[markerName];
        if (!range || !resultValue || resultValue === '—') return 'normal';
        const numericValue = parseFloat(resultValue);
        if (isNaN(numericValue)) return 'normal';
        if (markerName === 'HDL') return numericValue < range.low ? 'low' : 'normal';
        if (numericValue < range.low) return 'low';
        if (numericValue > range.high) return 'high';
        return 'normal';
    };

    const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop, noClick: false, multiple: false });

    const filteredClients = useMemo(() => clientList.filter(name => name.toLowerCase().includes(searchQuery.toLowerCase())).sort(), [clientList, searchQuery]);
    
    const renderResultsTable = () => {
        if (!activeClientData) {
             return (
                <div className="text-center p-10 bg-gray-50 rounded-lg h-full flex flex-col justify-center items-center">
                    <Users className="mx-auto h-16 w-16 text-gray-400" />
                    <h3 className="mt-4 text-lg font-medium text-gray-900">{selectedClientEmail ? `Loading...` : "Select a Client"}</h3>
                    <p className="mt-1 text-sm text-gray-500">Choose a client from the list to view their results.</p>
                </div>
            );
        }
        
        const headers = activeClientData[0] || [];
        const rows = activeClientData.slice(1);
        const hasDataColumns = headers.length > 2;

        return (
             <div className="p-4 sm:p-6">
                <div className="flex justify-between items-center mb-4 flex-wrap gap-4">
                    <h2 className="text-2xl font-bold text-gray-800">Results for: <span className="text-blue-600">{selectedClientEmail}</span></h2>
                    <div className="flex gap-2 flex-wrap">
                        <div className="relative">
                            <button 
                                onClick={() => setIsAnalysisDropdownOpen(!isAnalysisDropdownOpen)} 
                                className={`flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-md transition-all ${!hasDataColumns ? 'cursor-not-allowed bg-purple-300' : 'hover:bg-purple-700'}`}
                                disabled={!hasDataColumns}
                                title={!hasDataColumns ? "No data to analyze" : "Analyze lab results"}
                            >
                                <BrainCircuit size={20} /> Analyze Labs <ChevronDown size={20} />
                            </button>
                            {isAnalysisDropdownOpen && hasDataColumns && (
                                <div className="absolute right-0 mt-2 w-48 bg-white rounded-md shadow-lg z-10 border">
                                    <button onClick={() => handleAnalyzeLabs('Fertility')} className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100">Fertility Health</button>
                                    <button onClick={() => handleAnalyzeLabs('Thyroid')} className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100">Thyroid Health</button>
                                    <button onClick={() => handleAnalyzeLabs('Metabolic')} className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100">Metabolic Health</button>
                                </div>
                            )}
                        </div>
                        <button onClick={handleAddManualColumn} className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-all">
                            <PlusCircle size={20} /> Add Manual Column
                        </button>
                        {isDirty && (
                            <button onClick={() => handleSaveManualChanges()} className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-all animate-pulse">
                                <Save size={20} /> Save Changes
                            </button>
                        )}
                    </div>
                </div>
                <div className="overflow-x-auto bg-white rounded-xl shadow-md border border-gray-200">
                    <table className="w-full text-sm text-left text-gray-600">
                        <thead className="text-xs text-gray-700 uppercase bg-gray-100">
                            <tr>{headers.map((h, i) => {
                                if (i < 2) {
                                    return <th key={i} scope="col" className="px-6 py-3 whitespace-pre-wrap relative group">{h}</th>;
                                }
                                const [date, ...fileNameParts] = h.split('\n');
                                const fileName = fileNameParts.join('\n');
                                
                                return (
                                    <th key={i} scope="col" className="px-1 py-1 whitespace-pre-wrap relative group">
                                        <div className="flex flex-col items-center justify-center">
                                            <input type="text" value={date} onChange={(e) => handleHeaderChange(i, `${e.target.value}\n${fileName}`)} className="w-full p-1 bg-transparent border border-transparent focus:bg-white focus:border-blue-500 rounded-md outline-none text-center font-semibold text-gray-800" aria-label="Report Date" />
                                            <input type="text" value={fileName} onChange={(e) => handleHeaderChange(i, `${date}\n${e.target.value}`)} className="w-full p-1 bg-transparent border border-transparent focus:bg-white focus:border-blue-500 rounded-md outline-none text-center text-xs text-gray-500 truncate" title={fileName} aria-label="File Name" />
                                        </div>
                                        <button onClick={() => handleDeleteColumn(i)} className="absolute top-1 right-1 p-1 bg-red-100 text-red-600 rounded-full opacity-0 group-hover:opacity-100 transition-opacity">
                                            <Trash2 size={14} />
                                        </button>
                                    </th>
                                );
                            })}</tr>
                        </thead>
                        <tbody>
                            {rows.map((row, rowIndex) => (
                                <tr key={rowIndex} className="bg-white border-b hover:bg-blue-50/50">
                                    {row.map((cell, cellIndex) => {
                                        if (cellIndex < 2) return <td key={cellIndex} className="px-6 py-4 font-medium">{cell}</td>;
                                        const markerName = row[0];
                                        const status = getRangeStatus(cell, markerName);
                                        let statusClass = '';
                                        if (status === 'low') statusClass = 'bg-blue-100 text-blue-800 font-bold';
                                        if (status === 'high') statusClass = 'bg-red-100 text-red-800 font-bold';
                                        return (
                                            <td key={cellIndex} className={`px-1 py-1 ${statusClass}`}>
                                                <input type="text" value={cell} onChange={(e) => handleCellChange(rowIndex + 1, cellIndex, e.target.value)} className="w-full p-2 bg-transparent border border-transparent focus:bg-white focus:border-blue-500 rounded-md outline-none text-center" />
                                            </td>
                                        );
                                    })}
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        );
    };
    
    if (error && error.startsWith("Configuration Error:")) {
        return (
            <div className="flex items-center justify-center h-screen bg-red-50">
                <div className="text-center p-8 bg-white rounded-lg shadow-xl max-w-lg mx-4">
                    <XCircle className="mx-auto h-12 w-12 text-red-500" />
                    <h2 className="mt-4 text-2xl font-bold text-gray-800">Configuration Required</h2>
                    <p className="mt-2 text-gray-600">{error}</p>
                </div>
            </div>
        )
    }

    if (!gapiClient || !tokenClient) {
        return <div className="flex items-center justify-center h-screen"><Loader className="animate-spin h-10 w-10" /></div>;
    }

    if (!accessToken) {
        return (
            <div className="flex flex-col items-center justify-center h-screen bg-gray-100">
                <h1 className="text-4xl font-bold text-gray-800 mb-2">Lab Value Extractor</h1>
                <p className="text-lg text-gray-600 mb-8">Sign in with your Google Account to continue.</p>
                <button onClick={handleAuthClick} className="flex items-center gap-3 px-6 py-3 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 transition-all">
                    <LogIn/>
                    Sign in with Google
                </button>
                 {error && (<div className="mt-4 p-3 bg-red-100 text-red-700 rounded-md text-sm max-w-md text-center">{error}</div>)}
            </div>
        );
    }

    return (
        <div className="bg-gray-100 min-h-screen font-sans flex flex-col">
            <header className="bg-white shadow-sm p-4 border-b border-gray-200">
                <div className="container mx-auto flex justify-between items-center">
                    <h1 className="text-2xl font-bold text-gray-800">Lab Value Extractor</h1>
                     <button onClick={handleSignoutClick} className="flex items-center px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700">
                        <LogOut className="h-5 w-5 mr-2" /> Sign Out
                    </button>
                </div>
            </header>
            <div className="flex-grow container mx-auto flex flex-col md:flex-row gap-6 p-4 sm:p-6">
                <aside className="w-full md:w-1/3 lg:w-1/4 flex flex-col gap-4">
                     <div className="bg-white p-4 rounded-xl shadow-lg border border-gray-200">
                        <h2 className="text-lg font-semibold text-gray-700 mb-3">Add New Report</h2>
                        <div className="relative mb-3">
                            <UserPlus className="absolute left-3 top-1/2 -translate-y-1/2 h-5 w-5 text-gray-400" />
                            <input type="email" value={newClientEmail} onChange={(e) => setNewClientEmail(e.target.value)} placeholder="Enter Client Email" className="w-full pl-10 pr-4 py-2 border rounded-md"/>
                        </div>
                        <div {...getRootProps()} className={`w-full p-6 border-2 border-dashed rounded-lg text-center cursor-pointer ${isDragActive ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-blue-400'}`}>
                            <input {...getInputProps()} />
                            <UploadCloud className="mx-auto h-10 w-10 text-gray-400" />
                            <p className="mt-2 text-sm font-semibold text-gray-700">{isDragActive ? "Drop file..." : "Drag & drop file or click"}</p>
                        </div>
                         {isLoading && (<div className="flex items-center justify-center mt-3 text-sm text-blue-600"><Loader className="h-4 w-4 mr-2 animate-spin" /><span>{loadingMessage}</span></div>)}
                         {error && (<div className="flex items-start mt-3 p-2 bg-red-50 text-red-700 rounded-md text-sm"><XCircle className="h-4 w-4 mr-2 mt-0.5"/><span>{error}</span></div>)}
                    </div>
                     <div className="bg-white p-4 rounded-xl shadow-lg border flex-grow flex flex-col">
                         <h2 className="text-lg font-semibold text-gray-700 mb-3">Clients</h2>
                         <div className="relative mb-3">
                            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-5 w-5 text-gray-400" />
                            <input type="text" value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="Search by email..." className="w-full pl-10 pr-4 py-2 border rounded-md"/>
                        </div>
                        <div className="flex-grow overflow-y-auto">
                            {filteredClients.length > 0 ? (
                                <ul className="space-y-1">
                                    {filteredClients.map(email => (
                                        <li key={email} className="flex items-center justify-between group">
                                            <button onClick={() => handleClientSelection(email)} className={`w-full text-left px-3 py-2 rounded-md transition-colors duration-150 ${selectedClientEmail === email ? 'bg-blue-600 text-white font-semibold' : 'hover:bg-gray-100'}`}>
                                                {email}
                                            </button>
                                            {selectedClientEmail === email && (
                                                <button onClick={handleDeleteClient} className="p-2 text-red-500 hover:bg-red-100 rounded-md opacity-0 group-hover:opacity-100 transition-opacity">
                                                    <Trash2 size={16} />
                                                </button>
                                            )}
                                        </li>
                                    ))}
                                </ul>
                            ) : (<p className="text-sm text-gray-500 text-center py-4">No clients found.</p>)}
                        </div>
                    </div>
                </aside>
                <main className="w-full md:w-2/3 lg:w-3/4 bg-white rounded-xl shadow-lg border">
                    {isLoading && !activeClientData ? (
                         <div className="flex items-center justify-center h-full"><Loader className="animate-spin h-10 w-10 text-blue-500" /></div>
                    ) : renderResultsTable()}
                </main>
            </div>

            {isAnalysisModalOpen && (
                <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
                    <div className="bg-white rounded-lg shadow-2xl max-w-2xl w-full max-h-[90vh] flex flex-col">
                        <div className="p-4 border-b flex justify-between items-center">
                            <h2 className="text-xl font-bold text-gray-800">Lab Analysis</h2>
                            <button onClick={() => setIsAnalysisModalOpen(false)} className="p-1 rounded-full hover:bg-gray-200">
                                <XCircle size={24} className="text-gray-600" />
                            </button>
                        </div>
                        <div className="p-6 overflow-y-auto">
                            {isAnalyzing ? (
                                <div className="flex flex-col items-center justify-center py-12">
                                    <Loader className="animate-spin h-10 w-10 text-purple-600" />
                                    <p className="mt-4 text-lg font-semibold text-gray-700">Generating analysis...</p>
                                </div>
                            ) : (
                                <div className="prose prose-sm max-w-none" dangerouslySetInnerHTML={{ __html: analysisResult.replace(/\n/g, '<br />') }}></div>
                            )}
                        </div>
                        <div className="p-4 bg-gray-50 border-t flex justify-end items-center gap-4">
                            {ghlStatus && <p className="text-sm text-gray-600">{ghlStatus}</p>}
                            <button 
                                onClick={handleSendToGHL} 
                                disabled={isAnalyzing || isSendingToGHL}
                                className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-all disabled:bg-blue-300 disabled:cursor-not-allowed"
                            >
                                {isSendingToGHL ? <Loader className="animate-spin" size={20}/> : <Send size={20} />}
                                {isSendingToGHL ? 'Sending...' : 'Send to GoHighLevel'}
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
