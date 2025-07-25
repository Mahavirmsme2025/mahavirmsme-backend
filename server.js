const express = require('express');
const fs = require('fs').promises;
const path = require('path');
const XLSX = require('xlsx');
const cors = require('cors'); // Import cors at the top

const app = express();
const PORT = process.env.PORT || 3000;

// --- Configuration ---
const FILES_DIR = path.join(__dirname, 'serve_files'); // Directory for all servable files
const PROJECT_REPORTS_DIR = path.join(FILES_DIR, 'ProjectReports2');
const EXCEL_FILE = path.join(__dirname, 'contacts.xlsx');


// --- Middleware (MUST be before API endpoints) ---

// 1. Configure CORS to allow requests from your Netlify frontend
app.use(cors({ origin: "https://mahavirmsme.netlify.app" }));

// 2. Use the built-in Express middleware to parse JSON request bodies
app.use(express.json());

// 3. Serve the files in 'serve_files' so they are downloadable
// For example, a file at '/ProjectReports2/Category/file.pdf' can be downloaded
app.use(express.static(FILES_DIR));


// --- API Endpoints ---

// API: List all categories (subfolders) in ProjectReports2
app.get('/api/project-report-categories', async (req, res) => {
    try {
        const files = await fs.readdir(PROJECT_REPORTS_DIR, { withFileTypes: true });
        const categories = files
            .filter(file => file.isDirectory())
            .map(file => file.name)
            .sort((a, b) => a.localeCompare(b));
        res.json(categories);
    } catch (error) {
        console.error('Error reading categories directory:', error);
        res.status(500).json({ error: 'Unable to list categories' });
    }
});

// API: List all PDFs in a given category
app.get('/api/project-reports', async (req, res) => {
    const { category } = req.query;
    if (!category) {
        return res.status(400).json({ error: 'Category is required.' });
    }
    try {
        const categoryDir = path.join(PROJECT_REPORTS_DIR, category);
        const files = await fs.readdir(categoryDir);
        const pdfs = files
            .filter(file => file.toLowerCase().endsWith('.pdf'))
            .map(file => ({
                name: path.basename(file, '.pdf').replace(/_/g, ' '),
                // Provide the correct public URL for the file to be downloaded
                file: `/${encodeURIComponent('ProjectReports2')}/${encodeURIComponent(category)}/${encodeURIComponent(file)}`
            }));
        res.json(pdfs);
    } catch (error) {
        console.error('Error listing files in category:', error);
        res.status(500).json({ error: 'Unable to list files for the specified category' });
    }
});

// API: Save contact details to the Excel file
app.post('/api/contact', async (req, res) => {
    const { name, email, mobile } = req.body;
    if (!name || !email || !mobile) {
        return res.status(400).json({ error: 'All fields are required.' });
    }
    try {
        let contacts = [];
        try {
            await fs.access(EXCEL_FILE);
            const workbook = XLSX.readFile(EXCEL_FILE);
            const sheetName = workbook.SheetNames[0];
            contacts = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        } catch {
            console.log('Contacts file not found, creating a new one.');
        }
        contacts.push({ name, email, mobile, date: new Date().toISOString() });
        const worksheet = XLSX.utils.json_to_sheet(contacts);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Contacts');
        XLSX.writeFile(workbook, EXCEL_FILE);
        res.status(201).json({ success: true, message: 'Contact saved.' });
    } catch (error) {
        console.error('Error saving contact to Excel:', error);
        res.status(500).json({ error: 'Failed to save contact.' });
    }
});


// --- Start Server (ONLY ONE app.listen call at the end) ---
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});