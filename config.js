// Configuration - Paste your Google Apps Script Web App URL here
const GOOGLE_SHEETS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbz8e05jIbT0JyqpfhyotUchNtuhcGfVEFpu3CyOKXOmGyvj7_7vblRaiYX4I01EiQpL/exec';

// Sync to Google Sheets
async function syncToGoogleSheets(type, data) {
    if (!GOOGLE_SHEETS_WEB_APP_URL || GOOGLE_SHEETS_WEB_APP_URL === 'PASTE_YOUR_WEB_APP_URL_HERE') {
        console.log('Google Sheets not configured');
        return;
    }
    
    try {
        const now = new Date();
        const currentMonth = now.toLocaleDateString('en-US', { month: 'long' });
        const currentYear = now.getFullYear();
        
        const response = await fetch(GOOGLE_SHEETS_WEB_APP_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                type: type,
                department: typeof currentDepartment !== 'undefined' ? currentDepartment : 'rv',
                month: currentMonth,
                year: currentYear,
                monthYear: `${currentMonth} ${currentYear}`,
                ...data
            })
        });
        console.log(`✅ Synced to Google Sheets (${currentMonth} ${currentYear})`);
    } catch (error) {
        console.error('❌ Sync error:', error);
        // Don't throw the error to prevent breaking the application
        // Just log it and continue
    }
}
