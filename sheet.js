let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations
let subsheetNames = []; // This holds the names of subsheets

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        
        // Load the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        // Get all subsheet names
        subsheetNames = workbook.SheetNames.filter(name => name !== firstSheetName);
        populateSubsheetSelect(subsheetNames);

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the full sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Populate the subsheet dropdown
function populateSubsheetSelect(names) {
    const subsheetSelect = document.getElementById('subsheet-select');
    names.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        subsheetSelect.appendChild(option);
    });

    // Event listener for subsheet selection
    subsheetSelect.addEventListener('change', async (event) => {
        const selectedSubsheet = event.target.value;
        if (selectedSubsheet) {
            await loadSubsheet(selectedSubsheet);
        }
    });
}

// Function to load and display a selected subsheet
async function loadSubsheet(sheetName) {
    try {
        const response = await fetch(fileUrl); // You may need to store fileUrl globally or retrieve it appropriately
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading subsheet:", error);
    }
}

// The rest of your existing functions remain unchanged...

// Load the Excel sheet when the page is loaded (replace with your file URL)
window.addEventListener('load', () => {
    const fileUrl = getQueryParam('fileUrl'); // Assuming you get file URL from query params
    loadExcelSheet(fileUrl);
});
