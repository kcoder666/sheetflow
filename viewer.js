// Simple File Viewer JavaScript
// Minimal interface for viewing Excel/CSV files

// DOM elements
const fileName = document.getElementById('fileName');
const fileDetails = document.getElementById('fileDetails');
const openFileBtn = document.getElementById('openFileBtn');
const worksheetSelect = document.getElementById('worksheetSelect');
const tableContainer = document.getElementById('tableContainer');
const tableNav = document.getElementById('tableNav');
const tableWrapper = document.getElementById('tableWrapper');
const dataTable = document.getElementById('dataTable');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const loadingState = document.getElementById('loadingState');
const emptyState = document.getElementById('emptyState');
const errorState = document.getElementById('errorState');
const errorMessage = document.getElementById('errorMessage');

// Navigation elements
const rowStart = document.getElementById('rowStart');
const rowEnd = document.getElementById('rowEnd');
const rowTotal = document.getElementById('rowTotal');
const currentPage = document.getElementById('currentPage');
const totalPages = document.getElementById('totalPages');
const prevBtn = document.getElementById('prevBtn');
const nextBtn = document.getElementById('nextBtn');

// State variables
let currentFile = null;
let currentSheet = null;
let fileType = null;
let currentPageNum = 1;
const pageSize = 100;
let totalRowCount = 0;
let columns = [];
let worksheets = [];

// Initialize viewer
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  showState('empty');

  // Listen for direct file opening from main process
  window.electronAPI.onOpenFileDirectly((filePath) => {
    console.log('Received file to open directly:', filePath);
    initializeFile(filePath);
  });
});

// Event listeners
function setupEventListeners() {
  openFileBtn.addEventListener('click', openFile);
  worksheetSelect.addEventListener('change', onWorksheetChange);
  prevBtn.addEventListener('click', () => navigateToPage(currentPageNum - 1));
  nextBtn.addEventListener('click', () => navigateToPage(currentPageNum + 1));
}

// File operations
async function openFile() {
  try {
    const filePath = await window.electronAPI.selectViewerFile();
    if (!filePath) return;

    showState('loading');
    await initializeFile(filePath);
  } catch (error) {
    console.error('Failed to open file:', error);
    showError('Failed to open file: ' + error.message);
  }
}

async function initializeFile(filePath) {
  try {
    // Initialize file for viewing
    const result = await window.electronAPI.viewerInitFile(filePath);

    if (!result.success) {
      throw new Error(result.error);
    }

    currentFile = filePath;
    fileType = result.fileInfo.type;

    // Update UI
    fileName.textContent = getFileName(filePath);

    if (fileType === 'excel' && result.worksheets && result.worksheets.length > 0) {
      // Show worksheet selector for Excel files
      worksheets = result.worksheets;
      populateWorksheetSelect();
      worksheetSelect.style.display = 'block';

      // Auto-select first worksheet
      if (worksheets.length > 0) {
        currentSheet = worksheets[0];
        worksheetSelect.value = currentSheet;
        await loadWorksheet();
      }
    } else {
      // CSV file - load directly
      worksheetSelect.style.display = 'none';
      await loadData();
    }
  } catch (error) {
    console.error('Failed to initialize file:', error);
    showError('Failed to load file: ' + error.message);
  }
}

function populateWorksheetSelect() {
  worksheetSelect.innerHTML = '<option value="">Select worksheet...</option>';
  worksheets.forEach(sheet => {
    const option = document.createElement('option');
    option.value = sheet;
    option.textContent = sheet;
    worksheetSelect.appendChild(option);
  });
}

async function onWorksheetChange() {
  const selectedSheet = worksheetSelect.value;
  if (!selectedSheet) return;

  try {
    currentSheet = selectedSheet;
    currentPageNum = 1;
    await loadWorksheet();
  } catch (error) {
    console.error('Failed to change worksheet:', error);
    showError('Failed to load worksheet: ' + error.message);
  }
}

async function loadWorksheet() {
  try {
    showState('loading');

    // Set the worksheet for Excel files
    if (fileType === 'excel') {
      await window.electronAPI.viewerSetSheet(currentSheet);
    }

    await loadData();
  } catch (error) {
    console.error('Failed to load worksheet:', error);
    showError('Failed to load worksheet: ' + error.message);
  }
}

async function loadData() {
  try {
    // Get file info first
    const fileInfoResult = await window.electronAPI.viewerGetFileInfo();
    if (fileInfoResult.success) {
      totalRowCount = fileInfoResult.fileInfo.totalRows || 0;
      updateFileDetails();
    }

    // Load the first page
    await loadPage(currentPageNum);
  } catch (error) {
    console.error('Failed to load data:', error);
    showError('Failed to load data: ' + error.message);
  }
}

async function loadPage(pageNum) {
  try {
    const startRow = (pageNum - 1) * pageSize;
    const result = await window.electronAPI.viewerReadPage(startRow, pageSize);

    if (!result.success) {
      throw new Error(result.error);
    }

    // Update data
    currentPageNum = pageNum;
    columns = result.headers || [];

    // Render table
    renderTable(result.headers, result.rows);
    updateNavigation();
    showState('table');

  } catch (error) {
    console.error('Failed to load page:', error);
    showError('Failed to load page: ' + error.message);
  }
}

function renderTable(headers, rows) {
  // Clear existing content
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  if (!headers || !rows) return;

  // Render headers
  const headerRow = document.createElement('tr');
  headers.forEach((header, index) => {
    const th = document.createElement('th');
    th.textContent = header || `Column ${index + 1}`;
    th.title = header || `Column ${index + 1}`;
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  // Render rows
  rows.forEach((rowData, rowIndex) => {
    const tr = document.createElement('tr');

    headers.forEach((header, colIndex) => {
      const td = document.createElement('td');
      const cellValue = rowData[colIndex] !== undefined ? rowData[colIndex] : '';
      td.textContent = cellValue;
      td.title = cellValue; // Show full content on hover
      tr.appendChild(td);
    });

    tableBody.appendChild(tr);
  });
}

function updateFileDetails() {
  if (totalRowCount > 0) {
    const rowText = totalRowCount.toLocaleString() + ' rows';
    const sheetText = currentSheet ? ` • ${currentSheet}` : '';
    fileDetails.textContent = rowText + sheetText;
  } else {
    fileDetails.textContent = currentSheet || '';
  }
}

function updateNavigation() {
  const totalPagesCount = Math.ceil(totalRowCount / pageSize);
  const startRowNum = (currentPageNum - 1) * pageSize + 1;
  const endRowNum = Math.min(currentPageNum * pageSize, totalRowCount);

  // Update display
  rowStart.textContent = startRowNum.toLocaleString();
  rowEnd.textContent = endRowNum.toLocaleString();
  rowTotal.textContent = totalRowCount.toLocaleString();
  currentPage.textContent = currentPageNum.toLocaleString();
  totalPages.textContent = totalPagesCount.toLocaleString();

  // Update button states
  prevBtn.disabled = currentPageNum <= 1;
  nextBtn.disabled = currentPageNum >= totalPagesCount;
}

async function navigateToPage(pageNum) {
  const totalPagesCount = Math.ceil(totalRowCount / pageSize);

  if (pageNum < 1 || pageNum > totalPagesCount) return;
  if (pageNum === currentPageNum) return;

  await loadPage(pageNum);
}

// UI state management
function showState(state) {
  // Hide all states
  loadingState.style.display = 'none';
  emptyState.style.display = 'none';
  errorState.style.display = 'none';
  tableNav.style.display = 'none';
  tableWrapper.style.display = 'none';

  // Show requested state
  switch (state) {
    case 'loading':
      loadingState.style.display = 'flex';
      break;
    case 'empty':
      emptyState.style.display = 'flex';
      break;
    case 'error':
      errorState.style.display = 'flex';
      break;
    case 'table':
      tableNav.style.display = 'flex';
      tableWrapper.style.display = 'block';
      break;
  }
}

function showError(message) {
  errorMessage.textContent = message;
  showState('error');
}

// Utility functions
function getFileName(filePath) {
  return filePath.split(/[\\/]/).pop();
}

function formatFileSize(bytes) {
  if (!bytes) return 'Unknown';

  const units = ['B', 'KB', 'MB', 'GB'];
  let size = bytes;
  let unitIndex = 0;

  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex++;
  }

  return `${size.toFixed(1)} ${units[unitIndex]}`;
}