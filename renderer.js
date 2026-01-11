let selectedInputFile = null;
let selectedOutputDir = null;
let availableWorksheets = [];
let selectedWorksheets = [];
let currentWorkerId = null;
let isProcessing = false;

const inputFileEl = document.getElementById('inputFile');
const outputDirEl = document.getElementById('outputDir');
const selectFileBtn = document.getElementById('selectFileBtn');
const selectDirBtn = document.getElementById('selectDirBtn');
const enableLimitRowsEl = document.getElementById('enableLimitRows');
const limitRowsOptionsEl = document.getElementById('limitRowsOptions');
const maxRowsEl = document.getElementById('maxRows');
const enableLimitSizeEl = document.getElementById('enableLimitSize');
const limitSizeOptionsEl = document.getElementById('limitSizeOptions');
const maxFileSizeEl = document.getElementById('maxFileSize');
const convertBtn = document.getElementById('convertBtn');
const statusEl = document.getElementById('status');
const resultsEl = document.getElementById('results');
const worksheetSelectionEl = document.getElementById('worksheetSelection');
const worksheetListEl = document.getElementById('worksheetList');
const worksheetLoadingEl = document.getElementById('worksheetLoading');
const selectAllBtn = document.getElementById('selectAllBtn');
const deselectAllBtn = document.getElementById('deselectAllBtn');
const progressFillEl = document.getElementById('progressFill');
const progressTextEl = document.getElementById('progressText');
const cancelBtn = document.getElementById('cancelBtn');
const conversionProgressEl = document.getElementById('conversionProgress');
const conversionProgressFillEl = document.getElementById('conversionProgressFill');
const conversionProgressTextEl = document.getElementById('conversionProgressText');
const cancelConversionBtn = document.getElementById('cancelConversionBtn');
const currentWorksheetEl = document.getElementById('currentWorksheet');
const worksheetProgressEl = document.getElementById('worksheetProgress');
const totalRowsProcessedEl = document.getElementById('totalRowsProcessed');

// Tab elements
const converterTab = document.getElementById('converterTab');
const viewerTab = document.getElementById('viewerTab');
const converterContent = document.getElementById('converterContent');
const viewerContent = document.getElementById('viewerContent');

// File Viewer tab elements
const openFileViewerBtn = document.getElementById('openFileViewerBtn');
const openMultipleFilesBtn = document.getElementById('openMultipleFilesBtn');
const recentFiles = document.getElementById('recentFiles');
const recentFilesList = document.getElementById('recentFilesList');
const activeViewers = document.getElementById('activeViewers');
const activeViewersList = document.getElementById('activeViewersList');

// Set up progress event handlers
window.electronAPI.onExcelProgress((data) => {
  updateProgress(data.progress, data.message);
});

window.electronAPI.onExcelWorksheets((data) => {
  if (data.success) {
    handleWorksheetsReceived(data.worksheets);
  }
});

window.electronAPI.onExcelCancelled((data) => {
  handleProcessingCancelled();
});

window.electronAPI.onConversionProgress((data) => {
  updateConversionProgress(data);
});

// Function to update progress UI
function updateProgress(progress, message) {
  if (progressFillEl && progressTextEl) {
    progressFillEl.style.width = `${progress}%`;
    progressTextEl.textContent = `${Math.round(progress)}% - ${message}`;

    // Show progress container when progress starts
    if (progress > 0) {
      const progressContainer = document.querySelector('.progress-container');
      if (progressContainer) {
        progressContainer.style.display = 'block';
      }
    }
  }
}

// Function to handle received worksheets during processing
function handleWorksheetsReceived(worksheets) {
  // Update available worksheets in real-time
  if (worksheets && worksheets.length > 0) {
    availableWorksheets = worksheets;
    selectedWorksheets = worksheets.map(ws => ws.name || ws);

    // Show worksheets immediately while processing continues
    worksheetLoadingEl.style.display = 'none';
    worksheetListEl.style.display = 'block';
    displayWorksheets();
    convertBtn.disabled = false;

    showStatus(`Found ${availableWorksheets.length} worksheet(s) - Processing complete`, 'success');
  }
}

// Function to handle processing cancellation
function handleProcessingCancelled() {
  isProcessing = false;
  currentWorkerId = null;
  worksheetLoadingEl.style.display = 'none';
  showStatus('File processing cancelled', 'info');
  resetProgressUI();
}

// Function to reset progress UI
function resetProgressUI() {
  const progressContainer = document.querySelector('.progress-container');
  const cancelButton = document.getElementById('cancelBtn');

  if (progressContainer) progressContainer.style.display = 'none';
  if (cancelButton) cancelButton.style.display = 'none';
  if (progressFillEl) progressFillEl.style.width = '0%';
  if (progressTextEl) progressTextEl.textContent = '0%';
}

// Function to update conversion progress
function updateConversionProgress(data) {
  const { currentSheet, completed, total, totalRows } = data;

  // Update conversion progress container
  if (conversionProgressEl) {
    conversionProgressEl.style.display = 'block';
  }

  // Update progress bar
  const progressPercent = total > 0 ? (completed / total) * 100 : 0;
  if (conversionProgressFillEl) {
    conversionProgressFillEl.style.width = `${progressPercent}%`;
  }

  // Update progress text
  if (conversionProgressTextEl) {
    conversionProgressTextEl.textContent = `Converting ${currentSheet}... ${Math.round(progressPercent)}%`;
  }

  // Update detailed progress information
  if (currentWorksheetEl) {
    currentWorksheetEl.textContent = currentSheet || '-';
  }

  if (worksheetProgressEl) {
    worksheetProgressEl.textContent = `${completed} of ${total} worksheets`;
  }

  if (totalRowsProcessedEl) {
    totalRowsProcessedEl.textContent = totalRows ? totalRows.toLocaleString() : '0';
  }
}

// Function to reset conversion progress UI
function resetConversionProgressUI() {
  if (conversionProgressEl) {
    conversionProgressEl.style.display = 'none';
  }
  if (conversionProgressFillEl) {
    conversionProgressFillEl.style.width = '0%';
  }
  if (conversionProgressTextEl) {
    conversionProgressTextEl.textContent = 'Preparing conversion...';
  }
  if (currentWorksheetEl) {
    currentWorksheetEl.textContent = '-';
  }
  if (worksheetProgressEl) {
    worksheetProgressEl.textContent = '0 of 0 worksheets';
  }
  if (totalRowsProcessedEl) {
    totalRowsProcessedEl.textContent = '0';
  }
}

// Function to handle file selection and worksheet loading
async function handleFileSelection(filePath) {
  if (!filePath) return;

  // Prevent multiple simultaneous processing
  if (isProcessing) {
    showStatus('Another file is currently being processed. Please wait or cancel the current operation.', 'info');
    return;
  }

  selectedInputFile = filePath;
  inputFileEl.value = filePath;
  clearStatus();
  resetProgressUI();

  // Show worksheet selection area with loading state
  worksheetSelectionEl.style.display = 'block';
  worksheetLoadingEl.style.display = 'flex';
  worksheetListEl.style.display = 'none';
  convertBtn.disabled = true;
  isProcessing = true;

  // Show cancel button after a short delay (for larger files)
  setTimeout(() => {
    if (isProcessing) {
      const cancelButton = document.getElementById('cancelBtn');
      if (cancelButton) cancelButton.style.display = 'inline-block';
    }
  }, 2000);

  // Get worksheets from the file using background worker
  try {
    showStatus('Loading Excel file in background...', 'info');

    const worksheetsResult = await window.electronAPI.getWorksheets(filePath);

    isProcessing = false;
    resetProgressUI();

    // Hide loading, show worksheet list
    worksheetLoadingEl.style.display = 'none';
    worksheetListEl.style.display = 'block';

    if (worksheetsResult.success) {
      // If worksheets weren't already received via events, process them now
      if (!availableWorksheets.length) {
        availableWorksheets = worksheetsResult.worksheets;
        selectedWorksheets = worksheetsResult.worksheets.map(ws => ws.name || ws);
      }

      displayWorksheets();
      convertBtn.disabled = selectedWorksheets.length === 0;

      // Show complexity information if available
      const complexSheets = availableWorksheets.filter(ws => ws.complexity === 'large');
      if (complexSheets.length > 0) {
        showStatus(`Found ${availableWorksheets.length} worksheet(s). ${complexSheets.length} large worksheet(s) detected.`, 'success');
      } else {
        showStatus(`Found ${availableWorksheets.length} worksheet(s)`, 'success');
      }
    } else {
      showStatus(`Error loading worksheets: ${worksheetsResult.error}`, 'error');
      worksheetListEl.innerHTML = `<p style="padding: 16px; text-align: center; color: var(--color-error);">Failed to load worksheets: ${worksheetsResult.error}</p>`;
      convertBtn.disabled = true;
    }
  } catch (error) {
    isProcessing = false;
    resetProgressUI();
    worksheetLoadingEl.style.display = 'none';
    worksheetListEl.style.display = 'block';
    showStatus(`Error loading worksheets: ${error.message}`, 'error');
    worksheetListEl.innerHTML = `<p style="padding: 16px; text-align: center; color: var(--color-error);">Error: ${error.message}</p>`;
    convertBtn.disabled = true;
  }
}

// File selection from Browse button
selectFileBtn.addEventListener('click', async () => {
  const filePath = await window.electronAPI.selectFile();
  await handleFileSelection(filePath);
});

// File selection from menu (File > Open)
window.electronAPI.onFileSelected(async (filePath) => {
  await handleFileSelection(filePath);
});

// Display worksheets with checkboxes
function displayWorksheets() {
  if (availableWorksheets.length === 0) {
    worksheetListEl.innerHTML = '<p style="padding: 16px; text-align: center; color: var(--color-text-secondary);">No worksheets found</p>';
    return;
  }

  const html = availableWorksheets.map((ws, index) => {
    const worksheet = typeof ws === 'string' ? { name: ws, accessible: true } : ws;
    const isLarge = worksheet.complexity === 'large' || !worksheet.accessible;
    const estimatedInfo = worksheet.estimatedRows ? ` (~${worksheet.estimatedRows.toLocaleString()} rows)` : '';

    return `
    <label class="worksheet-item ${isLarge ? 'worksheet-item-large' : ''}">
      <input
        type="checkbox"
        data-worksheet="${worksheet.name}"
        ${selectedWorksheets.includes(worksheet.name) ? 'checked' : ''}
        title="${isLarge ? 'This worksheet is large and will use Python for conversion' : 'This worksheet can be processed quickly'}"
      >
      <span class="worksheet-name">${worksheet.name}${estimatedInfo}</span>
      ${isLarge ? '<span class="worksheet-warning">(large file - uses Python)</span>' : ''}
    </label>
    `;
  }).join('');

  worksheetListEl.innerHTML = html;

  // Add event listeners to checkboxes
  worksheetListEl.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
    checkbox.addEventListener('change', (e) => {
      const worksheetName = e.target.dataset.worksheet;
      if (e.target.checked) {
        if (!selectedWorksheets.includes(worksheetName)) {
          selectedWorksheets.push(worksheetName);
        }
      } else {
        selectedWorksheets = selectedWorksheets.filter(name => name !== worksheetName);
      }
      convertBtn.disabled = selectedWorksheets.length === 0;
    });
  });
}

// Select all worksheets
selectAllBtn.addEventListener('click', () => {
  selectedWorksheets = availableWorksheets.map(ws => ws.name);
  displayWorksheets();
  convertBtn.disabled = false;
});

// Deselect all worksheets
deselectAllBtn.addEventListener('click', () => {
  selectedWorksheets = [];
  displayWorksheets();
  convertBtn.disabled = true;
});

// Cancel button for file processing
cancelBtn.addEventListener('click', async () => {
  if (isProcessing && currentWorkerId) {
    try {
      await window.electronAPI.cancelExcelProcessing(currentWorkerId);
      showStatus('Cancelling file processing...', 'info');
    } catch (error) {
      console.error('Failed to cancel processing:', error);
      showStatus('Failed to cancel processing', 'error');
    }
  } else if (isProcessing) {
    // Cancel all if no specific worker ID
    try {
      await window.electronAPI.cancelExcelProcessing();
      showStatus('Cancelling file processing...', 'info');
    } catch (error) {
      console.error('Failed to cancel processing:', error);
      showStatus('Failed to cancel processing', 'error');
    }
  }
});

// Cancel button for conversion processing
cancelConversionBtn.addEventListener('click', async () => {
  try {
    await window.electronAPI.cancelExcelProcessing();
    showStatus('Cancelling conversion...', 'info');
    convertBtn.disabled = false;
    resetConversionProgressUI();
  } catch (error) {
    console.error('Failed to cancel conversion:', error);
    showStatus('Failed to cancel conversion', 'error');
  }
});

// Output directory selection
selectDirBtn.addEventListener('click', async () => {
  const dirPath = await window.electronAPI.selectOutputDir();
  if (dirPath) {
    selectedOutputDir = dirPath;
    outputDirEl.value = dirPath;
  }
});

// Toggle limit rows options
enableLimitRowsEl.addEventListener('change', () => {
  limitRowsOptionsEl.style.display = enableLimitRowsEl.checked ? 'block' : 'none';
});

// Toggle limit size options
enableLimitSizeEl.addEventListener('change', () => {
  limitSizeOptionsEl.style.display = enableLimitSizeEl.checked ? 'block' : 'none';
});

// Convert button
convertBtn.addEventListener('click', async () => {
  // Use input field value as fallback if selectedInputFile is not set
  const filePath = selectedInputFile || inputFileEl.value.trim();
  
  console.log('Renderer: Convert button clicked');
  console.log('Renderer: selectedInputFile:', selectedInputFile);
  console.log('Renderer: inputFileEl.value:', inputFileEl.value);
  console.log('Renderer: Using filePath:', filePath);
  
  if (!filePath) {
    console.error('Renderer: No file path available');
    showStatus('Please select an input file', 'error');
    return;
  }

  const maxRows = enableLimitRowsEl.checked ? parseInt(maxRowsEl.value) : null;
  const maxFileSize = enableLimitSizeEl.checked ? parseFloat(maxFileSizeEl.value) : null;

  if (enableLimitRowsEl.checked && (!maxRows || maxRows < 1)) {
    showStatus('Please enter a valid number of rows', 'error');
    return;
  }

  if (enableLimitSizeEl.checked && (!maxFileSize || maxFileSize <= 0)) {
    showStatus('Please enter a valid file size (greater than 0)', 'error');
    return;
  }

  // At least one splitting option should be enabled if either checkbox is checked
  if ((enableLimitRowsEl.checked || enableLimitSizeEl.checked) && !maxRows && !maxFileSize) {
    showStatus('Please enable at least one splitting option', 'error');
    return;
  }

  // Disable button and show progress
  convertBtn.disabled = true;
  showStatus('Converting...', 'info');
  resultsEl.innerHTML = '';
  resetConversionProgressUI();

  if (selectedWorksheets.length === 0) {
    showStatus('Please select at least one worksheet to convert', 'error');
    convertBtn.disabled = false;
    return;
  }

  const options = {
    inputFile: filePath,
    worksheets: selectedWorksheets,
    outputDir: selectedOutputDir,
    maxRows: maxRows,
    maxFileSize: maxFileSize
  };

  console.log('Renderer: Sending conversion request with options:', options);
  const result = await window.electronAPI.convertFile(options);
  console.log('Renderer: Received result:', result);

  convertBtn.disabled = false;
  resetConversionProgressUI();

  if (result.success) {
    const totalRows = result.totalRows || 0;
    const totalFiles = result.files ? result.files.length : 0;
    let statusMsg = `✓ Successfully converted ${selectedWorksheets.length} worksheet(s), ${totalRows} total rows, ${totalFiles} file(s) created`;
    if (result.errors && result.errors.length > 0) {
      statusMsg += `\n⚠ Some worksheets failed: ${result.errors.join('; ')}`;
    }
    showStatus(statusMsg, result.errors && result.errors.length > 0 ? 'info' : 'success');
    displayResults(result.files);
  } else {
    console.error('Renderer: Conversion failed:', result.error);
    showStatus(`Error: ${result.error}`, 'error');
  }
});

function showStatus(message, type) {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
}

function clearStatus() {
  statusEl.className = 'status';
  statusEl.textContent = '';
  resultsEl.innerHTML = '';
  resetConversionProgressUI();
}

function displayResults(files) {
  if (files.length === 0) return;

  const html = `
    <strong>Files created:</strong>
    <ul>
      ${files.map(file => `
        <li>
          <div>${file.rows} rows</div>
          <div class="file-path">${file.path}</div>
        </li>
      `).join('')}
    </ul>
  `;

  resultsEl.innerHTML = html;
}

// State for tracking viewer windows
let viewerWindows = [];
let recentFilesData = [];

// Tab switching
converterTab.addEventListener('click', () => {
  switchTab('converter');
});

viewerTab.addEventListener('click', () => {
  switchTab('viewer');
});

function switchTab(tab) {
  if (tab === 'converter') {
    converterTab.classList.add('active');
    viewerTab.classList.remove('active');
    converterContent.classList.add('active');
    viewerContent.classList.remove('active');
  } else {
    viewerTab.classList.add('active');
    converterTab.classList.remove('active');
    viewerContent.classList.add('active');
    converterContent.classList.remove('active');

    // Update the active viewers list when switching to viewer tab
    updateActiveViewersList();
  }
}

// File Viewer tab functionality
openFileViewerBtn.addEventListener('click', async () => {
  await openViewerWithFileSelection();
});

openMultipleFilesBtn.addEventListener('click', async () => {
  // Allow selecting multiple files
  const filePaths = await window.electronAPI.selectMultipleViewerFiles();
  if (filePaths && filePaths.length > 0) {
    // Open each file in a separate viewer window
    for (const filePath of filePaths) {
      await openSpecificFileInNewViewer(filePath);
    }
  }
});

// Open viewer with file selection (new streamlined workflow)
async function openViewerWithFileSelection() {
  try {
    console.log('openViewerWithFileSelection called');
    console.log('window.electronAPI available:', !!window.electronAPI);
    console.log('openViewerWithFile available:', !!window.electronAPI.openViewerWithFile);

    const result = await window.electronAPI.openViewerWithFile();

    if (result.success && result.filePath) {
      // Add to recent files
      addToRecentFiles(result.filePath);

      // Track the new window
      const windowId = Date.now();
      viewerWindows.push({
        id: windowId,
        filePath: result.filePath,
        fileName: getFileName(result.filePath),
        opened: new Date()
      });

      updateActiveViewersList();
      updateRecentFilesList();

      console.log(`Opened file viewer for: ${result.filePath}`);
    } else if (!result.cancelled) {
      console.error('Failed to open viewer with file:', result.error);
    }
  } catch (error) {
    console.error('Failed to open file viewer:', error);
  }
}

// Open a file in new viewer window (legacy function)
async function openFileInNewViewer() {
  const filePath = await window.electronAPI.selectViewerFile();
  if (filePath) {
    await openSpecificFileInNewViewer(filePath);
  }
}

// Open specific file in new viewer window
async function openSpecificFileInNewViewer(filePath) {
  try {
    const result = await window.electronAPI.openNewViewer();
    if (result.success) {
      // Add to recent files
      addToRecentFiles(filePath);

      // Track the new window
      const windowId = Date.now(); // Simple ID generation
      viewerWindows.push({
        id: windowId,
        filePath: filePath,
        fileName: getFileName(filePath),
        opened: new Date()
      });

      updateActiveViewersList();
      updateRecentFilesList();

      console.log(`Opened file viewer for: ${filePath}`);
    }
  } catch (error) {
    console.error('Failed to open file viewer:', error);
  }
}

// Utility functions
function getFileName(filePath) {
  return filePath.split(/[\\/]/).pop();
}

function addToRecentFiles(filePath) {
  // Remove if already exists
  recentFilesData = recentFilesData.filter(item => item.filePath !== filePath);

  // Add to beginning
  recentFilesData.unshift({
    filePath: filePath,
    fileName: getFileName(filePath),
    lastOpened: new Date()
  });

  // Keep only last 10
  recentFilesData = recentFilesData.slice(0, 10);
}

function updateRecentFilesList() {
  if (recentFilesData.length === 0) {
    recentFiles.style.display = 'none';
    return;
  }

  recentFiles.style.display = 'block';
  recentFilesList.innerHTML = '';

  recentFilesData.forEach((item, index) => {
    const li = document.createElement('li');
    li.textContent = item.fileName;
    li.title = item.filePath;
    li.addEventListener('click', () => {
      openSpecificFileInNewViewer(item.filePath);
    });
    recentFilesList.appendChild(li);
  });
}

function updateActiveViewersList() {
  if (viewerWindows.length === 0) {
    activeViewers.style.display = 'none';
    return;
  }

  activeViewers.style.display = 'block';
  activeViewersList.innerHTML = '';

  viewerWindows.forEach((window, index) => {
    const li = document.createElement('li');
    li.textContent = `${window.fileName} (${window.opened.toLocaleTimeString()})`;
    li.title = window.filePath;
    activeViewersList.appendChild(li);
  });
}
