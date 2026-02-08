// API Configuration
const API_URL_STORAGE_KEY = 'NODENBERG_API_BASE_URL';
const API_KEY_STORAGE_KEY = 'NODENBERG_API_KEY';

// Docker compose の既定（03_docker-version/docker-compose.yml: 3200:3100）
const DEFAULT_API_URL = 'http://localhost:3200';

// Get API Base URL and API Key from localStorage
let API_BASE_URL = localStorage.getItem(API_URL_STORAGE_KEY) || DEFAULT_API_URL;
let API_KEY = localStorage.getItem(API_KEY_STORAGE_KEY) || '';

// State
let templateBase64 = null;
let detectedPlaceholders = [];

// DOM Elements
const elements = {
  templateFile: document.getElementById('templateFile'),
  fileInfo: document.getElementById('file-info'),
  detectPlaceholders: document.getElementById('detectPlaceholders'),
  detectPlaceholdersDetailed: document.getElementById('detectPlaceholdersDetailed'),
  getTemplateInfo: document.getElementById('getTemplateInfo'),
  getTemplateSheets: document.getElementById('getTemplateSheets'),
  validateTemplate: document.getElementById('validateTemplate'),
  placeholdersResult: document.getElementById('placeholders-result'),
  dataInput: document.getElementById('dataInput'),
  useSampleData: document.getElementById('useSampleData'),
  generateExcel: document.getElementById('generateExcel'),
  generatePDF: document.getElementById('generatePDF'),
  generateExcelByDisplayOrder: document.getElementById('generateExcelByDisplayOrder'),
  generatePDFByDisplayOrder: document.getElementById('generatePDFByDisplayOrder'),
  sheetSelectBy: document.getElementById('sheetSelectBy'),
  sheetSelectValue: document.getElementById('sheetSelectValue'),
  generationResult: document.getElementById('generation-result'),
  errorDisplay: document.getElementById('error-display'),
  healthStatus: document.getElementById('health-status'),
  healthIndicator: document.getElementById('health-indicator'),
  healthText: document.getElementById('health-text'),
  apiBaseUrl: document.getElementById('api-base-url'),
  settingsBtn: document.getElementById('settings-btn'),
  settingsModal: document.getElementById('settings-modal'),
  modalClose: document.getElementById('modal-close'),
  apiUrlInput: document.getElementById('api-url-input'),
  apiKeyInput: document.getElementById('api-key-input'),
  saveUrlBtn: document.getElementById('save-url-btn'),
  resetUrlBtn: document.getElementById('reset-url-btn'),
};

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  elements.apiBaseUrl.textContent = API_BASE_URL;
  elements.apiUrlInput.value = API_BASE_URL;
  elements.apiKeyInput.value = API_KEY;
  checkHealth();
  setupEventListeners();
});

// Event Listeners
function setupEventListeners() {
  // Allow selecting the same file repeatedly by clearing current value before picker opens.
  elements.templateFile.addEventListener('click', () => {
    elements.templateFile.value = '';
  });
  elements.templateFile.addEventListener('change', handleFileUpload);
  elements.detectPlaceholders.addEventListener('click', () => detectPlaceholders(false));
  elements.detectPlaceholdersDetailed.addEventListener('click', () => detectPlaceholders(true));
  elements.getTemplateInfo.addEventListener('click', getTemplateInfo);
  elements.getTemplateSheets.addEventListener('click', getTemplateSheets);
  elements.validateTemplate.addEventListener('click', validateTemplate);
  elements.useSampleData.addEventListener('click', useSampleData);
  elements.generateExcel.addEventListener('click', generateExcel);
  elements.generatePDF.addEventListener('click', generatePDF);
  elements.generateExcelByDisplayOrder.addEventListener('click', generateExcelByDisplayOrder);
  elements.generatePDFByDisplayOrder.addEventListener('click', generatePDFByDisplayOrder);
  elements.sheetSelectBy.addEventListener('change', handleSheetSelectorChange);
  elements.sheetSelectValue.addEventListener('input', () => hideError());

  // Settings modal
  elements.settingsBtn.addEventListener('click', openSettings);
  elements.modalClose.addEventListener('click', closeSettings);
  elements.saveUrlBtn.addEventListener('click', saveApiUrl);
  elements.resetUrlBtn.addEventListener('click', resetApiUrl);

  // Close modal on outside click
  window.addEventListener('click', (e) => {
    if (e.target === elements.settingsModal) {
      closeSettings();
    }
  });
}

function handleSheetSelectorChange() {
  const by = elements.sheetSelectBy.value;
  if (by === 'none') {
    elements.sheetSelectValue.value = '';
    elements.sheetSelectValue.disabled = true;
    elements.sheetSelectValue.type = 'text';
    elements.sheetSelectValue.placeholder = '(disabled)';
    return;
  }

  elements.sheetSelectValue.disabled = false;
  if (by === 'id') {
    elements.sheetSelectValue.type = 'number';
    elements.sheetSelectValue.placeholder = 'e.g. 1';
  } else {
    elements.sheetSelectValue.type = 'text';
    elements.sheetSelectValue.placeholder = 'e.g. 参照シート';
  }
}

// Settings Modal Functions
function openSettings() {
  elements.settingsModal.style.display = 'block';
  elements.apiUrlInput.value = API_BASE_URL;
  elements.apiKeyInput.value = API_KEY;
  elements.apiUrlInput.focus();
}

function closeSettings() {
  elements.settingsModal.style.display = 'none';
}

function saveApiUrl() {
  const newUrl = elements.apiUrlInput.value.trim();
  const newApiKey = elements.apiKeyInput.value.trim();

  if (!newUrl) {
    alert('Please enter a valid URL');
    return;
  }

  // Remove trailing slash if present
  const cleanUrl = newUrl.replace(/\/$/, '');

  // Save to localStorage
  localStorage.setItem(API_URL_STORAGE_KEY, cleanUrl);
  localStorage.setItem(API_KEY_STORAGE_KEY, newApiKey);

  // Reload page to apply new settings
  location.reload();
}

function resetApiUrl() {
  localStorage.removeItem(API_URL_STORAGE_KEY);
  localStorage.removeItem(API_KEY_STORAGE_KEY);
  location.reload();
}

// Helper function to create headers with API key
function getHeaders() {
  const headers = {
    'Content-Type': 'application/json'
  };

  if (API_KEY) {
    headers['X-API-Key'] = API_KEY;
  }

  return headers;
}

function getSheetSelectorPayload() {
  const by = elements.sheetSelectBy.value;
  const rawValue = elements.sheetSelectValue.value;

  // Only effective when both are set
  // NOTE: name selector must preserve leading/trailing spaces for exact sheet-name match
  if (by === 'none' || rawValue.trim() === '') return {};

  if (by === 'id') {
    const id = Number(rawValue.trim());
    if (!Number.isInteger(id) || id <= 0) {
      throw new Error('Sheet ID must be an integer (>= 1)');
    }
    return { sheetSelectBy: 'id', sheetSelectValue: id };
  }

  if (by === 'name') {
    return { sheetSelectBy: 'name', sheetSelectValue: rawValue };
  }

  return {};
}

// Health Check
async function checkHealth() {
  try {
    const response = await fetch(`${API_BASE_URL}/health`);
    const data = await response.json();

    if (data.status === 'ok') {
      elements.healthIndicator.textContent = '✅';
      elements.healthText.textContent = `Server OK - ${data.service} v${data.version}`;
      elements.healthStatus.classList.add('ok');
    } else {
      throw new Error('Server status not OK');
    }
  } catch (error) {
    elements.healthIndicator.textContent = '❌';
    elements.healthText.textContent = 'Server connection failed';
    showError(`Failed to connect to server at ${API_BASE_URL}. Make sure the server is running, or update the API URL in settings.`);
  }
}

// File Upload Handler
async function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  try {
    // Convert file to Base64
    templateBase64 = await fileToBase64(file);

    // Update UI
    elements.fileInfo.innerHTML = `
      <strong>✅ File loaded:</strong> ${file.name} (${(file.size / 1024).toFixed(2)} KB)
    `;
    elements.fileInfo.classList.add('active');

    // Enable detection buttons
    elements.detectPlaceholders.disabled = false;
    elements.detectPlaceholdersDetailed.disabled = false;
    elements.getTemplateInfo.disabled = false;
    elements.getTemplateSheets.disabled = false;
    elements.validateTemplate.disabled = false;

    // Enable generation buttons immediately after template upload
    elements.generateExcel.disabled = false;
    elements.generatePDF.disabled = false;
    elements.generateExcelByDisplayOrder.disabled = false;
    elements.generatePDFByDisplayOrder.disabled = false;

    hideError();
  } catch (error) {
    showError('Failed to read file: ' + error.message);
  }
}

// Convert File to Base64
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// Detect Placeholders
async function detectPlaceholders(detailed = false) {
  if (!templateBase64) {
    showError('Please upload a template file first');
    return;
  }

  try {
    showLoading(elements.placeholdersResult);

    const response = await fetch(`${API_BASE_URL}/template/placeholders`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({
        templateBase64,
        detailed,
      }),
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || 'Failed to detect placeholders');
    }

    // Store detected placeholders
    detectedPlaceholders = data.placeholders || [];

    // Display results
    elements.placeholdersResult.innerHTML = `
      <div class="success-message">✅ Detected ${detectedPlaceholders.length} placeholders</div>
      <pre>${JSON.stringify(data, null, 2)}</pre>
    `;
    elements.placeholdersResult.classList.add('active');

    // Enable generation buttons if placeholders found
    if (detectedPlaceholders.length > 0) {
      elements.generateExcel.disabled = false;
      elements.generatePDF.disabled = false;
      elements.generateExcelByDisplayOrder.disabled = false;
      elements.generatePDFByDisplayOrder.disabled = false;

      // Auto-generate sample data
      generateSampleData();
    }

    hideError();
  } catch (error) {
    showError('Placeholder detection failed: ' + error.message);
  }
}

// Get Template Info
async function getTemplateInfo() {
  if (!templateBase64) {
    showError('Please upload a template file first');
    return;
  }

  try {
    showLoading(elements.placeholdersResult);

    const response = await fetch(`${API_BASE_URL}/template/info`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({ templateBase64 }),
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || 'Failed to get template info');
    }

    // Display results
    elements.placeholdersResult.innerHTML = `
      <div class="success-message">✅ Template information retrieved</div>
      <pre>${JSON.stringify(data, null, 2)}</pre>
    `;
    elements.placeholdersResult.classList.add('active');

    hideError();
  } catch (error) {
    showError('Failed to get template info: ' + error.message);
  }
}

// Get Template Sheets
async function getTemplateSheets() {
  if (!templateBase64) {
    showError('Please upload a template file first');
    return;
  }

  try {
    showLoading(elements.placeholdersResult);

    const response = await fetch(`${API_BASE_URL}/template/sheets`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({ templateBase64 }),
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || 'Failed to get template sheets');
    }

    // Display results
    elements.placeholdersResult.innerHTML = `
      <div class="success-message">✅ Sheet list retrieved (${data.sheetCount || 0} sheets)</div>
      <pre>${JSON.stringify(data, null, 2)}</pre>
    `;
    elements.placeholdersResult.classList.add('active');

    hideError();
  } catch (error) {
    showError('Failed to get template sheets: ' + error.message);
  }
}

// Validate Template
async function validateTemplate() {
  if (!templateBase64) {
    showError('Please upload a template file first');
    return;
  }

  try {
    showLoading(elements.placeholdersResult);

    const fileName = elements.templateFile.files && elements.templateFile.files[0]
      ? elements.templateFile.files[0].name.replace(/\.xlsx$/i, '')
      : 'template';

    const response = await fetch(`${API_BASE_URL}/template/validate`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({
        templateId: `${fileName}-${Date.now()}`,
        templateName: fileName,
        base64Data: templateBase64,
        generateJsonTemplate: true,
      }),
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || 'Failed to validate template');
    }

    elements.placeholdersResult.innerHTML = `
      <div class="success-message">✅ Template validated</div>
      <pre>${JSON.stringify(data, null, 2)}</pre>
    `;
    elements.placeholdersResult.classList.add('active');

    hideError();
  } catch (error) {
    showError('Template validation failed: ' + error.message);
  }
}

// Generate Sample Data
function generateSampleData() {
  const sampleData = buildSampleDataFromDetectedPlaceholders(detectedPlaceholders);
  elements.dataInput.value = JSON.stringify(sampleData, null, 2);
}

// Use Sample Data Button
function useSampleData() {
  if (detectedPlaceholders.length === 0) {
    showError('Please detect placeholders first');
    return;
  }
  generateSampleData();
}

function extractPlaceholderKeys(placeholders) {
  if (!Array.isArray(placeholders)) return [];

  return placeholders
    .map((p) => {
      if (typeof p === 'string') return p;
      if (p && typeof p === 'object' && typeof p.key === 'string') return p.key;
      return null;
    })
    .filter((k) => typeof k === 'string');
}

function sampleValueForField(fieldName) {
  const lower = String(fieldName).toLowerCase();
  if (lower.includes('番号') || lower.includes('no')) return 1;
  if (lower.includes('数量') || lower.includes('qty') || lower.includes('count')) return 2;
  if (lower.includes('単価') || lower.includes('price') || lower.includes('amount') || lower.includes('金額')) return 1000;
  if (lower.includes('日付') || lower.includes('date')) return '2026/02/07';
  return `Sample ${fieldName}`;
}

function buildSampleDataFromDetectedPlaceholders(placeholders) {
  const keys = extractPlaceholderKeys(placeholders);
  const sampleData = {};
  const legacyArrays = new Map();   // arrayName => Set(fieldPath)
  const sectionTables = new Map();  // section.table => { section, table, fields:Set(cellPath) }

  keys.forEach((key) => {
    // New syntax: ##section.table.cell
    if (key.startsWith('##')) {
      const parts = key.slice(2).split('.');
      if (parts.length >= 3) {
        const section = parts[0];
        const table = parts[1];
        const cellPath = parts.slice(2).join('.');
        const tableKey = `${section}.${table}`;
        if (!sectionTables.has(tableKey)) {
          sectionTables.set(tableKey, { section, table, fields: new Set() });
        }
        sectionTables.get(tableKey).fields.add(cellPath);
        return;
      }
    }

    // Legacy array syntax: #array.field
    if (key.startsWith('#')) {
      const parts = key.slice(1).split('.');
      if (parts.length >= 2) {
        const arrayName = parts[0];
        const fieldPath = parts.slice(1).join('.');
        if (!legacyArrays.has(arrayName)) {
          legacyArrays.set(arrayName, new Set());
        }
        legacyArrays.get(arrayName).add(fieldPath);
        return;
      }
    }

    // Primitive placeholder
    sampleData[key] = sampleValueForField(key);
  });

  // Build legacy array sample: data[arrayName] = [{...}, {...}]
  legacyArrays.forEach((fields, arrayName) => {
    const row1 = {};
    const row2 = {};
    Array.from(fields).forEach((field, idx) => {
      row1[field] = sampleValueForField(field);
      row2[field] = sampleValueForField(`${field}_${idx + 2}`);
    });
    sampleData[arrayName] = [row1, row2];
  });

  // Build section-table sample: data[section][table] = [{...}, {...}]
  sectionTables.forEach(({ section, table, fields }) => {
    if (!sampleData[section] || typeof sampleData[section] !== 'object' || Array.isArray(sampleData[section])) {
      sampleData[section] = {};
    }

    const row1 = {};
    const row2 = {};
    Array.from(fields).forEach((field, idx) => {
      row1[field] = sampleValueForField(field);
      row2[field] = sampleValueForField(`${field}_${idx + 2}`);
    });
    sampleData[section][table] = [row1, row2];
  });

  return sampleData;
}

function getDisplayOrderPayload() {
  const rawValue = elements.sheetSelectValue.value.trim();
  const displayOrder = Number(rawValue);
  if (!Number.isInteger(displayOrder) || displayOrder <= 0) {
    throw new Error('DisplayOrder must be an integer (>= 1). Enter a number in the Sheet input.');
  }
  return { displayOrder };
}

// Generate Excel
async function generateExcel() {
  if (!templateBase64) {
    showError('Please upload a template file first');
    return;
  }

  const dataText = elements.dataInput.value.trim();
  if (!dataText) {
    showError('Please enter data in JSON format');
    return;
  }

  try {
    const data = JSON.parse(dataText);
    const sheetSelector = getSheetSelectorPayload();

    showLoading(elements.generationResult);

    const response = await fetch(`${API_BASE_URL}/generate/excel`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({
        templateBase64,
        data,
        ...sheetSelector,
      }),
    });

    const result = await response.json();

    if (!response.ok) {
      throw new Error(result.error || 'Failed to generate Excel');
    }

    // Display success and download button
    const filename = `generated-${Date.now()}.xlsx`;
    elements.generationResult.innerHTML = `
      <div class="success-message">✅ Excel file generated successfully!</div>
      <p>File size: ${(Buffer.from(result.data, 'base64').length / 1024).toFixed(2)} KB</p>
      <a href="#" class="download-link" onclick="downloadFile('${result.data}', '${filename}', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); return false;">
        📥 Download Excel File
      </a>
    `;
    elements.generationResult.classList.add('active');

    hideError();
  } catch (error) {
    if (error instanceof SyntaxError) {
      showError('Invalid JSON format in data input');
    } else {
      showError('Excel generation failed: ' + error.message);
    }
  }
}

// Generate PDF
async function generatePDF() {
  if (!templateBase64) {
    showError('Please upload a template file first');
    return;
  }

  const dataText = elements.dataInput.value.trim();
  if (!dataText) {
    showError('Please enter data in JSON format');
    return;
  }

  try {
    const data = JSON.parse(dataText);
    const sheetSelector = getSheetSelectorPayload();

    showLoading(elements.generationResult);

    const response = await fetch(`${API_BASE_URL}/generate/pdf`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({
        templateBase64,
        data,
        ...sheetSelector,
      }),
    });

    const result = await response.json();

    if (!response.ok) {
      throw new Error(result.error || 'Failed to generate PDF');
    }

    // Display success and download button
    const filename = `generated-${Date.now()}.pdf`;
    elements.generationResult.innerHTML = `
      <div class="success-message">✅ PDF file generated successfully!</div>
      <p>File size: ${(Buffer.from(result.data, 'base64').length / 1024).toFixed(2)} KB</p>
      <a href="#" class="download-link" onclick="downloadFile('${result.data}', '${filename}', 'application/pdf'); return false;">
        📥 Download PDF File
      </a>
    `;
    elements.generationResult.classList.add('active');

    hideError();
  } catch (error) {
    if (error instanceof SyntaxError) {
      showError('Invalid JSON format in data input');
    } else {
      showError('PDF generation failed: ' + error.message);
    }
  }
}

// Generate Excel by displayOrder
async function generateExcelByDisplayOrder() {
  if (!templateBase64) {
    showError('Please upload a template file first');
    return;
  }

  const dataText = elements.dataInput.value.trim();
  if (!dataText) {
    showError('Please enter data in JSON format');
    return;
  }

  try {
    const data = JSON.parse(dataText);
    const payload = getDisplayOrderPayload();

    showLoading(elements.generationResult);

    const response = await fetch(`${API_BASE_URL}/generate/excel/by-display-order`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({
        templateBase64,
        data,
        ...payload,
      }),
    });

    const result = await response.json();

    if (!response.ok) {
      throw new Error(result.error || result.details || 'Failed to generate Excel by displayOrder');
    }

    const filename = `generated-display-order-${Date.now()}.xlsx`;
    elements.generationResult.innerHTML = `
      <div class="success-message">✅ Excel file generated by displayOrder!</div>
      <p>File size: ${(Buffer.from(result.data, 'base64').length / 1024).toFixed(2)} KB</p>
      <a href="#" class="download-link" onclick="downloadFile('${result.data}', '${filename}', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); return false;">
        📥 Download Excel File
      </a>
    `;
    elements.generationResult.classList.add('active');

    hideError();
  } catch (error) {
    if (error instanceof SyntaxError) {
      showError('Invalid JSON format in data input');
    } else {
      showError('Excel generation by displayOrder failed: ' + error.message);
    }
  }
}

// Generate PDF by displayOrder
async function generatePDFByDisplayOrder() {
  if (!templateBase64) {
    showError('Please upload a template file first');
    return;
  }

  const dataText = elements.dataInput.value.trim();
  if (!dataText) {
    showError('Please enter data in JSON format');
    return;
  }

  try {
    const data = JSON.parse(dataText);
    const payload = getDisplayOrderPayload();

    showLoading(elements.generationResult);

    const response = await fetch(`${API_BASE_URL}/generate/pdf/by-display-order`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({
        templateBase64,
        data,
        ...payload,
      }),
    });

    const result = await response.json();

    if (!response.ok) {
      throw new Error(result.error || result.details || 'Failed to generate PDF by displayOrder');
    }

    const filename = `generated-display-order-${Date.now()}.pdf`;
    elements.generationResult.innerHTML = `
      <div class="success-message">✅ PDF file generated by displayOrder!</div>
      <p>File size: ${(Buffer.from(result.data, 'base64').length / 1024).toFixed(2)} KB</p>
      <a href="#" class="download-link" onclick="downloadFile('${result.data}', '${filename}', 'application/pdf'); return false;">
        📥 Download PDF File
      </a>
    `;
    elements.generationResult.classList.add('active');

    hideError();
  } catch (error) {
    if (error instanceof SyntaxError) {
      showError('Invalid JSON format in data input');
    } else {
      showError('PDF generation by displayOrder failed: ' + error.message);
    }
  }
}

// Download File from Base64
window.downloadFile = function(base64Data, filename, mimeType) {
  try {
    // Convert Base64 to Blob
    const byteCharacters = atob(base64Data);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: mimeType });

    // Create download link
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (error) {
    showError('Failed to download file: ' + error.message);
  }
};

// Show Loading
function showLoading(element) {
  element.innerHTML = '<div class="loading">⏳ Processing...</div>';
  element.classList.add('active');
}

// Show Error
function showError(message) {
  elements.errorDisplay.innerHTML = `<strong>❌ Error</strong>${message}`;
  elements.errorDisplay.style.display = 'block';
}

// Hide Error
function hideError() {
  elements.errorDisplay.style.display = 'none';
}

// Buffer polyfill for browser
const Buffer = {
  from: (data, encoding) => {
    if (encoding === 'base64') {
      const binary = atob(data);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      return { length: bytes.length };
    }
    return { length: 0 };
  }
};
