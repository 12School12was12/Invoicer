// ===== State =====
const state = {
    csvData: {
        headers: [],
        rows: [],
        delimiter: ','
    },
    mapping: {},
    currentRowIndex: 0,
    hideEmptyFields: false,
    selectedRows: new Set(), // Selected row indices for export
    tableCollapsed: false,
    isExporting: false
};

// ===== DOM Elements =====
const elements = {
    // Upload
    uploadSection: document.getElementById('upload-section'),
    uploadArea: document.getElementById('upload-area'),
    csvInput: document.getElementById('csv-input'),
    uploadBtn: document.getElementById('upload-btn'),
    fileInfo: document.getElementById('file-info'),
    resetBtn: document.getElementById('reset-btn'),
    
    // Mapping
    mappingSection: document.getElementById('mapping-section'),
    autoMapBtn: document.getElementById('auto-map-btn'),
    applyMappingBtn: document.getElementById('apply-mapping-btn'),
    
    // Preview
    previewSection: document.getElementById('preview-section'),
    backToMappingBtn: document.getElementById('back-to-mapping-btn'),
    hideEmptyCheckbox: document.getElementById('hide-empty-fields'),
    prevBtn: document.getElementById('prev-btn'),
    nextBtn: document.getElementById('next-btn'),
    currentIndex: document.getElementById('current-index'),
    totalCount: document.getElementById('total-count'),
    invoiceCanvas: document.getElementById('invoice-canvas'),
    
    // Data Table
    toggleTableBtn: document.getElementById('toggle-table-btn'),
    tableToggleIcon: document.getElementById('table-toggle-icon'),
    dataTableContent: document.getElementById('data-table-content'),
    dataTableHeaderRow: document.getElementById('data-table-header-row'),
    dataTableBody: document.getElementById('data-table-body'),
    selectAllCheckbox: document.getElementById('select-all-checkbox'),
    selectAllBtn: document.getElementById('select-all-btn'),
    deselectAllBtn: document.getElementById('deselect-all-btn'),
    selectedCount: document.getElementById('selected-count'),
    exportBtn: document.getElementById('export-btn'),
    
    // Export Progress
    exportProgress: document.getElementById('export-progress'),
    progressFill: document.getElementById('progress-fill'),
    progressText: document.getElementById('progress-text')
};

// ===== Field Definitions =====
const fieldDefinitions = {
    invoice_number: { label: 'Invoice Number', aliases: ['invoice_number', 'inv_number', 'invoice', 'number', 'inv_no', 'invoice_no'] },
    issued_on_date: { label: 'Issued on', aliases: ['issued_on', 'issued_date', 'issue_date', 'date_issued', 'created', 'created_at'] },
    due_date: { label: 'Due date', aliases: ['due_date', 'due', 'payment_due', 'deadline'] },
    sale_date: { label: 'Date of sale', aliases: ['sale_date', 'supply_date', 'date_of_sale', 'service_date'] },
    
    client_contact_name: { label: 'Client Contact', aliases: ['client_name', 'contact_name', 'customer_name', 'billed_to', 'client', 'customer', 'name'] },
    client_company_name: { label: 'Client Company', aliases: ['client_company', 'company_name', 'customer_company', 'company', 'organization'] },
    client_email: { label: 'Client Email', aliases: ['client_email', 'email', 'customer_email', 'contact_email'] },
    client_address_line1: { label: 'Client Address 1', aliases: ['client_address', 'address', 'address_line1', 'street', 'client_address_1'] },
    client_address_line2: { label: 'Client Address 2', aliases: ['client_address_2', 'address_line2', 'city_state', 'postal'] },
    
    seller_company_name: { label: 'Seller Company', aliases: ['seller_company', 'from_company', 'seller', 'vendor', 'from'] },
    seller_address_line1: { label: 'Seller Address', aliases: ['seller_address', 'from_address', 'vendor_address'] },
    
    item_1_description: { label: 'Item Description', aliases: ['item_description', 'description', 'service', 'product', 'item', 'item_name'] },
    item_1_price: { label: 'Item Price', aliases: ['item_price', 'price', 'unit_price', 'rate'] },
    item_1_quantity: { label: 'Item Quantity', aliases: ['quantity', 'qty', 'item_quantity', 'count', 'units'] },
    item_1_tax_rate: { label: 'Tax Rate', aliases: ['tax_rate', 'tax', 'vat', 'item_tax'] },
    item_1_amount: { label: 'Item Amount', aliases: ['item_amount', 'line_total', 'line_amount', 'item_total', 'amount'] },
    
    subtotal_value: { label: 'Subtotal', aliases: ['subtotal', 'sub_total', 'net_total'] },
    total_value: { label: 'Total', aliases: ['total', 'grand_total', 'total_amount', 'invoice_total', 'final_total'] },
    invoice_note_text: { label: 'Invoice Note', aliases: ['note', 'notes', 'invoice_note', 'comment', 'comments', 'terms', 'memo'] }
};

// Canvas field mappings (field -> element ID)
const canvasFieldMap = {
    invoice_number: 'canvas-invoice-number',
    issued_on_date: 'canvas-issued-on',
    due_date: 'canvas-due-date',
    sale_date: 'canvas-sale-date',
    client_contact_name: 'canvas-client-name',
    client_company_name: 'canvas-client-company',
    client_email: 'canvas-client-email',
    client_address_line1: 'canvas-client-address1',
    client_address_line2: 'canvas-client-address2',
    seller_company_name: 'canvas-seller-name',
    seller_address_line1: 'canvas-seller-address',
    item_1_description: 'canvas-item-desc',
    item_1_price: 'canvas-item-price',
    item_1_quantity: 'canvas-item-qty',
    item_1_tax_rate: 'canvas-item-tax',
    item_1_amount: 'canvas-item-amount',
    subtotal_value: 'canvas-subtotal',
    total_value: 'canvas-total',
    invoice_note_text: 'canvas-note'
};

// ===== CSV Parsing =====
function detectDelimiter(text) {
    const firstLine = text.split('\n')[0];
    const delimiters = [',', ';', '\t'];
    let maxCount = 0;
    let bestDelimiter = ',';
    
    for (const delimiter of delimiters) {
        const count = (firstLine.match(new RegExp(escapeRegex(delimiter), 'g')) || []).length;
        if (count > maxCount) {
            maxCount = count;
            bestDelimiter = delimiter;
        }
    }
    
    return bestDelimiter;
}

function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function parseCSV(text) {
    const delimiter = detectDelimiter(text);
    const lines = text.split(/\r?\n/).filter(line => line.trim());
    
    if (lines.length < 2) {
        throw new Error('CSV должен содержать заголовок и хотя бы одну строку данных');
    }
    
    const headers = parseCSVLine(lines[0], delimiter);
    const rows = [];
    
    for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i], delimiter);
        if (values.length === headers.length) {
            const row = {};
            headers.forEach((header, index) => {
                row[header] = values[index];
            });
            rows.push(row);
        }
    }
    
    return { headers, rows, delimiter };
}

function parseCSVLine(line, delimiter) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        const nextChar = line[i + 1];
        
        if (inQuotes) {
            if (char === '"') {
                if (nextChar === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = false;
                }
            } else {
                current += char;
            }
        } else {
            if (char === '"') {
                inQuotes = true;
            } else if (char === delimiter) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
    }
    
    result.push(current.trim());
    return result;
}

// ===== Mapping =====
function populateMappingDropdowns() {
    const selects = document.querySelectorAll('.mapping-row select');
    const { headers, rows } = state.csvData;
    
    selects.forEach(select => {
        select.innerHTML = '<option value="">— Не выбрано —</option>';
        
        headers.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            
            // Get sample value from first non-empty row
            let sampleValue = '';
            for (const row of rows) {
                if (row[header] && row[header].trim()) {
                    sampleValue = row[header].trim();
                    if (sampleValue.length > 30) {
                        sampleValue = sampleValue.substring(0, 30) + '...';
                    }
                    break;
                }
            }
            
            option.textContent = sampleValue 
                ? `${header} — пример: "${sampleValue}"`
                : header;
            
            select.appendChild(option);
        });
    });
}

function autoMapFields() {
    const { headers } = state.csvData;
    const selects = document.querySelectorAll('.mapping-row select');
    
    selects.forEach(select => {
        const fieldName = select.dataset.field;
        const fieldDef = fieldDefinitions[fieldName];
        
        if (!fieldDef) return;
        
        // Try to find matching header
        const normalizedAliases = fieldDef.aliases.map(a => a.toLowerCase().replace(/[_\s-]/g, ''));
        
        for (const header of headers) {
            const normalizedHeader = header.toLowerCase().replace(/[_\s-]/g, '');
            
            if (normalizedAliases.includes(normalizedHeader)) {
                select.value = header;
                break;
            }
        }
    });
    
    saveMappingState();
}

function saveMappingState() {
    const selects = document.querySelectorAll('.mapping-row select');
    state.mapping = {};
    
    selects.forEach(select => {
        const fieldName = select.dataset.field;
        if (select.value) {
            state.mapping[fieldName] = select.value;
        }
    });
}

// ===== Invoice Rendering =====
function renderInvoice(rowIndex) {
    const row = state.csvData.rows[rowIndex];
    if (!row) return;
    
    const hideEmpty = state.hideEmptyFields;
    
    // Update each canvas field
    Object.entries(canvasFieldMap).forEach(([field, elementId]) => {
        const element = document.getElementById(elementId);
        if (!element) return;
        
        const csvColumn = state.mapping[field];
        const value = csvColumn ? (row[csvColumn] || '') : '';
        
        element.textContent = value;
        
        // Handle empty field visibility
        if (hideEmpty && !value.trim()) {
            element.classList.add('hidden-field');
        } else {
            element.classList.remove('hidden-field');
        }
    });
    
    // Update seller title in header (same as seller_company_name)
    const sellerTitle = document.getElementById('canvas-seller-title');
    if (sellerTitle) {
        const csvColumn = state.mapping['seller_company_name'];
        const value = csvColumn ? (row[csvColumn] || 'COMPANY NAME') : 'COMPANY NAME';
        sellerTitle.textContent = value;
    }
    
    // Update navigation
    updateNavigation();
}

function updateNavigation() {
    const { rows } = state.csvData;
    const total = rows.length;
    const current = state.currentRowIndex + 1;
    
    elements.currentIndex.textContent = current;
    elements.totalCount.textContent = total;
    
    elements.prevBtn.disabled = state.currentRowIndex === 0;
    elements.nextBtn.disabled = state.currentRowIndex === total - 1;
    
    // Highlight current row in data table
    highlightCurrentRow();
}

// ===== Data Table =====
function renderDataTable() {
    const { headers, rows } = state.csvData;
    
    // Clear existing content
    elements.dataTableHeaderRow.innerHTML = `
        <th class="col-checkbox"><input type="checkbox" id="select-all-checkbox"></th>
        <th class="col-row-num">#</th>
    `;
    elements.dataTableBody.innerHTML = '';
    
    // Add header columns (all columns from CSV in original order)
    const displayHeaders = headers;
    displayHeaders.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        elements.dataTableHeaderRow.appendChild(th);
    });
    
    // Add rows
    rows.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.dataset.rowIndex = index;
        
        // Checkbox cell
        const checkboxTd = document.createElement('td');
        checkboxTd.className = 'col-checkbox';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = state.selectedRows.has(index);
        checkbox.addEventListener('change', (e) => {
            e.stopPropagation();
            toggleRowSelection(index);
        });
        checkboxTd.appendChild(checkbox);
        tr.appendChild(checkboxTd);
        
        // Row number
        const numTd = document.createElement('td');
        numTd.className = 'col-row-num';
        numTd.textContent = index + 1;
        tr.appendChild(numTd);
        
        // Data cells
        displayHeaders.forEach(header => {
            const td = document.createElement('td');
            td.textContent = row[header] || '';
            td.title = row[header] || '';
            tr.appendChild(td);
        });
        
        // Click to preview
        tr.addEventListener('click', (e) => {
            if (e.target.type !== 'checkbox') {
                state.currentRowIndex = index;
                renderInvoice(index);
            }
        });
        
        // Apply selected class
        if (state.selectedRows.has(index)) {
            tr.classList.add('selected');
        }
        
        elements.dataTableBody.appendChild(tr);
    });
    
    // Re-attach select all checkbox listener
    const selectAllCheckbox = document.getElementById('select-all-checkbox');
    selectAllCheckbox.addEventListener('change', handleSelectAllChange);
    
    updateSelectedCount();
    highlightCurrentRow();
}

function toggleRowSelection(index) {
    if (state.selectedRows.has(index)) {
        state.selectedRows.delete(index);
    } else {
        state.selectedRows.add(index);
    }
    
    const row = elements.dataTableBody.querySelector(`tr[data-row-index="${index}"]`);
    if (row) {
        row.classList.toggle('selected', state.selectedRows.has(index));
        const checkbox = row.querySelector('input[type="checkbox"]');
        if (checkbox) checkbox.checked = state.selectedRows.has(index);
    }
    
    updateSelectedCount();
    updateSelectAllCheckbox();
}

function selectAllRows() {
    state.csvData.rows.forEach((_, index) => {
        state.selectedRows.add(index);
    });
    updateTableCheckboxes();
    updateSelectedCount();
}

function deselectAllRows() {
    state.selectedRows.clear();
    updateTableCheckboxes();
    updateSelectedCount();
}

function updateTableCheckboxes() {
    const rows = elements.dataTableBody.querySelectorAll('tr');
    rows.forEach(row => {
        const index = parseInt(row.dataset.rowIndex);
        const checkbox = row.querySelector('input[type="checkbox"]');
        const isSelected = state.selectedRows.has(index);
        
        if (checkbox) checkbox.checked = isSelected;
        row.classList.toggle('selected', isSelected);
    });
    updateSelectAllCheckbox();
}

function updateSelectAllCheckbox() {
    const selectAllCheckbox = document.getElementById('select-all-checkbox');
    if (!selectAllCheckbox) return;
    
    const totalRows = state.csvData.rows.length;
    const selectedCount = state.selectedRows.size;
    
    selectAllCheckbox.checked = selectedCount === totalRows && totalRows > 0;
    selectAllCheckbox.indeterminate = selectedCount > 0 && selectedCount < totalRows;
}

function handleSelectAllChange(e) {
    if (e.target.checked) {
        selectAllRows();
    } else {
        deselectAllRows();
    }
}

function updateSelectedCount() {
    elements.selectedCount.textContent = state.selectedRows.size;
}

function highlightCurrentRow() {
    const rows = elements.dataTableBody.querySelectorAll('tr');
    rows.forEach(row => {
        const index = parseInt(row.dataset.rowIndex);
        row.classList.toggle('active', index === state.currentRowIndex);
    });
}

function toggleTableCollapse() {
    state.tableCollapsed = !state.tableCollapsed;
    elements.dataTableContent.classList.toggle('collapsed', state.tableCollapsed);
    elements.tableToggleIcon.classList.toggle('collapsed', state.tableCollapsed);
}

// ===== Export =====
async function exportSelectedInvoices() {
    if (state.selectedRows.size === 0) {
        alert('Выберите хотя бы один инвойс для экспорта');
        return;
    }
    
    if (state.isExporting) return;
    state.isExporting = true;
    
    // Show progress
    elements.exportProgress.classList.remove('hidden');
    elements.exportBtn.disabled = true;
    
    const zip = new JSZip();
    const selectedIndices = Array.from(state.selectedRows).sort((a, b) => a - b);
    const total = selectedIndices.length;
    
    try {
        for (let i = 0; i < selectedIndices.length; i++) {
            const rowIndex = selectedIndices[i];
            
            // Update progress
            const progress = Math.round(((i + 1) / total) * 100);
            elements.progressFill.style.width = `${progress}%`;
            elements.progressText.textContent = `Генерация инвойса ${i + 1} из ${total}...`;
            
            // Render the invoice for this row
            renderInvoice(rowIndex);
            
            // Wait for render
            await new Promise(resolve => setTimeout(resolve, 100));
            
            // Capture canvas as image
            const canvas = await html2canvas(elements.invoiceCanvas, {
                scale: 2,
                useCORS: true,
                backgroundColor: '#ffffff'
            });
            
            // Convert to blob
            const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
            
            // Get invoice number for filename
            const row = state.csvData.rows[rowIndex];
            const invoiceNumCol = state.mapping['invoice_number'];
            const invoiceNum = invoiceNumCol ? (row[invoiceNumCol] || `invoice_${rowIndex + 1}`) : `invoice_${rowIndex + 1}`;
            const safeFilename = invoiceNum.replace(/[^a-zA-Z0-9_-]/g, '_');
            
            // Add to zip
            zip.file(`${safeFilename}.png`, blob);
        }
        
        // Generate and download zip
        elements.progressText.textContent = 'Создание архива...';
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        
        // Download
        const timestamp = new Date().toISOString().slice(0, 10);
        saveAs(zipBlob, `invoices_${timestamp}.zip`);
        
        elements.progressText.textContent = 'Готово!';
        
        // Hide progress after delay
        setTimeout(() => {
            elements.exportProgress.classList.add('hidden');
            elements.progressFill.style.width = '0%';
        }, 2000);
        
    } catch (error) {
        console.error('Export error:', error);
        alert('Ошибка при экспорте: ' + error.message);
        elements.exportProgress.classList.add('hidden');
    } finally {
        state.isExporting = false;
        elements.exportBtn.disabled = false;
        
        // Restore current row view
        renderInvoice(state.currentRowIndex);
    }
}

// ===== Event Handlers =====
function handleFileSelect(file) {
    if (!file) return;
    
    const reader = new FileReader();
    
    reader.onload = (e) => {
        try {
            const text = e.target.result;
            state.csvData = parseCSV(text);
            
            // Update UI
            elements.uploadArea.classList.add('hidden');
            elements.fileInfo.classList.remove('hidden');
            elements.fileInfo.querySelector('.file-name').textContent = file.name;
            elements.fileInfo.querySelector('.file-stats').textContent = 
                `${state.csvData.rows.length} строк · ${state.csvData.headers.length} колонок`;
            
            // Show mapping section
            elements.mappingSection.classList.remove('hidden');
            populateMappingDropdowns();
            
        } catch (error) {
            alert('Ошибка при чтении CSV: ' + error.message);
        }
    };
    
    reader.readAsText(file);
}

function handleDragOver(e) {
    e.preventDefault();
    elements.uploadArea.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    elements.uploadArea.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    elements.uploadArea.classList.remove('dragover');
    
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.csv')) {
        handleFileSelect(file);
    } else {
        alert('Пожалуйста, загрузите CSV-файл');
    }
}

function handleReset() {
    state.csvData = { headers: [], rows: [], delimiter: ',' };
    state.mapping = {};
    state.currentRowIndex = 0;
    state.selectedRows.clear();
    
    elements.uploadArea.classList.remove('hidden');
    elements.fileInfo.classList.add('hidden');
    elements.mappingSection.classList.add('hidden');
    elements.previewSection.classList.add('hidden');
    elements.csvInput.value = '';
}

function handleApplyMapping() {
    saveMappingState();
    
    if (Object.keys(state.mapping).length === 0) {
        alert('Сопоставьте хотя бы одно поле');
        return;
    }
    
    state.currentRowIndex = 0;
    state.selectedRows.clear(); // Reset selection
    elements.previewSection.classList.remove('hidden');
    renderInvoice(0);
    renderDataTable();
    
    // Scroll to preview
    elements.previewSection.scrollIntoView({ behavior: 'smooth' });
}

function handlePrevRow() {
    if (state.currentRowIndex > 0) {
        state.currentRowIndex--;
        renderInvoice(state.currentRowIndex);
    }
}

function handleNextRow() {
    if (state.currentRowIndex < state.csvData.rows.length - 1) {
        state.currentRowIndex++;
        renderInvoice(state.currentRowIndex);
    }
}

function handleHideEmptyToggle() {
    state.hideEmptyFields = elements.hideEmptyCheckbox.checked;
    renderInvoice(state.currentRowIndex);
}

function handleBackToMapping() {
    elements.previewSection.classList.add('hidden');
    elements.mappingSection.scrollIntoView({ behavior: 'smooth' });
}

// ===== Keyboard Navigation =====
function handleKeyDown(e) {
    // Only handle if preview section is visible
    if (elements.previewSection.classList.contains('hidden')) return;
    
    if (e.key === 'ArrowLeft') {
        handlePrevRow();
    } else if (e.key === 'ArrowRight') {
        handleNextRow();
    }
}

// ===== Initialize =====
function init() {
    // Upload events
    elements.uploadBtn.addEventListener('click', () => elements.csvInput.click());
    elements.csvInput.addEventListener('change', (e) => handleFileSelect(e.target.files[0]));
    elements.uploadArea.addEventListener('dragover', handleDragOver);
    elements.uploadArea.addEventListener('dragleave', handleDragLeave);
    elements.uploadArea.addEventListener('drop', handleDrop);
    elements.resetBtn.addEventListener('click', handleReset);
    
    // Mapping events
    elements.autoMapBtn.addEventListener('click', autoMapFields);
    elements.applyMappingBtn.addEventListener('click', handleApplyMapping);
    
    // Preview events
    elements.backToMappingBtn.addEventListener('click', handleBackToMapping);
    elements.hideEmptyCheckbox.addEventListener('change', handleHideEmptyToggle);
    elements.prevBtn.addEventListener('click', handlePrevRow);
    elements.nextBtn.addEventListener('click', handleNextRow);
    
    // Data Table events
    elements.toggleTableBtn.addEventListener('click', toggleTableCollapse);
    elements.selectAllBtn.addEventListener('click', selectAllRows);
    elements.deselectAllBtn.addEventListener('click', deselectAllRows);
    elements.exportBtn.addEventListener('click', exportSelectedInvoices);
    
    // Keyboard navigation
    document.addEventListener('keydown', handleKeyDown);
    
    // Save mapping on any dropdown change
    document.querySelectorAll('.mapping-row select').forEach(select => {
        select.addEventListener('change', saveMappingState);
    });
}

// Load test data if ?test parameter is present
async function loadTestData() {
    try {
        const response = await fetch('sample-invoices.csv');
        const text = await response.text();
        state.csvData = parseCSV(text);
        
        // Update UI
        elements.uploadArea.classList.add('hidden');
        elements.fileInfo.classList.remove('hidden');
        elements.fileInfo.querySelector('.file-name').textContent = 'sample-invoices.csv';
        elements.fileInfo.querySelector('.file-stats').textContent = 
            `${state.csvData.rows.length} строк · ${state.csvData.headers.length} колонок`;
        
        // Show mapping section
        elements.mappingSection.classList.remove('hidden');
        populateMappingDropdowns();
        
        // Auto-map fields
        autoMapFields();
        
    } catch (error) {
        console.error('Failed to load test data:', error);
    }
}

// Start the app
document.addEventListener('DOMContentLoaded', () => {
    init();
    
    // Check for test mode
    if (window.location.search.includes('test')) {
        loadTestData();
    }
});

