// UI Elements
const overlay = document.getElementById('loading-overlay');
const overlayText = document.getElementById('overlay-text');
const consoleBox = document.getElementById('console-box');
const downloadContainer = document.getElementById('generator-download-container');

// Stats Elements
const statSkusCount = document.getElementById('stat-skus-count');
const statEngineStatus = document.getElementById('stat-engine-status');

// Setup Drag & Drop zones
setupDropZone('dz-learn-directory', 'input-learn-directory', 'info-learn-directory', 'directory');
setupDropZone('dz-learn-content', 'input-learn-content', 'info-learn-content', 'content');
setupDropZone('dz-list-template', 'input-list-template', 'info-list-template', 'template');

function setupDropZone(dzId, inputId, infoId, fileType) {
    const dz = document.getElementById(dzId);
    const input = document.getElementById(inputId);
    const info = document.getElementById(infoId);

    if (!dz || !input || !info) return;

    dz.addEventListener('dragover', (e) => {
        e.preventDefault();
        dz.classList.add('dragover');
    });

    dz.addEventListener('dragleave', () => {
        dz.classList.remove('dragover');
    });

    dz.addEventListener('drop', (e) => {
        e.preventDefault();
        dz.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            input.files = e.dataTransfer.files;
            handleFileSelect(input.files[0], dz, info, fileType);
        }
    });

    dz.addEventListener('change', () => {
        if (input.files.length > 0) {
            handleFileSelect(input.files[0], dz, info, fileType);
        }
    });
}

function handleFileSelect(file, dz, info, fileType) {
    dz.classList.add('filled');
    info.textContent = `${file.name} (${formatBytes(file.size)})`;
    logSystem(`Uploaded ${fileType} sheet: ${file.name}`);
    
    // Auto-detect engine category if template is selected
    if (fileType === 'template') {
        detectCategoryFromTemplateName(file.name);
    }
}

function detectCategoryFromTemplateName(fileName) {
    const lowerName = fileName.toLowerCase();
    const selectCategory = document.getElementById('select-category');
    if (!selectCategory) return;

    if (lowerName.includes('footwear')) {
        selectCategory.value = 'footwear';
        logInfo("Auto-detected footwear category based on template filename.");
        if (statEngineStatus) statEngineStatus.textContent = "Footwear Engine Ready";
    } else if (lowerName.includes('clothing') || lowerName.includes('apparel')) {
        selectCategory.value = 'clothing';
        logInfo("Auto-detected clothing category based on template filename.");
        if (statEngineStatus) statEngineStatus.textContent = "Clothing Engine Ready";
    }
}

function formatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// Log messages inside panel console
function logSystem(msg) { writeLog(msg, 'system-line'); }
function logInfo(msg) { writeLog(msg, 'info-line'); }
function logSuccess(msg) { writeLog(msg, 'success-line'); }
function logWarning(msg) { writeLog(msg, 'warning-line'); }
function logError(msg) { writeLog(msg, 'error-line'); }

function writeLog(text, className) {
    if (!consoleBox) return;
    const line = document.createElement('div');
    line.className = `console-line ${className}`;
    line.textContent = `[${new Date().toLocaleTimeString()}] ${text}`;
    consoleBox.appendChild(line);
    consoleBox.scrollTop = consoleBox.scrollHeight;
}

function clearConsole() {
    if (consoleBox) consoleBox.innerHTML = '';
    logSystem("Console cleared.");
}

// Loading overlay display
function showLoading(text) {
    overlayText.textContent = text;
    overlay.classList.add('active');
}

function hideLoading() {
    overlay.classList.remove('active');
}

// Listing Generation Trigger
async function runListingGeneration() {
    const masterFile = document.getElementById('input-learn-directory').files[0];
    const contentFile = document.getElementById('input-learn-content').files[0];
    const templateFile = document.getElementById('input-list-template').files[0];
    
    if (!masterFile || !contentFile || !templateFile) {
        logError("Listing Error: Select Item Directory, Content Sheet and Template to proceed.");
        alert("Please ensure all three input files are uploaded before generating the listing.");
        return;
    }
    
    const itemCodes = document.getElementById('input-item-codes').value.trim();
    const categorySel = document.getElementById('select-category').value;
    
    showLoading("Running Firstcry Autolister engine...");
    logInfo("Starting listing generation run...");
    
    const formData = new FormData();
    formData.append("item_directory", masterFile);
    formData.append("content_sheet", contentFile);
    formData.append("template", templateFile);
    formData.append("category", categorySel);
    formData.append("mapping_config", "");
    if (itemCodes) {
        formData.append("item_codes", itemCodes);
    }
    
    // Hide old download button if any
    if (downloadContainer) {
        downloadContainer.style.display = 'none';
        downloadContainer.innerHTML = '';
    }
    
    try {
        const response = await fetch('/api/list', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            let errorMsg = "Listing compile run failed";
            try {
                const err = await response.json();
                if (err && err.detail) {
                    if (Array.isArray(err.detail)) {
                        errorMsg = err.detail.map(d => `${d.loc ? d.loc.join('.') : ''}: ${d.msg}`).join(', ');
                    } else if (typeof err.detail === 'string') {
                        errorMsg = err.detail;
                    } else {
                        errorMsg = JSON.stringify(err.detail);
                    }
                }
            } catch (e) {
                // Not JSON or parse issue
            }
            throw new Error(errorMsg);
        }
        
        const blob = await response.blob();
        
        // Generate download
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = templateFile.name || 'completed_listing.xlsm';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        
        logSuccess(`Listing generated successfully!`);
        
        // Show prominent orange download button in the same page container
        if (downloadContainer) {
            downloadContainer.innerHTML = `
                <a href="${url}" download="${templateFile.name || 'completed_listing.xlsm'}" class="btn btn-primary btn-full" style="text-decoration: none; display: flex; align-items: center; justify-content: center; gap: 8px; background-color: var(--primary-orange); box-shadow: 0 4px 14px 0 var(--orange-glow);">
                     Download Completed Listing 📥
                </a>
            `;
            downloadContainer.style.display = 'block';
        }
        
        // Print clickable download link in the console logs
        const logLine = document.createElement('div');
        logLine.className = 'console-line success-line';
        logLine.innerHTML = `[${new Date().toLocaleTimeString()}] 📥 <a href="${url}" download="${templateFile.name || 'completed_listing.xlsm'}" class="download-link" style="color: #34D399; font-weight: bold; text-decoration: underline;">Click Here to Download ${templateFile.name}</a>`;
        consoleBox.appendChild(logLine);
        consoleBox.scrollTop = consoleBox.scrollHeight;
        
        // Update stats
        if (statSkusCount) {
            let count = 0;
            if (itemCodes) {
                count = itemCodes.split('\n').filter(c => c.trim()).length;
            }
            statSkusCount.textContent = count > 0 ? `${count} SKUs` : "All SKUs";
        }
        
        if (statEngineStatus) {
            statEngineStatus.textContent = "Generated Successfully";
        }
        
    } catch (e) {
        logError(`Generation Failed: ${e.message}`);
        alert(`Failed to generate listing: ${e.message}`);
    } finally {
        hideLoading();
    }
}
