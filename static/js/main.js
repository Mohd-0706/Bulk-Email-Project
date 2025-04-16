document.addEventListener('DOMContentLoaded', function () {
    // DOM Elements
    const elements = {
        setupSteps: document.querySelectorAll('.setup-step'),
        nextButtons: document.querySelectorAll('.next-btn'),
        backButtons: document.querySelectorAll('.back-btn'),
        uploadArea: document.getElementById('uploadArea'),
        excelFileInput: document.getElementById('excelFile'),
        fileInfo: document.getElementById('fileInfo'),
        analyzeBtn: document.getElementById('analyzeBtn'),
        sendEmailsBtn: document.getElementById('sendEmailsBtn'),
        progressPanel: document.getElementById('progressPanel'),
        resultPanel: document.getElementById('resultPanel'),
        progressBar: document.getElementById('sendProgress'),
        sentCount: document.getElementById('sentCount'),
        totalCount: document.getElementById('totalCount'),
        progressLog: document.getElementById('progressLog'),
        successCount: document.getElementById('successCount'),
        failedCount: document.getElementById('failedCount'),
        timeTaken: document.getElementById('timeTaken'),
        restartBtn: document.getElementById('restartBtn'),
        showGuideBtn: document.getElementById('show-guide'),
        guideModal: new bootstrap.Modal(document.getElementById('guideModal')),
        emailEditor: document.getElementById('emailEditor'),
        htmlEditor: document.getElementById('htmlEditor'),
        updateVisualBtn: document.getElementById('updateVisualBtn'),
        variablesList: document.getElementById('variablesList'),
        togglePassword: document.querySelector('.toggle-password'),
        passwordInput: document.getElementById('smtp-password'),
        formatButtons: document.querySelectorAll('.format-btn'),
        emailSubject: document.getElementById('email-subject'),
        smtpEmail: document.getElementById('smtp-email'),
        smtpPassword: document.getElementById('smtp-password'),
        downloadSampleBtn: document.getElementById('downloadSampleBtn'),
        editExcelBtn: document.getElementById('editExcelBtn'),
        excelEditorModal: new bootstrap.Modal(document.getElementById('excelEditorModal')),
        saveExcelChangesBtn: document.getElementById('saveExcelChangesBtn'),
        tableHeader: document.getElementById('tableHeader'),
        tableBody: document.getElementById('tableBody'),
        themeSwitch: document.getElementById('checkbox'),
        addAttachmentBtn: document.getElementById('addAttachmentBtn'),
        attachmentInput: document.getElementById('attachmentInput'),
        attachmentsList: document.getElementById('attachmentsList'),
        creditBtn: document.getElementById('creditBtn'),
        creditModal: new bootstrap.Modal(document.getElementById('creditModal')),
        addRowBtn: document.getElementById('addRowBtn'),
        removeRowBtn: document.getElementById('removeRowBtn'),
        addColumnBtn: document.getElementById('addColumnBtn'),
        removeColumnBtn: document.getElementById('removeColumnBtn')
    };

    // State management
    const state = {
        currentStep: 1,
        excelData: null,
        columns: [],
        totalEmails: 0,
        sentEmails: 0,
        failedEmails: 0,
        attachments: [],
        excelFile: null,
        startTime: null,
        editedExcelData: null,
        sendingInProgress: false,
        currentRowCount: 0,
        currentColumnCount: 0
    };

    // Initialize the application
    init();

    function init() {
        setupEventListeners();
        validateStep1();
        checkThemePreference();
        setupExcelTableEditing();

        // Set initial focus
        if (elements.smtpEmail) {
            elements.smtpEmail.focus();
        }
    }

    function setupEventListeners() {
        // Navigation buttons
        elements.nextButtons.forEach(btn => {
            btn.addEventListener('click', handleNextStep);
        });

        elements.backButtons.forEach(btn => {
            btn.addEventListener('click', handlePreviousStep);
        });

        // Step 1 validation
        elements.smtpEmail.addEventListener('input', validateStep1);
        elements.smtpPassword.addEventListener('input', validateStep1);
        elements.togglePassword.addEventListener('click', togglePasswordVisibility);
        elements.showGuideBtn.addEventListener('click', showGuideModal);
        elements.creditBtn.addEventListener('click', showCreditModal);

        // File upload handling
        elements.uploadArea.addEventListener('click', () => {
            elements.excelFileInput.value = '';
            elements.excelFileInput.click();
        });

        elements.excelFileInput.addEventListener('change', handleFileSelect);

        // Editor functionality
        elements.updateVisualBtn.addEventListener('click', updateVisualEditor);
        elements.emailEditor.addEventListener('input', syncEditors);
        elements.emailEditor.addEventListener('mouseup', updateFormatButtonStates);
        elements.emailEditor.addEventListener('keyup', updateFormatButtonStates);

        // Format buttons
        elements.formatButtons.forEach(button => {
            button.addEventListener('click', handleFormatButtonClick);
        });

        // Sample Excel download
        elements.downloadSampleBtn.addEventListener('click', downloadSampleExcel);
        elements.editExcelBtn.addEventListener('click', openExcelEditor);
        elements.saveExcelChangesBtn.addEventListener('click', saveExcelChanges);

        // Theme switch
        elements.themeSwitch.addEventListener('change', toggleTheme);

        // Attachments
        elements.addAttachmentBtn.addEventListener('click', () => elements.attachmentInput.click());
        elements.attachmentInput.addEventListener('change', handleAttachmentUpload);
        elements.attachmentsList.addEventListener('click', handleAttachmentRemoval);

        // Other interactions
        elements.sendEmailsBtn.addEventListener('click', startSendingEmails);
        elements.restartBtn.addEventListener('click', resetForm);
        elements.variablesList.addEventListener('click', handleVariableInsertion);

        // Excel table controls
        if (elements.addRowBtn) elements.addRowBtn.addEventListener('click', addTableRow);
        if (elements.removeRowBtn) elements.removeRowBtn.addEventListener('click', removeTableRow);
        if (elements.addColumnBtn) elements.addColumnBtn.addEventListener('click', addTableColumn);
        if (elements.removeColumnBtn) elements.removeColumnBtn.addEventListener('click', removeTableColumn);
    }

    function checkThemePreference() {
        const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        const savedTheme = localStorage.getItem('theme');

        if (savedTheme) {
            setTheme(savedTheme);
        } else if (prefersDark) {
            setTheme('dark');
        }
    }

    function toggleTheme() {
        const newTheme = elements.themeSwitch.checked ? 'dark' : 'light';
        setTheme(newTheme);
    }

    function setTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem('theme', theme);
        elements.themeSwitch.checked = theme === 'dark';
    }

    function showStep(step) {
        elements.setupSteps.forEach(stepEl => {
            stepEl.classList.toggle('active', parseInt(stepEl.dataset.step) === step);
        });

        elements.progressPanel.style.display = 'none';
        elements.resultPanel.style.display = 'none';
    }

    function handleNextStep() {
        const nextStep = parseInt(this.dataset.next);
        if (validateStep(state.currentStep)) {
            state.currentStep = nextStep;
            showStep(state.currentStep);

            if (nextStep === 3 && state.columns.length > 0) {
                addColumnVariables(state.columns);
            }
        }
    }

    function handlePreviousStep() {
        const prevStep = parseInt(this.dataset.prev);
        state.currentStep = prevStep;
        showStep(state.currentStep);
    }

    function validateStep1() {
        const emailValid = elements.smtpEmail.value.includes('@');
        const passwordValid = elements.smtpPassword.value.trim().length > 0;
        const continueBtn = document.querySelector('.setup-step[data-step="1"] .next-btn');

        continueBtn.disabled = !(emailValid && passwordValid);
        return continueBtn.disabled === false;
    }

    function validateStep(step) {
        switch (step) {
            case 1:
                if (!validateStep1()) {
                    showToast('Please enter valid email and password', 'error');
                    return false;
                }
                return true;

            case 2:
                if (!state.excelData || state.excelData.length === 0) {
                    showToast('Please upload an Excel file with recipient data', 'error');
                    return false;
                }
                return true;

            case 3:
                if (!elements.emailSubject.value) {
                    showToast('Please enter an email subject', 'error');
                    return false;
                }

                const content = elements.emailEditor.innerHTML;
                if (!content || content === '<br>' || content === '<div><br></div>') {
                    showToast('Please compose your email content', 'error');
                    return false;
                }
                return true;

            default:
                return true;
        }
    }

    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            showToast('Please upload an Excel file (.xlsx or .xls)', 'error');
            elements.excelFileInput.value = '';
            return;
        }

        showFileInfo(file, 'analyzing');
        state.excelFile = file;

        const formData = new FormData();
        formData.append('file', file);

        fetch('/analyze_excel', {
            method: 'POST',
            body: formData
        })
            .then(handleResponse)
            .then(data => {
                if (data.status === 'success') {
                    state.excelData = data.preview;
                    state.columns = data.columns;
                    state.totalEmails = data.rowCount;
                    elements.totalCount.textContent = state.totalEmails;
                    showFileInfo(file, 'success', data);
                    elements.analyzeBtn.disabled = false;
                    elements.editExcelBtn.disabled = false;
                    showToast('File analyzed successfully!', 'success');
                } else {
                    throw new Error(data.message || 'Unknown error occurred');
                }
            })
            .catch(error => {
                console.error('Error analyzing Excel file:', error);
                showFileInfo(file, 'error', error);
                elements.excelFileInput.value = '';
                showToast(error.message, 'error');
            });
    }

    function showFileInfo(file, status, data = null) {
        const icon = elements.fileInfo.querySelector('.file-info-icon');
        const text = elements.fileInfo.querySelector('.file-info-text');

        switch (status) {
            case 'analyzing':
                icon.innerHTML = '<i class="fas fa-spinner fa-spin"></i>';
                text.innerHTML = `
                    <span>${file.name} (${formatFileSize(file.size)})</span>
                    <p>Analyzing file...</p>
                `;
                break;
            case 'success':
                icon.innerHTML = '<i class="fas fa-check-circle" style="color: var(--success);"></i>';
                text.innerHTML = `
                    <span>${file.name} (${formatFileSize(file.size)})</span>
                    <p>Successfully analyzed ${data.rowCount} rows with ${data.columns.length} columns</p>
                `;
                break;
            case 'error':
                icon.innerHTML = '<i class="fas fa-times-circle" style="color: var(--danger);"></i>';
                text.innerHTML = `
                    <span>${file.name} (${formatFileSize(file.size)})</span>
                    <p style="color: var(--danger);">Error: ${data.message}</p>
                `;
                break;
        }

        elements.fileInfo.classList.add('show');
    }

    function downloadSampleExcel() {
        window.location.href = '/download_sample';
    }

    function openExcelEditor() {
        if (!state.excelFile) return;

        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            state.currentRowCount = jsonData.length - 1; // Subtract header row
            state.currentColumnCount = jsonData[0] ? jsonData[0].length : 0;

            renderExcelTable(jsonData);
            elements.excelEditorModal.show();
        };
        reader.readAsArrayBuffer(state.excelFile);
    }

    function renderExcelTable(data) {
        elements.tableHeader.innerHTML = '';
        elements.tableBody.innerHTML = '';

        if (data.length === 0) return;

        // Create header row
        const headers = data[0];
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            elements.tableHeader.appendChild(th);
        });

        // Create data rows
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const tr = document.createElement('tr');

            headers.forEach((header, index) => {
                const td = document.createElement('td');
                td.contentEditable = true;
                td.textContent = row[index] || '';
                tr.appendChild(td);
            });

            elements.tableBody.appendChild(tr);
        }
    }

    function addTableRow() {
        if (state.currentRowCount >= 50) {
            showToast('Maximum 50 rows allowed', 'warning');
            return;
        }

        const newRow = document.createElement('tr');
        const headers = Array.from(elements.tableHeader.querySelectorAll('th'));
        
        headers.forEach(header => {
            const td = document.createElement('td');
            td.contentEditable = true;
            newRow.appendChild(td);
        });

        elements.tableBody.appendChild(newRow);
        state.currentRowCount++;
    }

    function removeTableRow() {
        if (state.currentRowCount <= 1) {
            showToast('Must keep at least one row', 'warning');
            return;
        }
        
        const rows = elements.tableBody.querySelectorAll('tr');
        if (rows.length > 0) {
            rows[rows.length - 1].remove();
            state.currentRowCount--;
        }
    }

    function addTableColumn() {
        if (state.currentColumnCount >= 50) {
            showToast('Maximum 50 columns allowed', 'warning');
            return;
        }

        // Add header
        const newHeader = document.createElement('th');
        newHeader.contentEditable = true;
        newHeader.textContent = `Column ${state.currentColumnCount + 1}`;
        elements.tableHeader.appendChild(newHeader);

        // Add cells to each row
        const rows = Array.from(elements.tableBody.querySelectorAll('tr'));
        rows.forEach(row => {
            const td = document.createElement('td');
            td.contentEditable = true;
            row.appendChild(td);
        });

        state.currentColumnCount++;
    }

    function removeTableColumn() {
        if (state.currentColumnCount <= 1) {
            showToast('Must keep at least one column', 'warning');
            return;
        }
        
        // Remove header
        const headers = elements.tableHeader.querySelectorAll('th');
        if (headers.length > 0) {
            headers[headers.length - 1].remove();
        }
        
        // Remove cells from each row
        const rows = elements.tableBody.querySelectorAll('tr');
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            if (cells.length > 0) {
                cells[cells.length - 1].remove();
            }
        });
        
        state.currentColumnCount--;
    }

    function saveExcelChanges() {
        const headers = Array.from(elements.tableHeader.querySelectorAll('th')).map(th => th.textContent);
        const rows = Array.from(elements.tableBody.querySelectorAll('tr')).map(tr => {
            const rowData = Array.from(tr.querySelectorAll('td')).map(td => td.textContent);
            // Filter out empty rows (where all cells are empty)
            const isEmptyRow = rowData.every(cell => cell.trim() === '');
            return isEmptyRow ? null : rowData;
        }).filter(row => row !== null); // Remove null rows (empty rows)

        // Store the edited data in state
        state.editedExcelData = {
            headers: headers,
            rows: rows
        };

        // Create a new Excel file from the edited data
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
        XLSX.utils.book_append_sheet(wb, ws, "Recipients");

        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const fileName = state.excelFile ? state.excelFile.name : 'recipients.xlsx';
        
        // Update the state with the new file
        state.excelFile = new File([blob], fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Update the preview data in state
        state.excelData = rows.map(row => {
            const obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index] || '';
            });
            return obj;
        });

        // Update the UI
        state.totalEmails = rows.length;
        elements.totalCount.textContent = rows.length;
        showFileInfo(state.excelFile, 'success', {
            rowCount: rows.length,
            columns: headers
        });
        
        elements.excelEditorModal.hide();
        showToast('Changes saved successfully!', 'success');
    }

    function setupExcelTableEditing() {
        elements.tableBody.addEventListener('input', function (e) {
            if (e.target.tagName === 'TD') {
                e.target.style.height = 'auto';
                e.target.style.height = (e.target.scrollHeight) + 'px';
            }
        });
    }

    function handleAttachmentUpload(event) {
        const files = Array.from(event.target.files);

        const validFiles = files.filter(file => {
            if (file.size > 5 * 1024 * 1024) {
                showToast(`File ${file.name} exceeds 5MB limit and will not be attached`, 'warning');
                return false;
            }
            return true;
        });

        state.attachments = [...state.attachments, ...validFiles];
        renderAttachments();

        event.target.value = '';
    }

    function renderAttachments() {
        elements.attachmentsList.innerHTML = '';

        state.attachments.forEach((file, index) => {
            const attachmentItem = document.createElement('div');
            attachmentItem.className = 'attachment-item';
            attachmentItem.innerHTML = `
                <div class="attachment-info">
                    <i class="fas fa-paperclip attachment-icon"></i>
                    <span class="attachment-name">${file.name}</span>
                    <span class="attachment-size">${formatFileSize(file.size)}</span>
                </div>
                <i class="fas fa-times remove-attachment" data-index="${index}"></i>
            `;
            elements.attachmentsList.appendChild(attachmentItem);
        });
    }

    function handleAttachmentRemoval(event) {
        if (event.target.classList.contains('remove-attachment')) {
            const index = parseInt(event.target.dataset.index);
            state.attachments.splice(index, 1);
            renderAttachments();
        }
    }

    async function startSendingEmails() {
        if (state.sendingInProgress) return;
        if (!validateStep(3)) return;

        // Disable send button during process
        elements.sendEmailsBtn.disabled = true;
        state.sendingInProgress = true;

        state.startTime = new Date();
        state.sentEmails = 0;
        state.failedEmails = 0;
        showProgressPanel();

        // Add loading effect
        elements.sendEmailsBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Sending...';

        try {
            const formData = createFormData();
            state.attachments.forEach((file, index) => {
                formData.append(`attachment_${index}`, file);
            });

            const response = await fetch('/send_emails', {
                method: 'POST',
                body: formData
            });

            const data = await handleResponse(response);

            if (data.status === 'success') {
                state.sentEmails = data.sentCount;
                state.failedEmails = data.totalCount - data.sentCount;
                showResults();
                showToast(`Successfully sent ${state.sentEmails} emails!`, 'success');
            } else {
                throw new Error(data.message || 'Error sending emails');
            }
        } catch (error) {
            console.error('Error sending emails:', error);
            elements.progressLog.innerHTML += `<p class="error">Error: ${error.message}</p>`;
            showToast(`Error: ${error.message}`, 'error');
        } finally {
            elements.sendEmailsBtn.disabled = false;
            state.sendingInProgress = false;
            elements.sendEmailsBtn.innerHTML = '<i class="fas fa-paper-plane"></i> Send Emails';
        }
    }

    function createFormData() {
        const formData = new FormData();
        formData.append('smtp_email', elements.smtpEmail.value);
        formData.append('smtp_password', elements.smtpPassword.value);
        formData.append('subject', elements.emailSubject.value);
        formData.append('template', elements.emailEditor.innerHTML);

        // Use the edited Excel file if available, otherwise use the original
        if (state.excelFile) {
            formData.append('excel_file', state.excelFile);
        }

        return formData;
    }

    function showProgressPanel() {
        elements.progressPanel.style.display = 'block';
        elements.progressBar.style.width = '0%';
        elements.sentCount.textContent = '0';
        elements.totalCount.textContent = state.totalEmails;
        elements.progressLog.innerHTML = '';
    }

    function updateProgress(percent, sentCount) {
        elements.progressBar.style.width = `${percent}%`;
        elements.sentCount.textContent = sentCount;
    }

    function showResults() {
        elements.progressPanel.style.display = 'none';
        elements.resultPanel.style.display = 'block';

        const endTime = new Date();
        const duration = Math.round((endTime - state.startTime) / 1000);

        elements.successCount.textContent = state.sentEmails;
        elements.failedCount.textContent = state.failedEmails;
        elements.timeTaken.textContent = duration;

        document.querySelector('.success-animation').classList.add('pulse');
    }

    function updateVisualEditor() {
        elements.emailEditor.innerHTML = elements.htmlEditor.value;
    }

    function syncEditors() {
        elements.htmlEditor.value = elements.emailEditor.innerHTML;
    }

    function handleFormatButtonClick() {
        const command = this.dataset.command;
        document.execCommand(command, false, null);
        elements.emailEditor.focus();
        updateFormatButtonStates();
    }

    function updateFormatButtonStates() {
        elements.formatButtons.forEach(button => {
            const command = button.dataset.command;
            if (['bold', 'italic', 'underline'].includes(command)) {
                button.classList.toggle('active', document.queryCommandState(command));
            }
        });
    }

    function handleVariableInsertion(e) {
        const varItem = e.target.closest('.variable-item');
        if (!varItem) return;

        const variable = `{${varItem.textContent.trim()}}`;
        const isHtmlTabActive = document.querySelector('#html-tab').classList.contains('active');

        if (isHtmlTabActive) {
            // For HTML editor - insert at cursor position
            const htmlEditor = elements.htmlEditor;
            const startPos = htmlEditor.selectionStart;
            const endPos = htmlEditor.selectionEnd;
            const currentValue = htmlEditor.value;
            
            // Insert the variable at cursor position
            htmlEditor.value = currentValue.substring(0, startPos) + variable + currentValue.substring(endPos);
            
            // Set cursor position after inserted variable
            htmlEditor.selectionStart = htmlEditor.selectionEnd = startPos + variable.length;
            htmlEditor.focus();
        } else {
            // For visual editor - insert at cursor position
            elements.emailEditor.focus();
            
            // Check if there's a selection
            const selection = window.getSelection();
            
            if (selection.rangeCount > 0) {
                const range = selection.getRangeAt(0);
                range.deleteContents();
                
                // Create a text node for our variable
                const textNode = document.createTextNode(variable);
                range.insertNode(textNode);
                
                // Move cursor to after the inserted variable
                range.setStartAfter(textNode);
                range.collapse(true);
                selection.removeAllRanges();
                selection.addRange(range);
            } else {
                // If no selection, insert at current cursor position
                document.execCommand('insertText', false, variable);
            }
        }
        
        // Sync editors
        syncEditors();
    }

    function addColumnVariables(columns) {
        elements.variablesList.innerHTML = '';
        columns.forEach(column => {
            const varItem = document.createElement('div');
            varItem.className = 'variable-item';
            varItem.innerHTML = `
                <i class="fas fa-tag"></i>
                <span>${column}</span>
            `;
            elements.variablesList.appendChild(varItem);
        });
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    function handleResponse(response) {
        if (!response.ok) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }
        return response.json();
    }

    function showGuideModal(e) {
        e.preventDefault();
        elements.guideModal.show();
    }

    function showCreditModal() {
        elements.creditModal.show();
    }

    function togglePasswordVisibility() {
        const type = elements.passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
        elements.passwordInput.setAttribute('type', type);
        this.classList.toggle('fa-eye-slash');
        this.classList.toggle('fa-eye');
    }

    function resetForm() {
        // Reload the page to start fresh
        window.location.reload();
    }

    function showToast(message, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast-notification toast-${type}`;
        toast.innerHTML = `
            <div class="toast-message">${message}</div>
            <button class="toast-close">&times;</button>
        `;

        document.body.appendChild(toast);

        setTimeout(() => {
            toast.classList.add('fade-out');
            setTimeout(() => toast.remove(), 300);
        }, 5000);

        toast.querySelector('.toast-close').addEventListener('click', () => {
            toast.classList.add('fade-out');
            setTimeout(() => toast.remove(), 300);
        });
    }
});