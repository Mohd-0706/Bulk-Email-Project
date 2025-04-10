document.addEventListener('DOMContentLoaded', function() {
    // DOM Elements
    const setupSteps = document.querySelectorAll('.setup-step');
    const nextButtons = document.querySelectorAll('.next-btn');
    const backButtons = document.querySelectorAll('.back-btn');
    const uploadArea = document.getElementById('uploadArea');
    const excelFileInput = document.getElementById('excelFile');
    const fileInfo = document.getElementById('fileInfo');
    const analyzeBtn = document.getElementById('analyzeBtn');
    const sendEmailsBtn = document.getElementById('sendEmailsBtn');
    const progressPanel = document.getElementById('progressPanel');
    const resultPanel = document.getElementById('resultPanel');
    const progressBar = document.getElementById('sendProgress');
    const sentCount = document.getElementById('sentCount');
    const totalCount = document.getElementById('totalCount');
    const progressLog = document.getElementById('progressLog');
    const successCount = document.getElementById('successCount');
    const failedCount = document.getElementById('failedCount');
    const timeTaken = document.getElementById('timeTaken');
    const restartBtn = document.getElementById('restartBtn');
    const showGuideBtn = document.getElementById('show-guide');
    const guideModal = new bootstrap.Modal(document.getElementById('guideModal'));
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    const emailEditor = document.getElementById('emailEditor');
    const emailPreview = document.getElementById('emailPreview');
    const variablesList = document.getElementById('variablesList');
    const togglePassword = document.querySelector('.toggle-password');
    const passwordInput = document.getElementById('smtp-password');
    
    // Editor Elements
    const textColorBtn = document.getElementById('textColorBtn');
    const colorPickerModal = new bootstrap.Modal(document.getElementById('colorPickerModal'));
    const colorOptions = document.querySelectorAll('.color-option');
    const customColorInput = document.getElementById('customColor');
    const applyColorBtn = document.getElementById('applyColorBtn');
    const formatButtons = document.querySelectorAll('.format-btn');
    
    // Variables
    let currentStep = 1;
    let excelData = null;
    let columns = [];
    let startTime = null;
    let currentColor = '#00ffcc';
    
    // Initialize
    showStep(currentStep);
    
    // Event Listeners
    nextButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const nextStep = parseInt(btn.dataset.next);
            if (validateStep(currentStep)) {
                currentStep = nextStep;
                showStep(currentStep);
                
                if (nextStep === 3 && columns.length > 0) {
                    addColumnVariables(columns);
                }
            }
        });
    });
    
    backButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const prevStep = parseInt(btn.dataset.prev);
            currentStep = prevStep;
            showStep(currentStep);
        });
    });
    
    // File upload handling
    uploadArea.addEventListener('click', () => {
        excelFileInput.click();
    });
    
    excelFileInput.addEventListener('change', handleFileSelect);
    
    // Drag and drop for file upload
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = 'var(--cyber-primary)';
        uploadArea.style.backgroundColor = 'rgba(0, 255, 204, 0.2)';
    });
    
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.style.borderColor = 'rgba(0, 255, 204, 0.3)';
        uploadArea.style.backgroundColor = 'rgba(0, 255, 204, 0.05)';
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = 'rgba(0, 255, 204, 0.3)';
        uploadArea.style.backgroundColor = 'rgba(0, 255, 204, 0.05)';
        
        if (e.dataTransfer.files.length) {
            excelFileInput.files = e.dataTransfer.files;
            handleFileSelect({ target: excelFileInput });
        }
    });
    
    // Analyze button click
    analyzeBtn.addEventListener('click', () => {
        if (excelData) {
            currentStep = 3;
            showStep(currentStep);
            addColumnVariables(columns);
        }
    });
    
    // Send emails button
    sendEmailsBtn.addEventListener('click', () => {
        if (validateStep(3)) {
            startSendingEmails();
        }
    });
    
    // Restart process
    restartBtn.addEventListener('click', () => {
        resetForm();
        currentStep = 1;
        showStep(currentStep);
    });
    
    // Show guide modal
    showGuideBtn.addEventListener('click', (e) => {
        e.preventDefault();
        guideModal.show();
    });
    
    // Tab switching
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const tabName = btn.dataset.tab;
            switchTab(tabName);
        });
    });
    
    // Format buttons (bold, italic, undo, redo)
    formatButtons.forEach(button => {
        button.addEventListener('click', (e) => {
            const command = e.currentTarget.dataset.command;
            document.execCommand(command, false, null);
            
            // Update button states
            updateButtonStates();
            
            emailEditor.focus();
        });
    });
    
    // Update button states when selection changes
    emailEditor.addEventListener('input', updateButtonStates);
    emailEditor.addEventListener('mouseup', updateButtonStates);
    emailEditor.addEventListener('keyup', updateButtonStates);
    
    // Text color button
    textColorBtn.addEventListener('click', () => {
        colorPickerModal.show();
    });
    
    // Color selection
    colorOptions.forEach(option => {
        option.addEventListener('click', () => {
            customColorInput.value = option.dataset.color;
        });
    });
    
    // Apply color
    applyColorBtn.addEventListener('click', () => {
        currentColor = customColorInput.value;
        document.execCommand('foreColor', false, currentColor);
        colorPickerModal.hide();
        emailEditor.focus();
    });
    
    // Make variables draggable
    variablesList.addEventListener('dragstart', (e) => {
        if (e.target.classList.contains('variable-item')) {
            e.dataTransfer.setData('text/plain', e.target.dataset.var);
            e.dataTransfer.effectAllowed = 'copy';
        }
    });
    
    // Allow dropping variables into editor
    emailEditor.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
    });
    
    emailEditor.addEventListener('drop', (e) => {
        e.preventDefault();
        const variable = e.dataTransfer.getData('text/plain');
        if (variable) {
            const selection = window.getSelection();
            const range = selection.getRangeAt(0);
            range.deleteContents();
            range.insertNode(document.createTextNode(variable));
            emailEditor.focus();
        }
    });
    
    // Toggle password visibility
    togglePassword.addEventListener('click', () => {
        const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
        passwordInput.setAttribute('type', type);
        togglePassword.classList.toggle('fa-eye-slash');
        togglePassword.classList.toggle('fa-eye');
    });
    
    // Functions
    function showStep(step) {
        setupSteps.forEach(stepEl => {
            stepEl.classList.remove('active');
            if (parseInt(stepEl.dataset.step) === step) {
                stepEl.classList.add('active');
            }
        });
        
        progressPanel.style.display = 'none';
        resultPanel.style.display = 'none';
    }
    
    function validateStep(step) {
        if (step === 1) {
            const email = document.getElementById('smtp-email').value;
            const password = document.getElementById('smtp-password').value;
            
            if (!email || !password) {
                alert('Please enter both email and password');
                return false;
            }
            
            if (!email.includes('@')) {
                alert('Please enter a valid email address');
                return false;
            }
            
            return true;
        }
        
        if (step === 2) {
            if (!excelData || excelData.length === 0) {
                alert('Please upload an Excel file with recipient data');
                return false;
            }
            
            return true;
        }
        
        if (step === 3) {
            const subject = document.getElementById('email-subject').value;
            const content = emailEditor.innerHTML;
            
            if (!subject) {
                alert('Please enter an email subject');
                return false;
            }
            
            if (!content || content === '<br>' || content === '<div><br></div>') {
                alert('Please compose your email content');
                return false;
            }
            
            return true;
        }
        
        return true;
    }
    
    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            alert('Please upload an Excel file (.xlsx or .xls)');
            return;
        }
        
        fileInfo.innerHTML = `
            <i class="fas fa-check-circle" style="color: var(--cyber-success);"></i>
            ${file.name} (${formatFileSize(file.size)})
            <p><i class="fas fa-spinner fa-spin"></i> Analyzing file...</p>
        `;
        fileInfo.classList.add('show');
        
        const formData = new FormData();
        formData.append('file', file);
        
        fetch('/analyze_excel', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! Status: ${response.status}`);
            }
            return response.json();
        })
        .then(data => {
            if (data.status === 'success') {
                excelData = data.preview;
                columns = data.columns;
                totalCount.textContent = data.rowCount;
                
                fileInfo.innerHTML = `
                    <i class="fas fa-check-circle" style="color: var(--cyber-success);"></i>
                    ${file.name} (${formatFileSize(file.size)})
                    <p>Successfully analyzed ${data.rowCount} rows with ${columns.length} columns</p>
                `;
                
                analyzeBtn.disabled = false;
            } else {
                throw new Error(data.message || 'Unknown error occurred');
            }
        })
        .catch(error => {
            console.error('Error analyzing Excel file:', error);
            fileInfo.innerHTML = `
                <i class="fas fa-times-circle" style="color: var(--cyber-danger);"></i>
                ${file.name} (${formatFileSize(file.size)})
                <p style="color: var(--cyber-danger);">Error: ${error.message}</p>
            `;
            alert('Error analyzing Excel file: ' + error.message);
        });
    }
    
    function startSendingEmails() {
        startTime = new Date();
        progressPanel.style.display = 'block';
        progressBar.style.width = '0%';
        sentCount.textContent = '0';
        totalCount.textContent = excelData.length;
        progressLog.innerHTML = '';
        
        const formData = new FormData();
        formData.append('smtp_email', document.getElementById('smtp-email').value);
        formData.append('smtp_password', document.getElementById('smtp-password').value);
        formData.append('subject', document.getElementById('email-subject').value);
        formData.append('template', emailEditor.innerHTML);
        
        if (excelFileInput.files.length > 0) {
            formData.append('excel_file', excelFileInput.files[0]);
        }
        
        fetch('/send_emails', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            if (data.status === 'success') {
                progressBar.style.width = '100%';
                sentCount.textContent = data.sentCount;
                progressLog.innerHTML += `<p>Successfully sent ${data.sentCount} out of ${data.totalCount} emails</p>`;
                
                setTimeout(() => {
                    showResults(data);
                }, 1000);
            } else {
                throw new Error(data.message || 'Unknown error occurred');
            }
        })
        .catch(error => {
            progressLog.innerHTML += `<p class="error">Error: ${error.message}</p>`;
            alert('Error sending emails: ' + error.message);
        });
    }
    
    function showResults(data) {
        progressPanel.style.display = 'none';
        resultPanel.style.display = 'block';
        
        const endTime = new Date();
        const duration = Math.round((endTime - startTime) / 1000);
        
        successCount.textContent = data.sentCount;
        failedCount.textContent = data.totalCount - data.sentCount;
        timeTaken.textContent = `${duration} seconds`;
    }
    
    function resetForm() {
        document.getElementById('smtp-email').value = '';
        document.getElementById('smtp-password').value = '';
        document.getElementById('email-subject').value = '';
        emailEditor.innerHTML = '';
        emailPreview.innerHTML = '';
        fileInfo.innerHTML = '';
        fileInfo.classList.remove('show');
        excelFileInput.value = '';
        excelData = null;
        columns = [];
        analyzeBtn.disabled = true;
        progressPanel.style.display = 'none';
        resultPanel.style.display = 'none';
    }
    
    function switchTab(tabName) {
        tabBtns.forEach(btn => {
            btn.classList.remove('active');
            if (btn.dataset.tab === tabName) {
                btn.classList.add('active');
            }
        });
    
        tabContents.forEach(content => {
            content.style.display = content.dataset.tab === tabName ? 'block' : 'none';
        });
    
        if (tabName === 'preview') {
            emailPreview.innerHTML = emailEditor.innerHTML;
        }
    }
    
    function addColumnVariables(columns) {
        variablesList.innerHTML = '';
        columns.forEach(column => {
            const varItem = document.createElement('div');
            varItem.className = 'variable-item';
            varItem.dataset.var = `{${column}}`;
            varItem.draggable = true;
            varItem.innerHTML = `
                <i class="fas fa-tag"></i>
                <span>${column}</span>
            `;
            variablesList.appendChild(varItem);
        });
    }
    
    function updateButtonStates() {
        formatButtons.forEach(button => {
            const command = button.dataset.command;
            if (command === 'bold' || command === 'italic') {
                const isActive = document.queryCommandState(command);
                button.classList.toggle('active', isActive);
            }
        });
    }
    
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
});