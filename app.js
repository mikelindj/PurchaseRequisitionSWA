// MSAL Configuration
// IMPORTANT: Replace these values with your Azure AD app registration details
const msalConfig = {
    auth: {
        clientId: 'YOUR_CLIENT_ID', // Replace with your Azure AD Application (client) ID
        authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID', // Replace with your tenant ID
        redirectUri: window.location.origin
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
    }
};

// Power Automate HTTP URL
// IMPORTANT: Replace with your actual Power Automate HTTP trigger URL
const POWER_AUTOMATE_URL = 'YOUR_POWER_AUTOMATE_HTTP_URL';

const msalInstance = new msal.PublicClientApplication(msalConfig);
let account = null;

// Initialize MSAL
msalInstance.initialize().then(() => {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        account = accounts[0];
        updateUI();
    }
});

// Login function
async function login() {
    try {
        const loginRequest = {
            scopes: ['User.Read'],
            account: account
        };

        const response = await msalInstance.loginPopup(loginRequest);
        account = response.account;
        updateUI();
    } catch (error) {
        console.error('Login error:', error);
        showError('Failed to sign in. Please try again.');
    }
}

// Logout function
function logout() {
    msalInstance.logoutPopup({
        account: account
    }).then(() => {
        account = null;
        updateUI();
        document.getElementById('purchase-form').reset();
        document.getElementById('file-list').innerHTML = '';
        hideMessages();
    });
}

// Update UI based on authentication state
function updateUI() {
    const loginBtn = document.getElementById('login-btn');
    const logoutBtn = document.getElementById('logout-btn');
    const userInfo = document.getElementById('user-info');
    const formContainer = document.getElementById('form-container');

    if (account) {
        loginBtn.style.display = 'none';
        logoutBtn.style.display = 'inline-block';
        userInfo.style.display = 'inline-block';
        userInfo.textContent = `Signed in as: ${account.name || account.username}`;
        formContainer.style.display = 'block';
    } else {
        loginBtn.style.display = 'inline-block';
        logoutBtn.style.display = 'none';
        userInfo.style.display = 'none';
        formContainer.style.display = 'none';
    }
}

// File handling
const documentsInput = document.getElementById('documents');
const fileList = document.getElementById('file-list');
const selectedFiles = [];

documentsInput.addEventListener('change', function(e) {
    const files = Array.from(e.target.files);
    
    // Validate file count
    if (selectedFiles.length + files.length > 10) {
        showError('Maximum 10 files allowed. Please remove some files first.');
        return;
    }

    // Validate file sizes and types
    const allowedTypes = [
        'application/msword',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-powerpoint',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'application/pdf',
        'image/jpeg',
        'image/jpg',
        'image/png',
        'image/gif',
        'video/mp4',
        'video/quicktime',
        'audio/mpeg',
        'audio/wav'
    ];

    files.forEach(file => {
        if (file.size > 10 * 1024 * 1024) {
            showError(`File "${file.name}" exceeds 10MB size limit.`);
            return;
        }

        if (!allowedTypes.includes(file.type)) {
            showError(`File "${file.name}" is not an allowed file type.`);
            return;
        }

        selectedFiles.push(file);
        displayFile(file);
    });

    // Reset input to allow selecting the same file again
    documentsInput.value = '';
});

function displayFile(file) {
    const fileItem = document.createElement('div');
    fileItem.className = 'file-item';
    fileItem.dataset.fileName = file.name;

    const fileSize = (file.size / 1024 / 1024).toFixed(2);
    
    fileItem.innerHTML = `
        <span>${file.name}</span>
        <span class="file-size">${fileSize} MB</span>
        <button type="button" class="remove-file" onclick="removeFile('${file.name}')">Remove</button>
    `;

    fileList.appendChild(fileItem);
}

function removeFile(fileName) {
    const index = selectedFiles.findIndex(f => f.name === fileName);
    if (index > -1) {
        selectedFiles.splice(index, 1);
    }

    const fileItem = fileList.querySelector(`[data-file-name="${fileName}"]`);
    if (fileItem) {
        fileItem.remove();
    }
}

// Form validation
function validateForm(formData) {
    const purchaseValue = formData.get('purchaseValue');
    const fileCount = selectedFiles.length;

    // Validate file count based on purchase value
    if (purchaseValue === 'Reimbursement Request (<$100)') {
        if (fileCount < 1) {
            throw new Error('At least 1 document is required for Reimbursement Requests.');
        }
    } else if (purchaseValue === '<=$1000') {
        if (fileCount < 1) {
            throw new Error('At least 1 quote is required for purchases <=$1000.');
        }
    } else if (purchaseValue === '$1001-$6000') {
        if (fileCount < 2) {
            throw new Error('At least 2 quotations are required for purchases $1001-$6000.');
        }
    }

    // Validate total cost
    const totalCost = parseFloat(formData.get('totalCost'));
    if (isNaN(totalCost) || totalCost < 0 || totalCost >= 6001) {
        throw new Error('Total cost must be a number less than 6001.');
    }

    return true;
}

// Convert files to base64
async function filesToBase64(files) {
    const filePromises = files.map(file => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                resolve({
                    name: file.name,
                    type: file.type,
                    size: file.size,
                    content: reader.result.split(',')[1] // Remove data:type;base64, prefix
                });
            };
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    });

    return Promise.all(filePromises);
}

// Form submission
document.getElementById('purchase-form').addEventListener('submit', async function(e) {
    e.preventDefault();

    if (!account) {
        showError('Please sign in to submit the form.');
        return;
    }

    const submitBtn = document.getElementById('submit-btn');
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<span class="loading"></span>Submitting...';

    hideMessages();

    try {
        const formData = new FormData(e.target);
        
        // Validate form
        validateForm(formData);

        // Convert files to base64
        const fileData = await filesToBase64(selectedFiles);

        // Prepare submission data
        const submissionData = {
            timestamp: new Date().toISOString(),
            user: {
                name: account.name,
                email: account.username,
                id: account.homeAccountId
            },
            purchaseValue: formData.get('purchaseValue'),
            budget: formData.get('budget'),
            items: formData.get('items'),
            totalCost: parseFloat(formData.get('totalCost')),
            remarks: formData.get('remarks') || '',
            documents: fileData
        };

        // Send to Power Automate
        const response = await fetch(POWER_AUTOMATE_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(submissionData)
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        // Show success message
        showSuccess();
        document.getElementById('purchase-form').reset();
        selectedFiles.length = 0;
        fileList.innerHTML = '';

    } catch (error) {
        console.error('Submission error:', error);
        showError(error.message || 'Failed to submit the form. Please try again.');
    } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = 'Submit';
    }
});

// Message display functions
function showSuccess() {
    document.getElementById('success-message').style.display = 'block';
    document.getElementById('error-message').style.display = 'none';
    document.getElementById('form-container').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function showError(message) {
    document.getElementById('error-text').textContent = message;
    document.getElementById('error-message').style.display = 'block';
    document.getElementById('success-message').style.display = 'none';
    document.getElementById('form-container').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function hideMessages() {
    document.getElementById('success-message').style.display = 'none';
    document.getElementById('error-message').style.display = 'none';
}

// Event listeners
document.getElementById('login-btn').addEventListener('click', login);
document.getElementById('logout-btn').addEventListener('click', logout);

// Make removeFile available globally
window.removeFile = removeFile;

