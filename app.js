// Power Automate HTTP URL
const POWER_AUTOMATE_URL = 'https://default6dff32de1cd04ada892b2298e1f616.98.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/548d5069dd3f4a25aee9f34b373b2a6a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=DVfvTJ58sD2AFdiaHsNkKuY2o8JAPyW-5gx3CQMods8';

let currentUser = null;

// Get authenticated user from Azure Static Web Apps Easy Auth
async function getCurrentUser() {
    try {
        const response = await fetch('/.auth/me');
        if (!response.ok) {
            throw new Error('Failed to get user info');
        }
        const authData = await response.json();
        
        if (authData && authData.clientPrincipal) {
            currentUser = {
                name: authData.clientPrincipal.userDetails,
                email: authData.clientPrincipal.userDetails,
                id: authData.clientPrincipal.userId,
                identityProvider: authData.clientPrincipal.identityProvider
            };
            return currentUser;
        }
        return null;
    } catch (error) {
        console.error('Error getting user:', error);
        return null;
    }
}

// Initialize app on page load
async function initializeApp() {
    try {
        // Try to get user info, but don't block the form if it fails
        // Since routes are protected, if user reached this page, they're authenticated
        const user = await getCurrentUser();
        if (user) {
            currentUser = user;
        }
        // Always show the form - route protection handles authentication
        document.getElementById('form-container').style.display = 'block';
    } catch (error) {
        console.error('Initialization error:', error);
        // Still show the form even if we can't get user info
        // The route protection ensures only authenticated users can reach here
        document.getElementById('form-container').style.display = 'block';
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
    if (!purchaseValue) {
        throw new Error('Please select a purchase value.');
    }

    if (purchaseValue === 'Reimbursement Request (<$100)') {
        if (fileCount < 1) {
            throw new Error('At least 1 document is required for Reimbursement Requests. Please upload at least one file.');
        }
    } else if (purchaseValue === '<=$1000') {
        if (fileCount < 1) {
            throw new Error('At least 1 quote is required for purchases <=$1000. Please upload at least one file.');
        }
    } else if (purchaseValue === '$1001-$6000') {
        if (fileCount < 2) {
            throw new Error('At least 2 quotations are required for purchases $1001-$6000. Please upload at least two files.');
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

    // Try to get user info if not already available
    if (!currentUser) {
        currentUser = await getCurrentUser();
    }
    
    // If still no user, use placeholder values (route protection ensures user is authenticated)
    const userInfo = currentUser || {
        name: 'Unknown User',
        email: 'unknown@example.com',
        id: 'unknown',
        identityProvider: 'microsoft'
    };

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
                name: userInfo.name,
                email: userInfo.email,
                id: userInfo.id,
                identityProvider: userInfo.identityProvider
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

// Make removeFile available globally
window.removeFile = removeFile;

// Initialize app when DOM is loaded
document.addEventListener('DOMContentLoaded', initializeApp);
