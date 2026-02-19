// File input handling
const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');
const uploadArea = document.getElementById('uploadArea');
const uploadForm = document.getElementById('uploadForm');
const generateBtn = document.getElementById('generateBtn');

// Display selected file name
fileInput.addEventListener('change', function(e) {
    if (this.files && this.files[0]) {
        fileName.textContent = '✓ Selected: ' + this.files[0].name;
        uploadArea.style.borderColor = '#1565c0';
        uploadArea.style.background = '#e3f2fd';
    } else {
        fileName.textContent = '';
        uploadArea.style.borderColor = '#90caf9';
        uploadArea.style.background = '#f5f5f5';
    }
});

// Drag and drop functionality
uploadArea.addEventListener('dragover', function(e) {
    e.preventDefault();
    e.stopPropagation();
    this.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', function(e) {
    e.preventDefault();
    e.stopPropagation();
    this.classList.remove('dragover');
});

uploadArea.addEventListener('drop', function(e) {
    e.preventDefault();
    e.stopPropagation();
    this.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            fileInput.files = files;
            fileName.textContent = '✓ Selected: ' + file.name;
            uploadArea.style.borderColor = '#1565c0';
            uploadArea.style.background = '#e3f2fd';
        } else {
            alert('Please upload an Excel file (.xlsx or .xls)');
        }
    }
});

// Form submission with loading state
uploadForm.addEventListener('submit', function(e) {
    if (!fileInput.files || fileInput.files.length === 0) {
        e.preventDefault();
        alert('Please select a file first');
        return;
    }
    
    // Show loading state
    generateBtn.disabled = true;
    generateBtn.classList.add('loading');
    generateBtn.innerHTML = '<span class="btn-icon">⏳</span> Processing...';
});

// Auto-hide messages after 5 seconds
const messages = document.querySelectorAll('.message');
messages.forEach(function(message) {
    setTimeout(function() {
        message.style.opacity = '0';
        setTimeout(function() {
            message.remove();
        }, 300);
    }, 5000);
});
