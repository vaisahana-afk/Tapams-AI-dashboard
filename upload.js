const fileInput = document.querySelector('#file-input');
const dropZone = document.querySelector('.drop-zone');
const browseButton = document.querySelector('.browse-button');
const filePreview = document.querySelector('.file-preview');
const fileName = document.querySelector('.file-preview__name');
const fileMeta = document.querySelector('.file-preview__meta');
const removeFileButton = document.querySelector('.remove-file');
const analyzeButton = document.querySelector('.analyze-button');
const uploadStatus = document.querySelector('.upload-status');
const sampleDownload = document.querySelector('.sample-download');

let selectedFile = null;

const sampleRows = [
  ['S.No', 'Customer Code', 'Customer Name', 'Value', 'Category', 'Product', 'Invoice Date', 'Effective From', 'Effective To', 'Last Enhancement Date'],
  [1, 'CUSI002', 'INDOCOOL COMPOSITES PRIVATE', 6697785.46, 'CCCP', 'Cool Composite 2021', '2025-01-15', '2021-01-01', '2026-12-31', '2021-06-30'],
  [2, 'CUSJ002', 'JS AUTO CAST FOUNDRY INDIA PRIVATE', 1044788, 'CCCP', 'Auto Cast Pro', '2025-02-10', '2020-04-01', '2026-03-31', '2020-04-01'],
  [3, 'CUSP004', 'PTERIS GLOBAL INTEGRATED SOLUTION', 1347071.5, 'NCNP', 'Pteris SmartLine', '2025-03-05', '2025-03-01', '2027-02-28', '2025-03-01'],
  [4, 'CUSR005', 'RISHI LASER LTD', 79024, 'NCNP', 'Laser Connect', '2025-03-20', '2025-03-15', '2026-03-14', '2025-03-15'],
  [5, 'CUSS001', 'SRI VENKATESWARA PRECISION COMPONENTS', 24575996.5, 'CCCP', 'Precision Core', '2025-04-12', '2019-01-01', '2025-12-31', '2019-01-01'],
  [6, 'CUS002', 'SVPC AUTOMOTIVE (P) LTD', 1700855.98, 'CCCP', 'AutoMotion Base', '2025-05-03', '2022-05-01', '2027-04-30', '2022-05-01'],
  [7, 'CUS019', 'SPECTRA FABTECH LLP', 9555, 'NCNP', 'FabTech Insight', '2025-05-19', '2025-05-01', '2026-04-30', '2025-05-01'],
  [8, 'CUST002', 'Tritium Pty Ltd', 96127.76, 'CCCP', 'Tritium Classic', '2025-06-08', '2018-06-01', '2025-05-31', '2018-06-01'],
  [9, 'CUST003', 'TVS SUPPLY CHAIN SOLUTIONS LIMITED', 1918889.74, 'NCCP', 'Supply Chain Edge', '2025-06-22', '2024-06-01', '2027-05-31', '2024-06-01'],
  [10, 'CUSV001', 'VESTAS WIND TECHNOLOGY INDIA', 11058408.92, 'CCCP', 'WindCore Legacy', '2025-07-15', '2017-01-01', '2025-12-31', '2017-01-01'],
  [11, 'CUSV002', 'VESTAS WIND TECHNOLOGY INDIA', 4680252, 'CCCP', 'WindCore Legacy', '2025-08-01', '2017-01-01', '2025-12-31', '2017-01-01'],
  [12, 'CUSV003', 'VARICON PUMPS & SYSTEMS PVT LTD', 87680, 'NCNP', 'PumpSense', '2025-08-18', '2025-08-01', '2026-07-31', '2025-08-01'],
];

const clearPreviousFile = indexedDB.open('northstar-workspace', 1);
clearPreviousFile.onupgradeneeded = () => clearPreviousFile.result.createObjectStore('files');
clearPreviousFile.onsuccess = () => {
  const transaction = clearPreviousFile.result.transaction('files', 'readwrite');
  transaction.objectStore('files').delete('latest');
};

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function showFile(file) {
  if (!file) return;

  selectedFile = file;
  fileName.textContent = file.name;
  fileMeta.textContent = `${file.type || 'File'} · ${formatFileSize(file.size)}`;
  filePreview.hidden = false;
  dropZone.hidden = true;
  analyzeButton.disabled = false;
  uploadStatus.textContent = 'File ready to be read.';
}

function clearFile() {
  selectedFile = null;
  fileInput.value = '';
  filePreview.hidden = true;
  dropZone.hidden = false;
  analyzeButton.disabled = true;
  uploadStatus.textContent = '';
}

browseButton.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' || event.key === ' ') {
    event.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener('change', () => showFile(fileInput.files[0]));

['dragenter', 'dragover'].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add('drop-zone--active');
  });
});

['dragleave', 'drop'].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove('drop-zone--active');
  });
});

dropZone.addEventListener('drop', (event) => showFile(event.dataTransfer.files[0]));
removeFileButton.addEventListener('click', clearFile);

analyzeButton.addEventListener('click', () => {
  if (!selectedFile) return;

  const request = indexedDB.open('northstar-workspace', 1);
  request.onupgradeneeded = () => request.result.createObjectStore('files');
  request.onsuccess = () => {
    const transaction = request.result.transaction('files', 'readwrite');
    transaction.objectStore('files').put(selectedFile, 'latest');
    transaction.oncomplete = () => {
      window.location.href = 'dashboard.html';
    };
  };
  request.onerror = () => {
    uploadStatus.textContent = 'This browser could not prepare the file. Please try again.';
  };
});

sampleDownload.addEventListener('click', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(sampleRows);
  sheet['!cols'] = [{ wch: 8 }, { wch: 16 }, { wch: 40 }, { wch: 16 }, { wch: 12 }, { wch: 24 }, { wch: 15 }, { wch: 17 }, { wch: 15 }, { wch: 23 }];
  XLSX.utils.book_append_sheet(workbook, sheet, 'Effective Dated Sales');
  XLSX.writeFile(workbook, 'tapams-effective-dated-sample.xlsx');
  uploadStatus.textContent = 'Sample Excel downloaded. Upload it to see the effective-dated intelligence.';
});
