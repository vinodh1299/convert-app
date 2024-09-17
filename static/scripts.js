// JavaScript to toggle between file and folder selection
function toggleSelectionType() {
    const selectionType = document.getElementById('selectionType').value;
    const fileInputContainer = document.getElementById('fileInputContainer');
    const folderInputContainer = document.getElementById('folderInputContainer');

    if (selectionType === 'file') {
        fileInputContainer.style.display = 'block';
        folderInputContainer.style.display = 'none';
    } else {
        fileInputContainer.style.display = 'none';
        folderInputContainer.style.display = 'block';
    }
}

document.getElementById('uploadForm').addEventListener('submit', async function (e) {
    e.preventDefault();
    let formData = new FormData();
    const selectionType = document.getElementById('selectionType').value;
    const conversionType = document.getElementById('conversionType').value;

    formData.append('conversionType', conversionType);

    if (selectionType === 'file') {
        let files = document.getElementById('fileInput').files;
        for (let i = 0; i < files.length; i++) {
            formData.append('files', files[i]);
        }
    } else {
        let folderFiles = document.getElementById('folderInput').files;
        for (let i = 0; i < folderFiles.length; i++) {
            formData.append('files', folderFiles[i]);
        }
    }

    // Show progress container
    document.getElementById('progressContainer').style.display = 'block';

    let response = await fetch('/convert', {
        method: 'POST',
        body: formData
    });

    let result = await response.json();
    document.getElementById('progressContainer').style.display = 'none';
    document.getElementById('resultContainer').style.display = 'block';
    document.getElementById('downloadLink').href = result.downloadLink;
});
