const dropAreaMaster = document.getElementById('drop-area-master');
const fileInputMaster = document.getElementById('master-input');
const dropAreaTests = document.getElementById('drop-area-tests');
const fileInputTests = document.getElementById('tests-input');
const masterLabel = document.getElementById('master-text');
const testsLabel = document.getElementById('tests-text');
const compareButton = document.getElementById('compare-button');

class FileSelector {
    masterFile = null;
    testFiles = null;
}

const fileSelector = new FileSelector();

['dragenter', 'dragover'].forEach(eventName => {
    dropAreaMaster.addEventListener(eventName, (e) => highlight(e, dropAreaMaster), false);
    dropAreaTests.addEventListener(eventName, (e) => highlight(e, dropAreaTests), false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropAreaMaster.addEventListener(eventName, (e) => unhighlight(e, dropAreaMaster), false);
    dropAreaTests.addEventListener(eventName, (e) => unhighlight(e, dropAreaTests), false);
});

dropAreaMaster.addEventListener('drop', handleDropMaster, false);
dropAreaTests.addEventListener('drop', handleDropTests, false);
fileInputMaster.addEventListener('change', handleSelectMaster, false);
fileInputTests.addEventListener('change', handleSelectTests, false);

function highlight(e, area) {
    e.preventDefault();
    e.stopPropagation();
    area.classList.add('highlight');
}

function unhighlight(e, area) {
    e.preventDefault();
    e.stopPropagation();
    area.classList.remove('highlight');
}

function handleSelectMaster(e) {
    const files = e.target.files;
    const filesLength = files.length;

    if (filesLength > 1) {
        alert("Puoi inserire un solo file master!");
        return;
    }

    setLabelText(e, filesLength, masterLabel);
    fileSelector.masterFile = files[0];
    checkCompareButton();
}

function handleDropMaster(e) {
    e.preventDefault();
    e.stopPropagation();

    const dt = e.dataTransfer;
    const files = dt.files;
    const filesLength = files.length;

    if (filesLength > 1) {
        alert("Puoi inserire un solo file master!");
        return;
    }

    setLabelText(e, filesLength, masterLabel);
    fileSelector.masterFile = files[0];
    checkCompareButton();
}

function handleSelectTests(e) {
    const files = e.target.files;
    const filesLength = files.length;

    setLabelText(e, filesLength, masterLabel);
    fileSelector.testFiles = files;
    checkCompareButton();
}

function handleDropTests(e) {
    e.preventDefault();
    e.stopPropagation();

    const dt = e.dataTransfer;
    const files = dt.files;
    const filesLength = files.length;

    setLabelText(e, filesLength, testsLabel);
    fileSelector.testFiles = files;
    checkCompareButton();
}

function setLabelText(e, filesLength, label) {
    label.textContent = filesLength + " file selezionat" + (filesLength===1 ? "o" : "i");
}

function checkCompareButton() {
    if (fileSelector.masterFile && fileSelector.testFiles && fileSelector.testFiles.length > 0) {
        compareButton.disabled = false;
    }
}