const startBtn = document.getElementById('start-btn');
const stopBtn = document.getElementById('stop-btn');
const resultsDiv = document.getElementById('results');
const saveBtn = document.getElementById('save-btn');

let recognition;

function startSpeechRecognition() {
    if ('webkitSpeechRecognition' in window) {
        recognition = new webkitSpeechRecognition();
        recognition.continuous = true;
        recognition.interimResults = true;

        recognition.onresult = (event) => {
            const transcript = [];
            for (let i = event.resultIndex; i < event.results.length; ++i) {
                if (event.results[i].isFinal) {
                    transcript.push(event.results[i][0].transcript);
                }
            }
            resultsDiv.textContent = transcript.join('. ');
            saveBtn.disabled = !transcript.length;
        };

        recognition.onstart = () => {
            startBtn.disabled = true;
            stopBtn.disabled = false;
        };

        recognition.onend = () => {
            startBtn.disabled = false;
            stopBtn.disabled = true;
        };

        recognition.start();
    } else {
        resultsDiv.textContent = 'Speech recognition not supported.';
        saveBtn.disabled = true;
    }
}

function stopSpeechRecognition() {
    recognition.stop();
}

function saveToExcel() {
    const data = [resultsDiv.textContent];
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const excelBuffer = XLSX.write(workbook, {
        bookType: 'xlsx',
        type: 'array'
    });
    const dataBlob = new Blob([excelBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = window.URL.createObjectURL(dataBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'transcription.xlsx';
    a.click();
}

startBtn.addEventListener('click', startSpeechRecognition);
stopBtn.addEventListener('click', stopSpeechRecognition);
saveBtn.addEventListener('click', saveToExcel);