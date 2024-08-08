document.getElementById('fileInput').addEventListener('change', handleFileSelect, false);

function handleFileSelect(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const contents = e.target.result;
        let textContent = '';

        if (file.name.endsWith('.xlsx')) {
            const workbook = XLSX.read(contents, { type: 'binary' });
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                textContent += XLSX.utils.sheet_to_csv(worksheet);
            });
        } else if (file.name.endsWith('.csv')) {
            Papa.parse(contents, {
                complete: function(results) {
                    textContent = results.data.map(row => row.join(', ')).join('\n');
                    const vttContent = textToVtt(textContent);
                    displayOutput(vttContent);
                }
            });
            return;
        } else if (file.name.endsWith('.docx')) {
            mammoth.extractRawText({ arrayBuffer: contents }).then(result => {
                textContent = result.value;
                const vttContent = textToVtt(textContent);
                displayOutput(vttContent);
            });
            return;
        } else {
            textContent = contents;
        }

        const vttContent = textToVtt(textContent);
        displayOutput(vttContent);
    };

    if (file.name.endsWith('.xlsx')) {
        reader.readAsBinaryString(file);
    } else if (file.name.endsWith('.csv')) {
        reader.readAsText(file);
    } else if (file.name.endsWith('.docx')) {
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsText(file);
    }
}

function textToVtt(text) {
    const vtt = ['WEBVTT\n\n'];
    const lines = text.split(/\r?\n/);
    let cueIndex = 1;

    lines.forEach((line, index) => {
        if (line.trim() !== '') {
            vtt.push(`${cueIndex}`);
            vtt.push(`00:00:${index.toString().padStart(2, '0')}.000 --> 00:00:${(index + 1).toString().padStart(2, '0')}.000`);
            vtt.push(line);
            vtt.push('');
            cueIndex++;
        }
    });

    return vtt.join('\n');
}

function textToSrt(text) {
    const srt = [];
    const lines = text.split(/\r?\n/);
    let cueIndex = 1;

    lines.forEach((line, index) => {
        if (line.trim() !== '') {
            srt.push(`${cueIndex}`);
            srt.push(`00:00:${index.toString().padStart(2, '0')},000 --> 00:00:${(index + 1).toString().padStart(2, '0')},000`);
            srt.push(line);
            srt.push('');
            cueIndex++;
        }
    });

    return srt.join('\n');
}

function displayOutput(content) {
    const output = document.getElementById('output');
    output.value = content;
}

function downloadFile() {
    const outputContent = document.getElementById('output').value;
    const format = document.getElementById('formatSelect').value;
    const blob = new Blob([outputContent], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `converted.${format}`;
    document.body.appendChild(a);
    a.style.display = 'none';
    a.click();
    document.body.removeChild(a);
}

function convertFile() {
    const fileInput = document.getElementById('fileInput');
    const event = new Event('change');
    fileInput.dispatchEvent(event);
}
