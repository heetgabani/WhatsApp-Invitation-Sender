const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const open = require('open');

const app = express();
const port = 3000;

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.use(express.static('public'));

app.post('/upload', upload.single('excelFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    let workbook;
    try {
        workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    } catch (err) {
        return res.status(500).send('Error reading Excel file.');
    }

    let sheetName = workbook.SheetNames[0];
    let sheet = workbook.Sheets[sheetName];

    let phoneNumbers = [];

    for (let i = 1; ; i++) {
        let cell = sheet['A' + i];
        if (!cell || !cell.v) break;

        let phoneNumber = cell.v.toString().trim();
        phoneNumbers.push(phoneNumber);
    }

    let message = "Your invitation message here";

    let index = 0;

    function sendMessage() {
        if (index < phoneNumbers.length) {
            let number = phoneNumbers[index];
            let url = `https://web.whatsapp.com/send?phone=${encodeURIComponent(number)}&text=${encodeURIComponent(message)}`;

            open(url).then(() => {
                index++;
                sendMessage();
            }).catch(err => {
                console.error('Error opening WhatsApp Web:', err);
                index++;
                sendMessage();
            });
        } else {
            res.send('Messages sent successfully!');
        }
    }

    sendMessage();
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
