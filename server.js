const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(bodyParser.json({limit: '200mb'}));
app.use(express.static('public'));

app.get('/pdf', (req, res) => {
    res.sendFile(path.resolve('C:/Users/명현/Downloads/111에서113어깨.pdf'));
});

app.post('/save-image', (req, res) => {
    const { pageNum, imageBase64 } = req.body;
    const base64Data = imageBase64.replace(/^data:image\/png;base64,/, "");
    const imgFolder = path.resolve('pages');
    if (!fs.existsSync(imgFolder)) {
        fs.mkdirSync(imgFolder, { recursive: true });
    }
    const filename = `page_${pageNum}.png`;
    fs.writeFileSync(path.join(imgFolder, filename), base64Data, 'base64');
    console.log(`Saved ${filename}`);
    res.json({ success: true });
});

app.post('/done', (req, res) => {
    console.log("All pages saved!");
    res.json({ success: true });
    setTimeout(() => process.exit(0), 1000);
});

app.listen(3001, () => {
    console.log('Server running on http://localhost:3001');
});
