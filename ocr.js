const Tesseract = require('tesseract.js');
const fs = require('fs');
const path = require('path');

async function processPage(pageNum) {
    const imagePath = path.join(__dirname, 'pages', `page_${pageNum}.png`);
    console.log(`Processing page ${pageNum}...`);
    
    const result = await Tesseract.recognize(
        imagePath,
        'kor+eng',
        { logger: m => {} }
    );
    const data = result.data;
    
    // Some versions output blocks, lines, or words
    // Let's get lines
    const lines = data.lines || [];
    
    const layout = lines.map(b => ({
        text: b.text.trim(),
        bbox: b.bbox
    })).filter(b => b.text.length > 0);
    
    fs.writeFileSync(path.join(__dirname, 'pages', `page_${pageNum}_layout.json`), JSON.stringify(layout, null, 2));
    console.log(`Done page ${pageNum}, found ${lines.length} lines.`);
}

async function run() {
    for (let i = 1; i <= 3; i++) {
        await processPage(i);
    }
}

run().catch(console.error);
