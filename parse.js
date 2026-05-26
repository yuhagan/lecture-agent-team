const fs = require('fs');
const path = require('path');
const { createCanvas } = require('canvas');

// Polyfill canvas for PDF.js in Node
const NodeCanvasFactory = {
  create(width, height) {
    const canvas = createCanvas(width, height);
    const context = canvas.getContext('2d');
    return { canvas, context };
  },
  reset(canvasAndContext, width, height) {
    canvasAndContext.canvas.width = width;
    canvasAndContext.canvas.height = height;
  },
  destroy(canvasAndContext) {
    canvasAndContext.canvas.width = 0;
    canvasAndContext.canvas.height = 0;
    canvasAndContext.canvas = null;
    canvasAndContext.context = null;
  }
};

async function extractPages() {
    const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
    
    const pdfPath = 'C:/Users/명현/Downloads/111에서113어깨.pdf';
    const data = new Uint8Array(fs.readFileSync(pdfPath));
    const loadingTask = pdfjsLib.getDocument({
        data,
        cMapUrl: path.join(__dirname, 'node_modules/pdfjs-dist/cmaps/'),
        cMapPacked: true,
        standardFontDataUrl: path.join(__dirname, 'node_modules/pdfjs-dist/standard_fonts/')
    });
    
    const pdfDocument = await loadingTask.promise;
    console.log(`PDF loaded. Pages: ${pdfDocument.numPages}`);
    
    for (let i = 1; i <= pdfDocument.numPages; i++) {
        const page = await pdfDocument.getPage(i);
        const viewport = page.getViewport({ scale: 2.0 });
        
        const canvasFactory = NodeCanvasFactory;
        const canvasAndContext = canvasFactory.create(viewport.width, viewport.height);
        
        const renderContext = {
            canvasContext: canvasAndContext.context,
            viewport: viewport,
            canvasFactory: canvasFactory
        };
        
        await page.render(renderContext).promise;
        const image = canvasAndContext.canvas.toBuffer('image/png');
        fs.writeFileSync(path.join(__dirname, 'pages', `page_${i}.png`), image);
        console.log(`Saved page_${i}.png`);
    }
}

extractPages().catch(console.error);
