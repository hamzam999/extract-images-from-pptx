const fs = require('fs');
const yauzl = require('yauzl');

// Function to extract images from pptx file
function extractImagesFromPPTX(pptxFilePath, outputFolder) {
  yauzl.open(pptxFilePath, { lazyEntries: true }, (err, zipfile) => {
    if (err) throw err;

    zipfile.readEntry();
    zipfile.on('entry', (entry) => {
      if (/ppt\/media\/image\d+\.(jpeg|png|gif|jpg)$/i.test(entry.fileName)) {
        zipfile.openReadStream(entry, (err, readStream) => {
          if (err) throw err;

          const outputPath = `${outputFolder}/${entry.fileName.split('/').pop()}`;
          const writeStream = fs.createWriteStream(outputPath);

          readStream.pipe(writeStream);

          writeStream.on('close', () => {
            console.log(`Image extracted: ${outputPath}`);
            zipfile.readEntry();
          });
        });
      } else {
        zipfile.readEntry();
      }
    });

    zipfile.on('end', () => {
      console.log('Extraction completed.');
    });
  });
}

// Get command line arguments
const [, , pptxFilePath, outputFolder] = process.argv;

// Check if required arguments are provided
if (!pptxFilePath || !outputFolder) {
  console.error('Usage: node extractImages.js path/to/your/presentation.pptx path/to/output/folder');
  process.exit(1);
}

// Call the function with provided arguments
extractImagesFromPPTX(pptxFilePath, outputFolder);
