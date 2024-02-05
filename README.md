# extract-from-docx

Extracts text and table data from .docx files and returns a Promise in the form of an object.

## Installation

**You can install extract-from-docx via npm or yarn:**

**Node**
    ```bash
    npm install extract-from-docx

**Yarn**
    ```bash
    yarn add extract-from-docx

## Usage

    ```javascript
    const { extractText, extractTables } = require('extract-from-docx');

    // Specify the path to the .docx file
    const docxFilePath = 'path/to/your/document.docx';

    // Extract text from the document
    extractText(docxFilePath)
    .then(text => {
        console.log('Extracted Text:', text);
    })
    .catch(error => {
        console.error('Error extracting text:', error);
    });

    // Extract tables from the document
    extractTables(docxFilePath)
    .then(tables => {
        console.log('Extracted Tables:', tables);
    })
    .catch(error => {
        console.error('Error extracting tables:', error);
    });