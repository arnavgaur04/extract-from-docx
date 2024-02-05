const fs = require('fs');
const unzipper = require('unzipper');
var DOMParser = new (require('xmldom')).DOMParser;

async function readXmlFromDocx(docxFilePath) {
    const xmlContents = {};

    await fs.createReadStream(docxFilePath)
        .pipe(unzipper.Parse())
        .on('entry', entry => {
            // Only read XML files from the 'word' directory
            if (entry.path.startsWith('word/') && entry.path.endsWith('.xml')) {
                entry.buffer().then(buffer => {
                    xmlContents[entry.path] = buffer.toString('utf-8');
                });
            } else {
                entry.autodrain();
            }
        })
        .promise();

    return xmlContents;
}

var documentXML;

// Function for returning text in line by line format from .docx

async function extractText(docxFilePath) {
return readXmlFromDocx(docxFilePath)
    .then(xmlContents => {
        var arr = [];
        documentXML = xmlContents['word/document.xml'];
        var document = DOMParser.parseFromString(documentXML);

        for (let i = 0; i < document.getElementsByTagName('w:t').length; i++) {
            arr.push(document.getElementsByTagName('w:t')[i].childNodes[0].nodeValue);
        }

        return arr;
    })
    .catch(error => {
        console.error('Error reading XML file');
    });
}

// Function for extracting tables from .docx. It returns a promise in the form of Object.
async function extractTables(docxFilePath) {
    return readXmlFromDocx(docxFilePath)
    .then(xmlContents => {
        var dicTable = {};
        documentXML = xmlContents['word/document.xml'];
        var document = DOMParser.parseFromString(documentXML);

        for (let i = 0; i < document.getElementsByTagName('w:tbl').length; i++) {
            var dicRow = {};
            for (let j = 0; j < document.getElementsByTagName('w:tbl')[i].getElementsByTagName('w:tr').length; j++) {
                var arr = [];
                for (let k = 0; k < document.getElementsByTagName('w:tbl')[i].getElementsByTagName('w:tr')[j].getElementsByTagName('w:t').length; k++) {
                    arr.push(document.getElementsByTagName('w:tbl')[i].getElementsByTagName('w:tr')[j].getElementsByTagName('w:t')[k].childNodes[0].nodeValue);
                }
                dicRow["row_"+j] = arr;
            }
            dicTable["table_"+i] = dicRow;
        }

        return dicTable;
    })
    .catch(error => {
        console.error('Error reading XML file');
    });
}



