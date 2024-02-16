const fs = require('fs');
const ExcelJS = require('exceljs');
const path = require('path');
const mkdirp = require('mkdirp');
const AdmZip = require('adm-zip');

var contador = 0;

// Ruta del archivo Excel
const excelFilePath = './Excel/Pacientes.xlsx';

// Carpeta de destino para los archivos Word
const outputFolder = './Destino';

async function processExcelAndWord() {
  try {
    // Crear la carpeta de destino si no existe
    await mkdirp(outputFolder);

    // Leer el archivo Excel
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    const worksheet = workbook.getWorksheet(1);

    // Leer nombres de la primera columna del archivo Excel
    const names = worksheet.getColumn(1).values.slice(1); // Saltar el encabezado

    // Ruta de la plantilla Word
    const templateFilePath = './word/carta.docx';

    // Leer la plantilla Word en formato zip
    const zip = new AdmZip(templateFilePath);
    const zipEntries = zip.getEntries();

    for (const name of names) {
      contador = contador + 1;
      const outputFileName = ` ${contador} LISTO ${name}.docx`;
      const outputPath = path.join(outputFolder, outputFileName);

      // Clonar la plantilla Word
      const clonedZip = new AdmZip();
      for (const entry of zipEntries) {
        const entryData = zip.readFile(entry);
        clonedZip.addFile(entry.entryName, entryData);
      }

      // Reemplazar el texto en el documento clonado
      const contentXml = clonedZip.readAsText('word/document.xml');
      const replacedContent = contentXml.replace(/{NOMBREPACIENTE}/g, name);
      clonedZip.updateFile('word/document.xml', Buffer.from(replacedContent));

      // Guardar el archivo Word clonado con el nombre adecuado
      clonedZip.writeZip(outputPath);
      console.log(`Archivo ${outputFileName} creado`);
    }

    console.log('Proceso completado.');
  } catch (error) {
    console.error(`Error: ${error.message}`);
  }
}

// Llamar a la función para procesar el Excel y crear archivos Word
processExcelAndWord();
