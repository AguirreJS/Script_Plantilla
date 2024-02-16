const fs = require('fs');
const ExcelJS = require('exceljs');
const path = require('path');
const mkdirp = require('mkdirp');
const AdmZip = require('adm-zip');

// Ruta del archivo Excel
const excelFilePath = './Excel/Pacientes.xlsx';
var contador = 0;

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

    // Leer nombres, profesionales, DNIs profesionales, referentes, diagnósticos y profesiones de las columnas correspondientes
      const names = worksheet.getColumn(3).values.slice(1);  // un nombre a const (Nombre) numero de la columna (.getColumn(3))
    const paciente = worksheet.getColumn(3).values.slice(1);
    const dnipaciente = worksheet.getColumn(4).values.slice(1);
    const horariod = worksheet.getColumn(16).values.slice(1);
    const horahasta = worksheet.getColumn(17).values.slice(1);
    const prestacion = worksheet.getColumn(18).values.slice(1);

  

    // Ruta de la plantilla Word
    const templateFilePath = './word/carta.docx';

    // Leer la plantilla Word en formato zip
    const zip = new AdmZip(templateFilePath);
    const zipEntries = zip.getEntries();

    for (let i = 0; i < names.length; i++) {
      const name = names[i]; // nombre de la iteracion diferente al nombre de constante  const names = worksheet.getColumn(3).values.slice(1)
      const pacienteValue = paciente[i];
      const dnipacienteValue = dnipaciente[i];
      const horariodValue = horariod[i];
      const horahastaValue = horahasta[i];
      const prestacionValue = prestacion[i]; // Cambio de variable para evitar conflicto de nombres

contador = contador + 1;

      const outputFileName = `  ${name} - ${dnipacienteValue}.docx`;
      const outputPath = path.join(outputFolder, outputFileName);

      // Clonar la plantilla Word
      const clonedZip = new AdmZip();
      for (const entry of zipEntries) {
        const entryData = zip.readFile(entry);
        clonedZip.addFile(entry.entryName, entryData);
      }

      // Reemplazar el texto en el documento clonado
      let contentXml = clonedZip.readAsText('word/document.xml');
      contentXml = contentXml.replace(/{PACIENTE}/g, pacienteValue); // variable que debe buscar y el valor de la celda  const name = names[i];
      contentXml = contentXml.replace(/{DNIPACIENTE}/g, dnipacienteValue);
      contentXml = contentXml.replace(/{HORARIOD}/g, horariodValue);
      contentXml = contentXml.replace(/{HORAHASTA}/g, horahastaValue);
      contentXml = contentXml.replace(/{PRESTACION}/g, prestacionValue);

      clonedZip.updateFile('word/document.xml', Buffer.from(contentXml));

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
