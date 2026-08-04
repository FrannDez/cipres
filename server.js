const express = require('express');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const { PDFDocument } = require('pdf-lib');
const xlsx = require('xlsx');
const admin = require('firebase-admin');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Inicialización de Firebase Realtime Database
admin.initializeApp({
  databaseURL: "https://cipres-control-default-rtdb.firebaseio.com"
});
const db = admin.database();

app.use(express.json());
app.use(express.static('public'));

// -------------------------------------------------------------------
// 1. ENDPOINT: Procesar PDF de Cuasifactura
// -------------------------------------------------------------------
app.post('/api/parse-pdf', upload.single('pdf'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, message: 'No se recibió archivo PDF' });
        }

        const dataBuffer = req.file.buffer;
        const pdfData = await pdfParse(dataBuffer);
        const texto = pdfData.text;

        // 1. EXTRACCIÓN DEL MONTO TOTAL CUASIFACTURA
        let total = 0;
        // Busca "Total Cuasifactura" seguido de saltos de línea/espacios y captura el valor con signo $ o sin él
        const matchMonto = texto.match(/Total\s+Cuasifactura[\s\S]*?\$?\s*([\d\.,]+)/i);
        if (matchMonto) {
            // Convierte formato "104.000" o "104.000,00" a número entero/flotante
            total = parseFloat(matchMonto[1].replace(/\./g, '').replace(',', '.'));
        }

        // 2. EXTRACCIÓN DE CÓDIGOS Y SUMA REAL DE PRESTACIONES
        // Formato objetivo: "CT C001 A97", "CT C008 A97", etc.
        const codigosAgrupados = {};
        
        // Regex que busca el código de prestación y captura el primer número que le sigue en la misma línea (las prestaciones)
        const regexFilaPrestacion = /([A-Z]{2}\s+[A-Z0-9]{4}\s+[A-Z0-9]{3})\s+(\d+)/g;
        let match;

        while ((match = regexFilaPrestacion.exec(texto)) !== null) {
            const codigo = match[1].trim();
            const cantidadPrestaciones = parseInt(match[2], 10);

            if (!codigosAgrupados[codigo]) {
                codigosAgrupados[codigo] = { cantidadTotal: 0 };
            }
            // Suma la cantidad real de la columna "Prestaciones"
            codigosAgrupados[codigo].cantidadTotal += cantidadPrestaciones;
        }

        return res.json({
            success: true,
            totalCuasifactura: total,
            codigosAgrupados: codigosAgrupados
        });

    } catch (error) {
        console.error("Error al procesar PDF SUMAR:", error);
        return res.status(500).json({ success: false, message: 'Error interno al procesar el archivo' });
    }
});

// -------------------------------------------------------------------
// 2. ENDPOINT: Procesar Excel de Rechazos (Cualquier Obra Social)
// -------------------------------------------------------------------
app.post('/api/parse-excel', upload.single('excel'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'No se subió ningún archivo Excel.' });

    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convertir hoja a matriz de filas
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // Localizar la fila que contiene las cabeceras (NOMBRE, DOCUMENTO, MOTIVO)
    let headerIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i] && rows[i].includes('NOMBRE') && rows[i].includes('MOTIVO')) {
        headerIndex = i;
        break;
      }
    }

    if (headerIndex === -1) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ success: false, message: 'Estructura de Excel no válida. No se halló la cabecera NOMBRE/MOTIVO.' });
    }

    const headers = rows[headerIndex].map(h => String(h).trim().toUpperCase());
    const idxNombre = headers.indexOf('NOMBRE');
    const idxDoc = headers.indexOf('DOCUMENTO');
    const idxMotivo = headers.indexOf('MOTIVO');

    const rechazosObraSocial = [];

    // Recorrer filas de datos
    for (let i = headerIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[idxMotivo]) continue;

      const motivoTexto = String(row[idxMotivo]).trim().toUpperCase();

      // Evalúa si el motivo comienza o contiene la clave "EL PACIENTE REGISTRA:"
      if (motivoTexto.includes('EL PACIENTE REGISTRA')) {
        rechazosObraSocial.push({
          nombre: row[idxNombre] ? String(row[idxNombre]).trim() : 'SIN NOMBRE',
          documento: row[idxDoc] ? String(row[idxDoc]).trim() : 'SIN DOCUMENTO',
          motivo: String(row[idxMotivo]).trim()
        });
      }
    }

    fs.unlinkSync(req.file.path);
    res.json({ success: true, pacientes: rechazosObraSocial });

  } catch (error) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ success: false, error: error.message });
  }
});

// -------------------------------------------------------------------
// 3. ENDPOINT: Firmar PDFs y Generar el Consolidado Unificado
// -------------------------------------------------------------------
app.post('/api/export-consolidated', upload.fields([
  { name: 'consolidado', maxCount: 1 },
  { name: 'efectores', maxCount: 17 }
]), async (req, res) => {
  try {
    const mergedPdf = await PDFDocument.create();

    // 1. Unir Consolidado (escaneado con firmas oficiales)
    if (req.files['consolidado'] && req.files['consolidado'][0]) {
      const consolidadoPath = req.files['consolidado'][0].path;
      const consolidadoBuffer = fs.readFileSync(consolidadoPath);
      const consolidadoDoc = await PDFDocument.load(consolidadoBuffer);
      const copiedPages = await mergedPdf.copyPages(consolidadoDoc, consolidadoDoc.getPageIndices());
      copiedPages.forEach((page) => mergedPdf.addPage(page));
      fs.unlinkSync(consolidadoPath);
    }

    // 2. Cargar imagen de la firma digitalizada
    const firmaPath = path.join(__dirname, 'assets', 'firma.png'); 
    let firmaImage = null;
    if (fs.existsSync(firmaPath)) {
      const firmaBytes = fs.readFileSync(firmaPath);
      firmaImage = await mergedPdf.embedPng(firmaBytes);
    }

    // 3. Unir los PDFs de cada efector e incrustar firma
    if (req.files['efectores']) {
      for (const file of req.files['efectores']) {
        const pdfBuffer = fs.readFileSync(file.path);
        const pdfDoc = await PDFDocument.load(pdfBuffer);
        const pages = pdfDoc.getPages();

        // Estampar la firma en la última página sobre "Firma y Aclaración del Responsable del Centro"
        if (firmaImage && pages.length > 0) {
          const lastPage = pages[pages.length - 1];
          lastPage.drawImage(firmaImage, {
            x: 70,       // Posición horizontal sobre el pie de página
            y: 40,       // Posición vertical sobre el pie de página
            width: 140,  // Ancho de la firma
            height: 45,  // Alto de la firma
          });
        }

        const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
        copiedPages.forEach((page) => mergedPdf.addPage(page));
        fs.unlinkSync(file.path);
      }
    }

    const pdfBytes = await mergedPdf.save();
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=Consolidado_Final.pdf');
    res.send(Buffer.from(pdfBytes));

  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// -------------------------------------------------------------------
// 4. ENDPOINT: Guardar Registros en Firebase Realtime Database
// -------------------------------------------------------------------
app.post('/api/guardar-base', async (req, res) => {
  try {
    const { mesAno, datosEfectores, rechazos } = req.body;
    const claveMes = mesAno.replace('/', '-');
    
    const ref = db.ref(`cuasifacturas/${claveMes}`);
    await ref.set({
      datosEfectores,
      rechazos,
      fechaActualizacion: new Date().toISOString()
    });

    res.json({ success: true, message: 'Datos guardados correctamente en Firebase.' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor iniciado correctamente en puerto ${PORT}`));