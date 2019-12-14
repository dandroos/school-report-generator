const PDFDocument = require("pdfkit");
const GS = require("google-spreadsheet");
const moment = require("moment");
const async = require("async");
const fs = require("fs");

moment.locale('es')

const doc = new GS("1fZs7RAyY0bFrWxQuhIa3GoahT6p10lByaAFxLOkHrV8");
let sheet;

const fonts = {
  normal: "Helvetica",
  bold: "Helvetica-Bold",
  italic: "Helvetica-Italic"
};

async.series([
  step => {
    doc.useServiceAccountAuth(require("./creds.json"), step);
  },
  step => {
    doc.getInfo((err, info) => {
      info.worksheets.map(sh => {
        if (sh.title === "Reports") {
          sheet = sh;
          step();
        }
      });
    });
  },
  step => {
    doc.getRows(sheet.id, (err, rows) => {
      const PDF = new PDFDocument({ autoFirstPage: false });
      PDF.pipe(fs.createWriteStream("./reports.pdf"));
      rows.map(student => {
        PDF.addPage()
          .image(
            "./logo.png",
            PDF.page.margins.left,
            PDF.page.margins.top - 10,
            { width: 150 }
          )
          .font(fonts.normal)
          .fontSize(12)
          .text("El Rincón de Idiomas", { align: "right" })
          .text("Calle Alcaldes Mayores, 1", { align: "right" })
          .text("6 Majada Marcial", { align: "right" })
          .text("Puerto del Rosario", { align: "right" })
          .text("35600", { align: "right" })
          .moveDown()
          .text("636 32 34 41", { align: "right" })
          .moveDown()
          .fontSize(10)
          .text("Directora: Joanne Bowker", { align: "right" })
          .text("X2016168B", { align: "right" })
          .moveDown()
          .fontSize(12)
          .text(moment().format("LL"), { align: "right" })
          .moveDown(2)
          .fontSize(25)
          .font(fonts.bold)
          .text("Evaluación de Ingles")
          .fontSize(18)
          .moveDown()
          .text(student.name)
          .font(fonts.normal)
          .fontSize(15)
          .moveDown()
          .text("Oral", { width: 200, continued: true })
          .text(student.speaking, { align: "right" })
          .text("Escrito", { width: 200, continued: true })
          .text(student.writing, { align: "right" })
          .text("Comprensión Lectura", { width: 200, continued: true })
          .text(student.reading, { align: "right" })
          .text("Comprensión Auditiva", { width: 200, continued: true })
          .text(student.listening, { align: "right" })
          .text("Asistencia", { width: 200, continued: true })
          .text(student.attendance, { align: "right" })
          .text("Comportamiento", { width: 200, continued: true })
          .text(student.behaviour, { align: "right" })
          .moveDown(2)
          .font(fonts.bold)
          .text("Observaciones del profesor")
          .moveDown()
          .font(fonts.normal)
          .text(student.comments, {align: 'justify'})
          .moveDown();
        if (student.exam) {
          PDF.font(fonts.bold)
            .text("Su hij@ esta en el nivel de Cambridge: ", {
              continued: true
            })
            .font(fonts.normal)
            .text(student.exam)
            .moveDown();
        }
        PDF.image("./cambridge.jpeg", { width: 250 });
      });
      PDF.end();
    });
  }
]);
