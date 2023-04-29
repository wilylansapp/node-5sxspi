const express = require('express');
const app = express();
const Excel = require('exceljs');
const port = 3000;

app.get('/', (req, res) => {
  genererTableauExcel()
    .then((data) => res.send(data))
    .catch((err) => res.send(`Erreur : ${err}`));
});

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`);
});

async function genererTableauExcel() {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Feuille 1');

  // Vérifier si le tableau existe et initialiser s'il n'existe pas
  let dateColumns = null;
  if (!Array.isArray(dateColumns)) {
    dateColumns = Array.from({ length: 31 }, (_, i) => (i + 1).toString());
  }

  // Ajouter les activités avec les valeurs pour chaque jour du mois
  worksheet.addRow(['Activité', ...dateColumns, 'Total']);
  worksheet.addRow(['Site']);
  worksheet.addRow(['Télétravail']);
  worksheet.addRow(['Intercontrat']);
  worksheet.addRow(['suspension du contrat']);
  worksheet.addRow(['Congés payés']);
  worksheet.addRow(['RTT']);
  worksheet.addRow(['autres absences']);
  worksheet.addRow(['Maladie ou maternité']);
  worksheet.addRow(['TOTAL *']);

  // Adapter la taille des colonnes selon la taille des caractères
  const columnLengths = worksheet.columns.map((column) =>
    column.values.reduce((prev, next) => {
      const nextLength = next ? next.toString().length : 0;
      return Math.max(prev, nextLength);
    }, 0)
  );
  worksheet.columns.forEach((column, index) => {
    column.width = columnLengths[index] + 2;
  });

  // Mettre les weekends en gris et les autres jours en vert
  const daysInMonth = new Date(
    new Date().getFullYear(),
    new Date().getMonth() + 1,
    0
  ).getDate();
  for (let i = 1; i <= daysInMonth; i++) {
    const date = new Date(
      Date.UTC(new Date().getFullYear(), new Date().getMonth(), i)
    );
    const dayOfWeek = date.getUTCDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    worksheet.getColumn(i + 1).eachCell({ includeEmpty: true }, (cell) => {
      if (isWeekend) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '808080' },
        };
      } else {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '00FF00' },
        };
      }
    });
  }

  // Ajouter les formules pour calculer les totaux
  for (let i = 1; i <= daysInMonth; i++) {
    const column = worksheet.getColumn(i + 1);
    if (column.values[1]) {
      let totalActivites = 0;
      column.eachCell((cell, rowNumber) => {
        if (rowNumber > 1 && rowNumber < worksheet.rowCount) {
          const value = cell.value || 0;
          totalActivites += value;
        }
      });
      const totalFormula = `SUM(${column.getgetCell(2).address}:${
        column.lastCell ? column.lastCell.address : 'B2'
      })`;
      const totalCell = worksheet.getRow(worksheet.rowCount).getCell(i + 1);
      totalCell.value = { formula: totalFormula };
      totalCell.style = { font: { bold: true } };
      if (!isNaN(totalActivites) && totalActivites > 0) {
        const cellStyle = {
          font: { bold: true },
          fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '00FF00' },
          },
        };
        column.eachCell((cell, rowNumber) => {
          if (rowNumber > 1 && rowNumber < worksheet.rowCount) {
            const siteCell = worksheet.getRow(2).getCell(i + 1);
            const teletravailCell = worksheet.getRow(3).getCell(i + 1);
            const autresAbsencesCell = worksheet.getRow(8).getCell(i + 1);
            if (cell.value && cell.style.fill.fgColor.argb !== '808080') {
              cell.style = cellStyle;
              siteCell.style = cellStyle;
              teletravailCell.style = cellStyle;
              autresAbsencesCell.style = { font: { bold: true } };
            }
          }
        });
      }
    }
  }

  // Ajouter la formule pour le total du total
  const totalTotalFormula = `SUM(${
    worksheet.getCell(`$B$${worksheet.rowCount}`).address
  }:${
    worksheet.getCell(
      `$${Excel.utils.getExcelAlpha(daysInMonth + 1)}${worksheet.rowCount}`
    ).address
  })-SUM(${worksheet.getCell(`${Excel.utils.getExcelAlpha(i + 1)}2`).address}:${
    worksheet.getCell(
      `$${Excel.utils.getExcelAlpha(i + 1)}${worksheet.rowCount - 1}`
    ).address
  })`;
  const totalTotalCell = worksheet
    .getRow(worksheet.rowCount)
    .getCell(daysInMonth + 1);
  totalTotalCell.value = { formula: totalTotalFormula };
  totalTotalCell.style = { font: { bold: true } };

  // Enregistrer le fichier Excel
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}
