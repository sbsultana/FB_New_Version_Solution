import { Injectable } from '@angular/core';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

@Injectable({ providedIn: 'root' })
export class ExportService {

    getTable(tableId: string): HTMLTableElement {
        return document.getElementById(tableId) as HTMLTableElement;
    }
    exportExcel(tableId: string, fileName: string) {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Summary');

        // ✅ HEADER ROW 1 (Months)
        const months = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December', 'Year End'
        ];

        let colIndex = 2;

        months.forEach(month => {
            worksheet.mergeCells(1, colIndex, 1, colIndex + 2);
            worksheet.getCell(1, colIndex).value = month;
            worksheet.getCell(1, colIndex).alignment = { horizontal: 'center' };
            worksheet.getCell(1, colIndex).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF0554EF' }
            };
            worksheet.getCell(1, colIndex).font = { bold: true, color: { argb: 'FFFFFFFF' } };

            colIndex += 3;
        });

        // ✅ HEADER ROW 2 (Budget / Actual / Variance)
        const subHeaders = ['Budget', 'Actual', 'Variance'];

        colIndex = 2;
        months.forEach(() => {
            subHeaders.forEach((sub, i) => {
                const cell = worksheet.getCell(2, colIndex + i);
                cell.value = sub;
                cell.alignment = { horizontal: 'center' };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF4584FF' }
                };
                cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            });
            colIndex += 3;
        });

        // ✅ FIRST COLUMN HEADER
        worksheet.mergeCells(1, 1, 2, 1);
        worksheet.getCell('A1').value = 'Summary';
        worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('A1').font = { bold: true };

        // ✅ DATA
        const data = [
            {
                name: 'Total Income',
                values: [
                    50000, 77000, 27000,
                    5000, 4500, -500,
                    35000, 35000, 0,
                    58000, 63000, 5000
                ]
            },
            {
                name: 'Total Expenses',
                values: [
                    17500, 38500, 21000
                ]
            },
            {
                name: 'Net Income',
                values: [
                    32500, 38500, 6000
                ]
            }
        ];

        let rowIndex = 3;

        data.forEach((row, rIndex) => {

            const excelRow = worksheet.getRow(rowIndex);

            // First column
            excelRow.getCell(1).value = row.name;
            excelRow.getCell(1).font = { bold: true };

            row.values.forEach((val, cIndex) => {
                const cell = excelRow.getCell(cIndex + 2);
                cell.value = val;

                // Currency format
                cell.numFmt = '"$"#,##0.00';

                // Color negative
                if (val < 0) {
                    cell.font = { color: { argb: 'FFFF0000' } };
                }
            });

            // Alternate row color
            if (rIndex % 2 === 0) {
                excelRow.eachCell(cell => {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFF9FBFF' }
                    };
                });
            }

            rowIndex++;
        });

        // ✅ AUTO WIDTH
        worksheet.columns.forEach(col => {
            col.width = 14;
        });

        // ✅ EXPORT
        workbook.xlsx.writeBuffer().then(buffer => {
            saveAs(new Blob([buffer]), 'Summary.xlsx');
        });
    }

    exportPDF(tableId: string, title: string) {
        const table = this.getTable(tableId);

        const doc = new jsPDF('l', 'pt', 'a4');

        doc.text(title, 40, 30);

        autoTable(doc, {
            html: table,
            startY: 50,

            styles: {
                fontSize: 7,
                cellPadding: 3,
                valign: 'middle'
            },

            theme: 'grid',

            didParseCell: function (data) {

                // ✅ HEADER STYLING
                if (data.section === 'head') {

                    const rowIndex = data.row.index;

                    // 🔵 First header row (dark blue)
                    if (rowIndex === 0) {
                        data.cell.styles.fillColor = [5, 84, 239];
                        data.cell.styles.textColor = [255, 255, 255];
                        data.cell.styles.fontStyle = 'bold';
                    }

                    // 🔵 Second header row (light blue)
                    else if (rowIndex === 1) {
                        data.cell.styles.fillColor = [69, 132, 255];
                        data.cell.styles.textColor = [255, 255, 255];
                        data.cell.styles.fontStyle = 'bold';
                    }

                    // ⚪ Extra headers (multi-header fallback)
                    else {
                        data.cell.styles.fillColor = [217, 231, 255];
                        data.cell.styles.textColor = [0, 0, 0];
                    }
                }

                // ✅ BODY STYLING
                if (data.section === 'body') {

                    const rowIndex = data.row.index;

                    // even rows
                    if (rowIndex % 2 === 0) {
                        data.cell.styles.fillColor = [249, 251, 255];
                    }

                    // odd rows
                    else {
                        data.cell.styles.fillColor = [255, 255, 255];
                    }

                    // right align except first column
                    if (data.column.index !== 0) {
                        data.cell.styles.halign = 'right';
                    } else {
                        data.cell.styles.halign = 'left';
                    }
                }
            }
        });

        doc.save(title + '.pdf');
    }

    print(tableId: string, title: string) {
        const tableHtml = this.getTable(tableId).outerHTML;

        const win = window.open('', '', 'width=1200,height=800');

        win?.document.write(`
      <html>
        <head>
          <title>${title}</title>
          <style>
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ccc; padding: 4px; }
            thead { background: #0554ef; color: #fff; }
          </style>
        </head>
        <body>
          <h3>${title}</h3>
          ${tableHtml}
        </body>
      </html>
    `);

        win?.document.close();
        win?.print();
    }

    // ✅ EMAIL PDF
    async emailPDF(tableId: string, title: string) {
        const table = this.getTable(tableId);
        const doc = new jsPDF('l', 'pt', 'a4');
        autoTable(doc, {
            html: table,
            startY: 40
        });
        const blob = doc.output('blob');
        const formData = new FormData();
        formData.append('file', blob, `${title}.pdf`);
        formData.append('title', title);
        return fetch('/api/send-email', {
            method: 'POST',
            body: formData
        });
    }
}