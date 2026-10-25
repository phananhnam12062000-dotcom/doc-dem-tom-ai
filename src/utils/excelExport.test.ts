import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { generateShrimpReportWorkbookBuffer } from './excelExport.ts';

describe('generateShrimpReportWorkbookBuffer (.xlsx multi-sheet report)', () => {
  it('generates a valid 2-sheet Excel workbook with Navy Blue #1F4E79 headers, #F2F5F8 zebra rows, Segoe UI fonts, and live Excel formulas', async () => {
    const buffer = await generateShrimpReportWorkbookBuffer({
      pondName: 'Ao Số 1 - Khu A',
      staffName: 'Kỹ thuật viên 1',
      weightRecords: [
        { stt: 1, weight: 35.0, gender: 'Male', time: '08:15:10' },
        { stt: 2, weight: 45.0, gender: 'Female', time: '08:15:20' },
        { stt: 3, weight: 50.0, gender: 'Male', time: '08:15:30' },
      ],
      batchRecords: [
        {
          id: 1,
          batchCode: 'Mẻ #1',
          count: 35,
          cumulative: 35,
          time: '08:16:00',
          status: 'Hợp lệ',
        },
        {
          id: 2,
          batchCode: 'Mẻ #2',
          count: 40,
          cumulative: 75,
          time: '08:16:15',
          status: 'Hợp lệ',
        },
      ],
    });

    assert.ok(buffer.byteLength > 1000, 'XLSX buffer should be non-empty');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    assert.equal(wb.worksheets.length, 2, 'Workbook must contain exactly 2 sheets');

    const ws1 = wb.getWorksheet('1_Can_Trong_Luong_Gioi_Tinh');
    const ws2 = wb.getWorksheet('2_Dem_So_Luong_Cong_Don');
    assert.ok(ws1, 'Sheet 1 must exist');
    assert.ok(ws2, 'Sheet 2 must exist');

    // Verify Navy Blue #1F4E79 title and Segoe UI 16pt font
    const titleCell1 = ws1.getCell('A1');
    assert.equal(titleCell1.font?.name, 'Segoe UI');
    assert.equal(titleCell1.font?.size, 16);
    assert.equal(titleCell1.font?.bold, true);
    assert.equal(
      (titleCell1.fill as ExcelJS.FillPattern)?.fgColor?.argb,
      'FF1F4E79',
    );

    // Verify dynamic Excel formulas on Sheet 1 summary block
    const countFormula = (ws1.getCell('C7').value as ExcelJS.CellFormulaValue)
      ?.formula;
    const sumFormula = (ws1.getCell('C8').value as ExcelJS.CellFormulaValue)
      ?.formula;
    const avgFormula = (ws1.getCell('C9').value as ExcelJS.CellFormulaValue)
      ?.formula;
    const sizeFormula = (ws1.getCell('C10').value as ExcelJS.CellFormulaValue)
      ?.formula;

    assert.equal(countFormula, 'COUNT(B15:B17)');
    assert.equal(sumFormula, 'SUM(B15:B17)');
    assert.equal(avgFormula, 'IFERROR(AVERAGE(B15:B17), 0)');
    assert.equal(sizeFormula, 'IFERROR((C7/C8)*1000, 0)');

    // Verify alternating row fills (#FFFFFF and #F2F5F8) and alignments on Sheet 1 detail rows
    const row15Weight = ws1.getCell('B15');
    const row16Weight = ws1.getCell('B16');
    assert.equal(
      (row15Weight.fill as ExcelJS.FillPattern)?.fgColor?.argb,
      'FFFFFFFF',
    );
    assert.equal(
      (row16Weight.fill as ExcelJS.FillPattern)?.fgColor?.argb,
      'FFF2F5F8',
    );
    assert.equal(row15Weight.alignment?.horizontal, 'right');
    assert.equal(ws1.getCell('C15').alignment?.horizontal, 'center');

    // Verify Sheet 2 highlighted total formula and cumulative formula
    const totalBatchFormula = (
      ws2.getCell('E6').value as ExcelJS.CellFormulaValue
    )?.formula;
    assert.equal(totalBatchFormula, 'SUMIF(F10:F11,"Hợp lệ",C10:C11)');

    const runningCumulativeFormula = (
      ws2.getCell('D11').value as ExcelJS.CellFormulaValue
    )?.formula;
    assert.equal(
      runningCumulativeFormula,
      'SUMIF($F$10:F11,"Hợp lệ",$C$10:C11)',
    );
  });
});
