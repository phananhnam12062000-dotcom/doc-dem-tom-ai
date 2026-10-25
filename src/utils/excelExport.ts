import ExcelJS from 'exceljs';
import { WeightRecord } from './counterLogic';

export interface BatchExportRecord {
  id: number;
  batchCode: string;
  count: number;
  cumulative: number;
  time: string;
  status: 'Hợp lệ' | 'Đã hoàn tác (Undo)';
}

const NAVY_HEADER_HEX = 'FF1F4E79'; // #1F4E79 Navy Blue
const ZEBRA_ROW_HEX = 'FFF2F5F8';   // #F2F5F8 Light Blue-Grey
const WHITE_HEX = 'FFFFFFFF';       // #FFFFFF
const ORANGE_HIGHLIGHT_HEX = 'FFFCE4D6'; // Highlight cam nhạt cho khối Tổng số Sheet 2
const BORDER_COLOR_HEX = 'FFD0D7DE';

const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: BORDER_COLOR_HEX } },
  left: { style: 'thin', color: { argb: BORDER_COLOR_HEX } },
  bottom: { style: 'thin', color: { argb: BORDER_COLOR_HEX } },
  right: { style: 'thin', color: { argb: BORDER_COLOR_HEX } },
};

/**
 * Tự động căn chỉnh độ rộng các cột (Auto-fit column width) dựa theo nội dung dài nhất,
 * đảm bảo không bao giờ bị che khuất chữ hoặc tràn lề.
 */
function autoFitWorksheetColumns(ws: ExcelJS.Worksheet, minWidth = 14) {
  ws.columns.forEach((column) => {
    let maxLen = minWidth;
    if (column && column.eachCell) {
      column.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
        // Bỏ qua dòng tiêu đề gộp ô (dòng 1) và khối gộp ô lớn để tránh làm cột A quá rộng
        if (rowNumber === 1) return;
        let cellText = '';
        const val = cell.value;
        if (val === null || val === undefined) {
          cellText = '';
        } else if (typeof val === 'object' && 'formula' in val) {
          cellText = val.result !== undefined ? String(val.result) : '123,456.78';
        } else {
          cellText = String(val);
        }
        if (cellText.length > maxLen) {
          maxLen = cellText.length;
        }
      });
    }
    column.width = Math.min(Math.max(maxLen + 5, minWidth), 42);
  });
}

/**
 * Tạo Workbook Excel (.xlsx) chuyên nghiệp gồm 2 Sheet với đầy đủ công thức tự động
 * (SUM, AVERAGE, COUNT, COUNTIF, SUMIF) theo đúng chuẩn giao diện Thủy sản (#1F4E79 & #F2F5F8).
 */
export async function generateShrimpReportWorkbookBuffer(params: {
  pondName: string;
  staffName: string;
  weightRecords: WeightRecord[];
  batchRecords: BatchExportRecord[];
}): Promise<ArrayBuffer> {
  const { pondName, staffName, weightRecords, batchRecords } = params;
  const exportTimeStr = new Date().toLocaleString('vi-VN');

  const wb = new ExcelJS.Workbook();
  wb.creator = 'AI Shrimp Counter';
  wb.created = new Date();

  // ===========================================================================
  // SHEET 1: CÂN TRỌNG LƯỢNG & GIỚI TÍNH (Mẫu tôm thương phẩm)
  // ===========================================================================
  const ws1 = wb.addWorksheet('1_Can_Trong_Luong_Gioi_Tinh', {
    views: [{ showGridLines: true }],
  });

  // PHẦN THÔNG TIN CHUNG (Dòng 1 đến Dòng 5)
  ws1.mergeCells('A1:D1');
  const title1 = ws1.getCell('A1');
  title1.value = 'BÁO CÁO CÂN TRỌNG LƯỢNG, GIỚI TÍNH & ĐẾM TÔM THÔNG MINH';
  title1.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: WHITE_HEX } };
  title1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_HEADER_HEX } };
  title1.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(1).height = 36;

  const metaRows1: Array<[string, string]> = [
    ['Mã Lô / Ao:', pondName || 'Ao Số 1 - Khu A'],
    ['Kỹ thuật viên:', staffName || 'Kỹ thuật viên 1'],
    ['Thời gian xuất báo cáo:', exportTimeStr],
  ];
  metaRows1.forEach(([label, val], idx) => {
    const rIdx = idx + 2;
    ws1.getCell(`A${rIdx}`).value = label;
    ws1.getCell(`A${rIdx}`).font = { name: 'Segoe UI', size: 11, bold: true };
    ws1.mergeCells(`B${rIdx}:D${rIdx}`);
    ws1.getCell(`B${rIdx}`).value = val;
    ws1.getCell(`B${rIdx}`).font = { name: 'Segoe UI', size: 11 };
    ws1.getCell(`B${rIdx}`).alignment = { horizontal: 'left', vertical: 'middle' };
    ws1.getRow(rIdx).height = 20;
  });
  ws1.getRow(5).height = 10;

  // BẢNG THỐNG KÊ TỔNG HỢP (Dòng 6 đến Dòng 12 - Dùng công thức Excel tự động)
  ws1.mergeCells('A6:D6');
  const statHeader = ws1.getCell('A6');
  statHeader.value = 'BẢNG THỐNG KÊ TỔNG HỢP MẪU TÔM (TỰ ĐỘNG TÍNH TOÁN)';
  statHeader.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: WHITE_HEX } };
  statHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_HEADER_HEX } };
  statHeader.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(6).height = 26;

  const detailStartRow1 = 15;
  const detailEndRow1 = Math.max(detailStartRow1, detailStartRow1 + weightRecords.length - 1);

  const totalSamples = weightRecords.length;
  const totalWeightG = weightRecords.reduce((s, r) => s + r.weight, 0);
  const avgWeightG = totalSamples > 0 ? totalWeightG / totalSamples : 0;
  const sizePerKg = totalWeightG > 0 ? (totalSamples / totalWeightG) * 1000 : 0;
  const maleCount = weightRecords.filter((r) => r.gender === 'Male').length;
  const femaleCount = weightRecords.filter((r) => r.gender === 'Female').length;
  const malePct = totalSamples > 0 ? (maleCount / totalSamples) * 100 : 0;
  const femalePct = totalSamples > 0 ? (femaleCount / totalSamples) * 100 : 0;

  const summaryItems1: Array<{
    label: string;
    formula: string;
    result: number;
    numFmt: string;
    unit: string;
  }> = [
    {
      label: 'Tổng số mẫu đã cân',
      formula: `COUNT(B${detailStartRow1}:B${detailEndRow1})`,
      result: totalSamples,
      numFmt: '#,##0',
      unit: 'con',
    },
    {
      label: 'Tổng trọng lượng (g)',
      formula: `SUM(B${detailStartRow1}:B${detailEndRow1})`,
      result: Number(totalWeightG.toFixed(2)),
      numFmt: '#,##0.00',
      unit: 'g',
    },
    {
      label: 'Trọng lượng trung bình',
      formula: `IFERROR(AVERAGE(B${detailStartRow1}:B${detailEndRow1}), 0)`,
      result: Number(avgWeightG.toFixed(2)),
      numFmt: '0.00',
      unit: 'g/con',
    },
    {
      label: 'Size tôm trung bình (con/kg)',
      formula: `IFERROR((C7/C8)*1000, 0)`,
      result: Number(sizePerKg.toFixed(1)),
      numFmt: '0.0',
      unit: 'con/kg',
    },
    {
      label: 'Tỷ lệ Tôm Đực (Male %)',
      formula: `IFERROR((COUNTIF(C${detailStartRow1}:C${detailEndRow1},"Male")/C7)*100, 0)`,
      result: Number(malePct.toFixed(1)),
      numFmt: '0.0',
      unit: '%',
    },
    {
      label: 'Tỷ lệ Tôm Cái (Female %)',
      formula: `IFERROR((COUNTIF(C${detailStartRow1}:C${detailEndRow1},"Female")/C7)*100, 0)`,
      result: Number(femalePct.toFixed(1)),
      numFmt: '0.0',
      unit: '%',
    },
  ];

  summaryItems1.forEach((item, idx) => {
    const rIdx = 7 + idx;
    const isEven = idx % 2 === 1;
    const bgHex = isEven ? ZEBRA_ROW_HEX : WHITE_HEX;

    ws1.mergeCells(`A${rIdx}:B${rIdx}`);
    const lblCell = ws1.getCell(`A${rIdx}`);
    lblCell.value = item.label;
    lblCell.font = { name: 'Segoe UI', size: 11, bold: idx === 3 };

    const valCell = ws1.getCell(`C${rIdx}`);
    valCell.value = { formula: item.formula, result: item.result };
    valCell.numFmt = item.numFmt;
    valCell.font = {
      name: 'Segoe UI',
      size: 11,
      bold: true,
      color: idx === 3 ? { argb: 'FFD35400' } : { argb: 'FF1F4E79' },
    };
    valCell.alignment = { horizontal: 'right', vertical: 'middle' };

    const unitCell = ws1.getCell(`D${rIdx}`);
    unitCell.value = item.unit;
    unitCell.font = { name: 'Segoe UI', size: 11 };
    unitCell.alignment = { horizontal: 'center', vertical: 'middle' };

    ['A', 'B', 'C', 'D'].forEach((col) => {
      const c = ws1.getCell(`${col}${rIdx}`);
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgHex } };
      c.border = THIN_BORDER;
    });
    ws1.getRow(rIdx).height = 22;
  });

  ws1.getRow(13).height = 12;

  // BẢNG CHI TIẾT SHEET 1 (Dòng 14 trở đi: Cột A STT, Cột B Trọng Lượng (g), Cột C Giới Tính, Cột D Thời Gian Nhận Diện)
  const headers1 = ['STT', 'Trọng Lượng (g)', 'Giới Tính', 'Thời Gian Nhận Diện'];
  const headerRow1 = ws1.getRow(14);
  headers1.forEach((text, i) => {
    const cell = headerRow1.getCell(i + 1);
    cell.value = text;
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: WHITE_HEX } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_HEADER_HEX } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = THIN_BORDER;
  });
  headerRow1.height = 26;

  if (weightRecords.length === 0) {
    const emptyRow = ws1.getRow(15);
    emptyRow.getCell(1).value = 1;
    emptyRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    emptyRow.getCell(2).value = 0;
    emptyRow.getCell(2).numFmt = '0.00';
    emptyRow.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };
    emptyRow.getCell(3).value = 'Chưa có mẫu';
    emptyRow.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
    emptyRow.getCell(4).value = '--:--:--';
    emptyRow.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
    for (let c = 1; c <= 4; c++) {
      emptyRow.getCell(c).font = { name: 'Segoe UI', size: 11 };
      emptyRow.getCell(c).border = THIN_BORDER;
    }
  } else {
    weightRecords.forEach((rec, idx) => {
      const rIdx = detailStartRow1 + idx;
      const row = ws1.getRow(rIdx);
      const bgHex = idx % 2 === 1 ? ZEBRA_ROW_HEX : WHITE_HEX;

      row.getCell(1).value = rec.stt;
      row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

      row.getCell(2).value = Number(rec.weight.toFixed(2));
      row.getCell(2).numFmt = '0.00';
      row.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };

      row.getCell(3).value = rec.gender;
      row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };

      row.getCell(4).value = rec.time;
      row.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };

      for (let c = 1; c <= 4; c++) {
        const cell = row.getCell(c);
        cell.font = { name: 'Segoe UI', size: 11 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgHex } };
        cell.border = THIN_BORDER;
      }
      row.height = 21;
    });
  }

  // Dòng Tổng Cộng cuối bảng Sheet 1 (Dùng công thức SUM và AVERAGE)
  const totalRowIdx1 = detailEndRow1 + 1;
  const totalRow1 = ws1.getRow(totalRowIdx1);
  totalRow1.getCell(1).value = 'TỔNG / TB';
  totalRow1.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
  totalRow1.getCell(2).value = {
    formula: `SUM(B${detailStartRow1}:B${detailEndRow1})`,
    result: Number(totalWeightG.toFixed(2)),
  };
  totalRow1.getCell(2).numFmt = '#,##0.00';
  totalRow1.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };
  totalRow1.getCell(3).value = {
    formula: `IFERROR(AVERAGE(B${detailStartRow1}:B${detailEndRow1}), 0)`,
    result: Number(avgWeightG.toFixed(2)),
  };
  totalRow1.getCell(3).numFmt = '"TB: "0.00" g"';
  totalRow1.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
  totalRow1.getCell(4).value = `${totalSamples} mẫu`;
  totalRow1.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };

  for (let c = 1; c <= 4; c++) {
    const cell = totalRow1.getCell(c);
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF1F4E79' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA_ROW_HEX } };
    cell.border = THIN_BORDER;
  }
  totalRow1.height = 24;

  autoFitWorksheetColumns(ws1, 16);

  // ===========================================================================
  // SHEET 2: ĐẾM SỐ LƯỢNG CỘNG DỒN (Mẻ đếm tốc độ cao)
  // ===========================================================================
  const ws2 = wb.addWorksheet('2_Dem_So_Luong_Cong_Don', {
    views: [{ showGridLines: true }],
  });

  // PHẦN THÔNG TIN CHUNG (Dòng 1 đến Dòng 5)
  ws2.mergeCells('A1:F1');
  const title2 = ws2.getCell('A1');
  title2.value = 'BÁO CÁO CÂN TRỌNG LƯỢNG, GIỚI TÍNH & ĐẾM TÔM THÔNG MINH';
  title2.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: WHITE_HEX } };
  title2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_HEADER_HEX } };
  title2.alignment = { horizontal: 'center', vertical: 'middle' };
  ws2.getRow(1).height = 36;

  metaRows1.forEach(([label, val], idx) => {
    const rIdx = idx + 2;
    ws2.getCell(`A${rIdx}`).value = label;
    ws2.getCell(`A${rIdx}`).font = { name: 'Segoe UI', size: 11, bold: true };
    ws2.mergeCells(`B${rIdx}:F${rIdx}`);
    ws2.getCell(`B${rIdx}`).value = val;
    ws2.getCell(`B${rIdx}`).font = { name: 'Segoe UI', size: 11 };
    ws2.getCell(`B${rIdx}`).alignment = { horizontal: 'left', vertical: 'middle' };
    ws2.getRow(rIdx).height = 20;
  });
  ws2.getRow(5).height = 10;

  const detailStartRow2 = 10;
  const detailEndRow2 = Math.max(detailStartRow2, detailStartRow2 + batchRecords.length - 1);
  const validTotalCount = batchRecords
    .filter((b) => b.status === 'Hợp lệ')
    .reduce((s, b) => s + b.count, 0);

  // KHỐI TỔNG SỐ (Dòng 6 đến Dòng 7 - Highlight màu cam nhạt, gộp ô, chữ to in đậm, công thức SUMIF)
  ws2.mergeCells('A6:D7');
  const highlightLbl = ws2.getCell('A6');
  highlightLbl.value = 'TỔNG SỐ LƯỢNG TÔM ĐÃ ĐẾM (CON)';
  highlightLbl.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF9C3400' } };
  highlightLbl.alignment = { horizontal: 'center', vertical: 'middle' };

  ws2.mergeCells('E6:F7');
  const highlightVal = ws2.getCell('E6');
  highlightVal.value = {
    formula: `SUMIF(F${detailStartRow2}:F${detailEndRow2},"Hợp lệ",C${detailStartRow2}:C${detailEndRow2})`,
    result: validTotalCount,
  };
  highlightVal.numFmt = '#,##0 "con"';
  highlightVal.font = { name: 'Segoe UI', size: 18, bold: true, color: { argb: 'FFC0392B' } };
  highlightVal.alignment = { horizontal: 'right', vertical: 'middle' };

  for (let r = 6; r <= 7; r++) {
    ws2.getRow(r).height = 22;
    for (let c = 1; c <= 6; c++) {
      const cell = ws2.getRow(r).getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ORANGE_HIGHLIGHT_HEX } };
      cell.border = THIN_BORDER;
    }
  }

  ws2.getRow(8).height = 10;

  // BẢNG CHI TIẾT SHEET 2 (Dòng 9 trở đi: Cột A STT, B Mã Mẻ, C Số Lượng Mẻ, D Tổng Tích Lũy, E Thời Gian, F Trạng Thái)
  const headers2 = [
    'STT',
    'Mã Mẻ',
    'Số Lượng Mẻ',
    'Tổng Tích Lũy',
    'Thời Gian',
    'Trạng Thái',
  ];
  const headerRow2 = ws2.getRow(9);
  headers2.forEach((text, i) => {
    const cell = headerRow2.getCell(i + 1);
    cell.value = text;
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: WHITE_HEX } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_HEADER_HEX } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = THIN_BORDER;
  });
  headerRow2.height = 26;

  if (batchRecords.length === 0) {
    const emptyRow2 = ws2.getRow(10);
    emptyRow2.getCell(1).value = 1;
    emptyRow2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    emptyRow2.getCell(2).value = 'Mẻ #1';
    emptyRow2.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
    emptyRow2.getCell(3).value = 0;
    emptyRow2.getCell(3).numFmt = '#,##0';
    emptyRow2.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
    emptyRow2.getCell(4).value = { formula: 'SUMIF($F$10:F10,"Hợp lệ",$C$10:C10)', result: 0 };
    emptyRow2.getCell(4).numFmt = '#,##0';
    emptyRow2.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };
    emptyRow2.getCell(5).value = '--:--:--';
    emptyRow2.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
    emptyRow2.getCell(6).value = 'Hợp lệ';
    emptyRow2.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
    for (let c = 1; c <= 6; c++) {
      emptyRow2.getCell(c).font = { name: 'Segoe UI', size: 11 };
      emptyRow2.getCell(c).border = THIN_BORDER;
    }
  } else {
    batchRecords.forEach((rec, idx) => {
      const rIdx = detailStartRow2 + idx;
      const row = ws2.getRow(rIdx);
      const bgHex = idx % 2 === 1 ? ZEBRA_ROW_HEX : WHITE_HEX;

      row.getCell(1).value = rec.id;
      row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

      row.getCell(2).value = rec.batchCode;
      row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

      row.getCell(3).value = rec.count;
      row.getCell(3).numFmt = '#,##0';
      row.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };

      row.getCell(4).value = {
        formula: `SUMIF($F$${detailStartRow2}:F${rIdx},"Hợp lệ",$C$${detailStartRow2}:C${rIdx})`,
        result: rec.cumulative,
      };
      row.getCell(4).numFmt = '#,##0';
      row.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };

      row.getCell(5).value = rec.time;
      row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

      row.getCell(6).value = rec.status;
      row.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };

      for (let c = 1; c <= 6; c++) {
        const cell = row.getCell(c);
        cell.font = {
          name: 'Segoe UI',
          size: 11,
          bold: c === 4,
          strike: rec.status !== 'Hợp lệ',
        };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgHex } };
        cell.border = THIN_BORDER;
      }
      row.height = 21;
    });
  }

  // Dòng Tổng Cộng & Trung Bình Mẻ ở cuối bảng Sheet 2
  const totalRowIdx2 = detailEndRow2 + 1;
  const totalRow2 = ws2.getRow(totalRowIdx2);
  ws2.mergeCells(`A${totalRowIdx2}:B${totalRowIdx2}`);
  totalRow2.getCell(1).value = 'TỔNG CỘNG / TRUNG BÌNH MẺ';
  totalRow2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  totalRow2.getCell(3).value = {
    formula: `SUMIF(F${detailStartRow2}:F${detailEndRow2},"Hợp lệ",C${detailStartRow2}:C${detailEndRow2})`,
    result: validTotalCount,
  };
  totalRow2.getCell(3).numFmt = '#,##0';
  totalRow2.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };

  const validBatchCount = batchRecords.filter((b) => b.status === 'Hợp lệ').length;
  const avgBatch = validBatchCount > 0 ? validTotalCount / validBatchCount : 0;
  totalRow2.getCell(4).value = {
    formula: `IFERROR(AVERAGEIF(F${detailStartRow2}:F${detailEndRow2},"Hợp lệ",C${detailStartRow2}:C${detailEndRow2}), 0)`,
    result: Number(avgBatch.toFixed(1)),
  };
  totalRow2.getCell(4).numFmt = '"TB: "#,##0.0" con/mẻ"';
  totalRow2.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };

  totalRow2.getCell(5).value = `${validBatchCount} mẻ hợp lệ`;
  totalRow2.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
  totalRow2.getCell(6).value = 'Đã đối soát';
  totalRow2.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };

  for (let c = 1; c <= 6; c++) {
    const cell = totalRow2.getCell(c);
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF1F4E79' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA_ROW_HEX } };
    cell.border = THIN_BORDER;
  }
  totalRow2.height = 24;

  autoFitWorksheetColumns(ws2, 16);

  return wb.xlsx.writeBuffer();
}
