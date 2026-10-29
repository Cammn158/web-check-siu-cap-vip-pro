import ExcelJS from 'exceljs';

/**
 * Áp dụng màu chữ đỏ và/hoặc highlight nền đỏ trực tiếp vào 1 ô Excel bị sai/thiếu
 */
function markErrorCell(cell, markStyle = 'both', fallbackRedText = '') {
  const currentVal = cell.value;
  const isEmpty =
    currentVal == null ||
    (typeof currentVal === 'string' && currentVal.trim() === '');

  if (isEmpty && fallbackRedText) {
    cell.value = fallbackRedText;
  }

  const prevFont = cell.font || {};
  if (markStyle === 'both' || markStyle === 'red_text') {
    cell.font = {
      ...prevFont,
      name: prevFont.name || 'Times New Roman',
      size: prevFont.size || 11,
      bold: true,
      color: { argb: 'FFDC2626' }, // Đỏ đậm rõ nét
    };
  }

  if (markStyle === 'both' || markStyle === 'highlight') {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFECACA' }, // Highlight nền đỏ rõ ràng
    };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FFEF4444' } },
      left: { style: 'medium', color: { argb: 'FFEF4444' } },
      bottom: { style: 'medium', color: { argb: 'FFEF4444' } },
      right: { style: 'medium', color: { argb: 'FFEF4444' } },
    };
  }
}

/**
 * Tạo bản sao của File Bản đẹp (CTDT_... hoặc KHGD_...) giữ nguyên 100% định dạng gốc (.xlsx),
 * đánh dấu chữ đỏ / highlight ô màu đỏ trực tiếp tại ô bị sai hoặc bị thiếu,
 * kèm cột Check & Note bên phải bảng.
 */
export async function exportAnnotatedPrettyCopy({
  filePretty,
  prettyRows,
  prettySheetName,
  comparisonResult,
  userNotes = {},
  checkedCourses = {},
  generalUserNote = '',
  customFileName = '',
  options = {
    markStyle: 'both', // 'both' (Chữ đỏ + Highlight ô đỏ) | 'highlight' (Highlight ô đỏ) | 'red_text' (Chữ màu đỏ)
    fillMissingWithRedHint: true, // Điền chữ đỏ [THIẾU: ...] vào ô bị bỏ trống
    includeSummarySheet: true,
  },
}) {
  const wb = new ExcelJS.Workbook();
  let ws = null;
  let loadedOriginalXlsx = false;

  const arrayBuffer = await filePretty.arrayBuffer();
  const isXlsx = /\.xlsx$/i.test(filePretty.name || '');

  if (isXlsx) {
    try {
      await wb.xlsx.load(arrayBuffer);
      ws = wb.getWorksheet(prettySheetName) || wb.worksheets[0];
      loadedOriginalXlsx = Boolean(ws);
    } catch {
      loadedOriginalXlsx = false;
    }
  }

  // Fallback nếu File tải lên là .csv hoặc .xls
  if (!loadedOriginalXlsx) {
    ws = wb.addWorksheet(prettySheetName || 'Ban_Sao_Check_Note');
    prettyRows.forEach((rowArr, rIdx) => {
      const exRow = ws.getRow(rIdx + 1);
      rowArr.forEach((val, cIdx) => {
        exRow.getCell(cIdx + 1).value = val;
      });
    });
    ws.getColumn(1).width = 8;
    ws.getColumn(2).width = 14;
    ws.getColumn(3).width = 52;
    ws.getColumn(4).width = 10;
    ws.getColumn(5).width = 28;
  }

  const {
    fileLabel,
    prettyParsed,
    prettyRowMap,
    missingInPretty,
    stats,
  } = comparisonResult;

  const {
    headerRow = 10,
    codeCol = 1,
    nameCol = 2,
    credCol = 3,
    prereqCol = 4,
  } = prettyParsed.config;

  const headerExcelRow = headerRow + 1;
  let maxTableCol = Math.max(5, prereqCol + 1);

  const hdrRowObj = ws.getRow(headerExcelRow);
  hdrRowObj.eachCell({ includeEmpty: false }, (_cell, colNumber) => {
    if (colNumber > maxTableCol && colNumber <= 15) {
      maxTableCol = colNumber;
    }
  });

  const colStatus = maxTableCol + 1;
  const colAutoDetail = maxTableCol + 2;
  const colCheck = maxTableCol + 3;
  const colUserNote = maxTableCol + 4;

  ws.getColumn(colStatus).width = 22;
  ws.getColumn(colAutoDetail).width = 50;
  ws.getColumn(colCheck).width = 18;
  ws.getColumn(colUserNote).width = 40;

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  if (generalUserNote && generalUserNote.trim()) {
    const topNoteCell = ws.getRow(1).getCell(colStatus);
    topNoteCell.value = `📌 NOTE RÀ SOÁT (${fileLabel}): ${generalUserNote.trim()}`;
    topNoteCell.font = {
      name: 'Times New Roman',
      size: 11,
      bold: true,
      color: { argb: 'FFDC2626' },
    };
  }

  // Tạo tiêu đề cho 4 cột Rà soát & Note bên phải bảng
  const reviewHeaders = [
    { col: colStatus, title: 'Kết quả Rà soát', sub: `(${fileLabel})` },
    {
      col: colAutoDetail,
      title: 'Chi tiết Sai / Thiếu tại File này',
      sub: '(Nhóm, Tín chỉ, Chính tả, Tiên quyết)',
    },
    { col: colCheck, title: 'Tự kiểm tra', sub: '(Checklist)' },
    { col: colUserNote, title: 'Ghi chú sửa (Note)', sub: '(Note màu đỏ để chỉnh sửa)' },
  ];

  for (const h of reviewHeaders) {
    const cell = ws.getRow(headerExcelRow).getCell(h.col);
    cell.value = h.title;
    cell.font = {
      name: 'Times New Roman',
      size: 11,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF991B1B' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = thinBorder;

    const subRowIdx = headerExcelRow + 1;
    if (subRowIdx < prettyParsed.config.startRow + 1) {
      const subCell = ws.getRow(subRowIdx).getCell(h.col);
      subCell.value = h.sub;
      subCell.font = {
        name: 'Times New Roman',
        size: 10,
        italic: true,
        color: { argb: 'FF991B1B' },
      };
      subCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFEE2E2' },
      };
      subCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      subCell.border = thinBorder;
    }
  }

  // Duyệt từng dòng môn học của Bản đẹp: Đánh dấu ĐỎ trực tiếp vào Ô BỊ SAI / THIẾU + điền cột Note chữ đỏ
  for (const [excelRowIdx, info] of prettyRowMap.entries()) {
    const exRow = ws.getRow(excelRowIdx);
    const noteKey = `${info.code}`;
    const isChecked = Boolean(checkedCourses[noteKey]);
    const personalNote = userNotes[noteKey] || '';
    const ce = info.cellErrors || {};

    // 1. Đánh dấu ĐỎ trực tiếp trên đúng ô của bảng gốc bị lỗi/thiếu
    if (ce.codeCol) {
      markErrorCell(exRow.getCell(codeCol + 1), options.markStyle);
    }
    if (ce.viNameCol) {
      markErrorCell(exRow.getCell(nameCol + 1), options.markStyle);
    }
    if (ce.enNameCol) {
      const targetEnRow = info.engExcelRow ? ws.getRow(info.engExcelRow) : exRow;
      markErrorCell(
        targetEnRow.getCell(nameCol + 1),
        options.markStyle,
        options.fillMissingWithRedHint ? '[THIẾU TÊN TIẾNG ANH]' : ''
      );
    }
    if (ce.credCol) {
      markErrorCell(
        exRow.getCell(credCol + 1),
        options.markStyle,
        options.fillMissingWithRedHint ? '[THIẾU TC]' : ''
      );
    }
    if (ce.prereqCol) {
      markErrorCell(
        exRow.getCell(prereqCol + 1),
        options.markStyle,
        options.fillMissingWithRedHint && ce.missingPrereqText
          ? `[THIẾU: ${ce.missingPrereqText}]`
          : ''
      );
    }

    // Nếu người dùng có tự gõ Note cho môn này (kể cả khi máy không bắt lỗi), đánh dấu nhẹ mã môn chữ đỏ để dễ dò
    if (personalNote.trim() && info.severity === 'ok') {
      markErrorCell(exRow.getCell(codeCol + 1), 'red_text');
    }

    // 2. Điền 4 cột Check & Note bên phải
    const cStat = exRow.getCell(colStatus);
    const cDet = exRow.getCell(colAutoDetail);
    const cChk = exRow.getCell(colCheck);
    const cNote = exRow.getCell(colUserNote);

    const hasIssueOrNote = info.severity !== 'ok' || Boolean(personalNote.trim());

    cStat.value = info.statusText;
    cDet.value = `[${info.group || '—'}${info.category ? ` • ${info.category}` : ''}] ${
      info.autoNoteText
    }`;
    cChk.value = isChecked ? '[✓] Đã check' : '[ ] Chưa check';
    cNote.value = personalNote;

    const statusFg =
      info.severity === 'error'
        ? 'FFDC2626'
        : info.severity === 'warn'
        ? 'FFB45309'
        : 'FF166534';
    const statusBg =
      info.severity === 'error'
        ? 'FFFECACA'
        : info.severity === 'warn'
        ? 'FFFEF3C7'
        : 'FFDCFCE7';

    cStat.font = {
      name: 'Times New Roman',
      size: 10,
      bold: true,
      color: { argb: statusFg },
    };
    cStat.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: statusBg } };
    cStat.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cStat.border = thinBorder;

    cDet.font = {
      name: 'Times New Roman',
      size: 10,
      bold: info.severity === 'error',
      color: { argb: info.severity === 'ok' ? 'FF334155' : statusFg },
    };
    if (info.severity !== 'ok') {
      cDet.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: statusBg } };
    }
    cDet.alignment = { vertical: 'middle', wrapText: true };
    cDet.border = thinBorder;

    cChk.font = {
      name: 'Times New Roman',
      size: 10,
      bold: isChecked,
      color: { argb: isChecked ? 'FF166534' : 'FF475569' },
    };
    cChk.alignment = { vertical: 'middle', horizontal: 'center' };
    cChk.border = thinBorder;

    // Cột Note cá nhân luôn hiển thị chữ ĐỎ ĐẬM trên nền hồng/vàng nhạt khi có nội dung note
    cNote.font = {
      name: 'Times New Roman',
      size: 10.5,
      bold: true,
      color: { argb: 'FFDC2626' },
    };
    if (hasIssueOrNote) {
      cNote.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: personalNote.trim() ? 'FFFEE2E2' : 'FFFFF1F2' },
      };
    }
    cNote.alignment = { vertical: 'middle', wrapText: true };
    cNote.border = thinBorder;

    if (info.engExcelRow) {
      const enRow = ws.getRow(info.engExcelRow);
      [colStatus, colAutoDetail, colCheck, colUserNote].forEach((cIdx) => {
        enRow.getCell(cIdx).border = thinBorder;
      });
    }
  }

  // 3. Nếu Bản đẹp bị SÓT MÔN hoàn toàn so với File gốc MyBK: chèn khối ĐỎ ĐẬM cuối bảng
  if (missingInPretty.length > 0) {
    let appendRowIdx = ws.rowCount + 2;
    const warnHeaderRow = ws.getRow(appendRowIdx++);
    warnHeaderRow.getCell(1).value = `⚠️ CÁC MÔN CÓ TRONG GỐC MYBK NHƯNG ĐANG BỊ THIẾU (SÓT) TRONG ${fileLabel.toUpperCase()}:`;
    warnHeaderRow.getCell(1).font = {
      name: 'Times New Roman',
      size: 11,
      bold: true,
      color: { argb: 'FFDC2626' },
    };

    missingInPretty.forEach((m, i) => {
      const r = ws.getRow(appendRowIdx++);
      r.getCell(1).value = i + 1;
      r.getCell(codeCol + 1).value = m.code;
      r.getCell(nameCol + 1).value = `[THIẾU MÔN] ${m.name}${
        m.eng ? ` / ${m.eng}` : ''
      }`;
      r.getCell(credCol + 1).value = m.credits;
      r.getCell(prereqCol + 1).value = m.prereq || '';
      r.getCell(colStatus).value = `✗ THIẾU Ở ${fileLabel.toUpperCase()}`;
      r.getCell(colAutoDetail).value = `Cần bổ sung vào ${m.group || '—'} [${
        m.category || '—'
      }]`;
      r.getCell(colCheck).value = '[ ] Bổ sung ngay';
      r.getCell(colUserNote).value = userNotes[m.code] || 'Bổ sung môn bị sót';

      for (let c = 1; c <= colUserNote; c++) {
        const cell = r.getCell(c);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFECACA' },
        };
        cell.font = {
          name: 'Times New Roman',
          size: 10.5,
          bold: true,
          color: { argb: 'FFDC2626' },
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFEF4444' } },
          left: { style: 'thin', color: { argb: 'FFEF4444' } },
          bottom: { style: 'thin', color: { argb: 'FFEF4444' } },
          right: { style: 'thin', color: { argb: 'FFEF4444' } },
        };
      }
    });
  }

  // 4. Sheet phụ Tổng hợp Rà soát của riêng File Bản đẹp này
  if (options.includeSummarySheet) {
    let sumSheetName = 'Tong_Hop_Loi_File_Nay';
    if (wb.getWorksheet(sumSheetName)) {
      sumSheetName = 'Check_Note_Summary';
    }
    const sumWs = wb.addWorksheet(sumSheetName);
    sumWs.columns = [
      { header: 'Bước / Hạng mục', key: 'step', width: 24 },
      { header: 'Dòng Excel', key: 'row', width: 12 },
      { header: 'Mã môn / Nhóm', key: 'code', width: 16 },
      { header: 'Tên môn / Hạng mục', key: 'name', width: 44 },
      { header: `Chi tiết sai / thiếu trên ${fileLabel}`, key: 'detail', width: 54 },
      { header: 'Ghi chú sửa (Note)', key: 'note', width: 36 },
    ];

    const hdr = sumWs.getRow(1);
    hdr.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF991B1B' },
      };
    });

    const addRedSummaryRow = (obj) => {
      const added = sumWs.addRow(obj);
      added.getCell('detail').font = { color: { argb: 'FFDC2626' }, bold: true };
      if (obj.note) {
        added.getCell('note').font = { color: { argb: 'FFDC2626' }, bold: true };
      }
    };

    if (generalUserNote) {
      addRedSummaryRow({
        step: 'GHI CHÚ CHUNG',
        row: '-',
        code: 'ALL',
        name: `Tổng quan ${fileLabel}`,
        detail: `Tổng dòng kiểm: ${stats.total} | Khớp: ${stats.match} | Lỗi tại file này: ${stats.totalFileErrors}`,
        note: generalUserNote,
      });
    }

    for (const [exRow, info] of prettyRowMap.entries()) {
      const pNote = userNotes[info.code] || '';
      if (info.severity !== 'ok' || pNote.trim()) {
        addRedSummaryRow({
          step: info.severity === 'error' ? 'CẦN SỬA TRÊN BẢN ĐẸP' : 'LƯU Ý RÀ SOÁT',
          row: `Dòng ${exRow}`,
          code: info.code,
          name: info.name,
          detail: info.autoNoteText,
          note: pNote,
        });
      }
    }

    missingInPretty.forEach((m) => {
      addRedSummaryRow({
        step: 'B2 - BỊ SÓT MÔN',
        row: 'Cuối bảng',
        code: m.code,
        name: m.name,
        detail: `Thiếu môn này trong ${fileLabel} (Nhóm gốc: ${m.group})`,
        note: userNotes[m.code] || 'Cần bổ sung',
      });
    });
  }

  const outBuffer = await wb.xlsx.writeBuffer();
  const baseName = (filePretty.name || 'Ban_Dep')
    .replace(/\.(xlsx|xls|csv)$/i, '')
    .trim();
  const finalName =
    customFileName && customFileName.trim()
      ? customFileName.trim().endsWith('.xlsx')
        ? customFileName.trim()
        : `${customFileName.trim()}.xlsx`
      : `${baseName}_check_note_do.xlsx`;

  const blob = new Blob([outBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = finalName;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 5000);
}
