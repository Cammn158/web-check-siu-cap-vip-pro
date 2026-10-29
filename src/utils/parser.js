import * as XLSX from 'xlsx';
import { checkCourseSpelling } from './spellcheck.js';

export const clean = (val) =>
  val == null ? '' : String(val).trim().replace(/\s+/g, ' ');

export function sheetToRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) : null;
}

const ROMAN_TO_NUM = {
  I: 1,
  II: 2,
  III: 3,
  IV: 4,
  V: 5,
  VI: 6,
  VII: 7,
  VIII: 8,
};

/**
 * Nhận diện mã môn học:
 * - Chuẩn: 2 chữ cái + >= 4 chữ số (VD: AS5113, EE5359, ME6055...)
 * - Mã điều kiện/đặc biệt: chữ in hoa + gạch dưới + chữ/số (VD: ENG_B2...)
 */
export function isCourseCode(str) {
  const s = clean(str);
  if (!s) return false;
  return /^(?:[A-Z]{2}\d{4,}|[A-Z]{2,}_[A-Z0-9]+)$/.test(s);
}

/**
 * Nhận diện mã nhóm kiến thức (A, A.1, A.2, A.3, B, C.1...) hoặc học kỳ (HK1, Học kỳ 1, Học kỳ I, KHGD_..._HK1)
 */
export function parseGroupId(cellText, mode) {
  const t = clean(cellText);
  if (!t) return null;
  if (mode === 'khgd') {
    const m1 = t.match(/^(?:[IVX]+\.\s*)?HK\s*(\d+)\b/i);
    if (m1) return `HK${m1[1]}`;
    const m2 = t.match(/^(?:[IVX]+\.\s*)?H[oọ]c\s*k[ỳì]\s*(\d+)\b/i);
    if (m2) return `HK${m2[1]}`;
    const mRoman = t.match(/^(?:[IVX]+\.\s*)?H[oọ]c\s*k[ỳì]\s*([IVX]+)\b/i);
    if (mRoman && ROMAN_TO_NUM[mRoman[1].toUpperCase()]) {
      return `HK${ROMAN_TO_NUM[mRoman[1].toUpperCase()]}`;
    }
    const m3 = t.match(/(?:^|[_-\s])HK\s*(\d+)\b/i);
    if (m3) return `HK${m3[1]}`;
    return null;
  }
  const m = t.match(/^([A-E])(?:[.\s]?(\d+))?[.\s]/);
  return m ? (m[2] ? `${m[1]}.${m[2]}` : m[1]) : null;
}

export function cleanGroupLabel(str) {
  return clean(str)
    .replace(/\[(BB|TC|bb|tc)\]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Suy luận tính chất Bắt buộc (BB) / Tự chọn (TC) từ tiêu đề nhóm trong Khối kiến thức
 */
export function inferGroupCategory(rawText, gid) {
  const text = clean(rawText);
  const explicitMatch = text.match(/\[(BB|TC)\]/i);
  const explicitTag = explicitMatch ? explicitMatch[1].toUpperCase() : null;

  const lower = text.toLowerCase();
  let inferredByKeyword = null;
  if (
    lower.includes('tự chọn') ||
    lower.includes('elective') ||
    lower.includes('optional')
  ) {
    inferredByKeyword = 'TC';
  } else if (
    lower.includes('bắt buộc') ||
    lower.includes('compulsory') ||
    lower.includes('cơ sở ngành') ||
    lower.includes('core courses') ||
    lower.includes('foundation') ||
    lower.includes('ngoại ngữ') ||
    lower.includes('foreign language') ||
    lower.includes('tốt nghiệp') ||
    lower.includes('graduation')
  ) {
    inferredByKeyword = 'BB';
  }

  const isParentContainer = gid === 'A' || gid === 'C';
  if (isParentContainer) {
    return {
      category: 'PARENT',
      explicitTag,
      inferredByKeyword,
    };
  }

  const fallbackByGid =
    gid === 'A.2' || gid === 'C.2' ? 'TC' : gid ? 'BB' : null;

  return {
    category: explicitTag || inferredByKeyword || fallbackByGid || 'BB',
    explicitTag,
    inferredByKeyword,
  };
}

/**
 * Nhận diện tiêu đề phân mục Bắt buộc / Tự chọn trong Kế hoạch giảng dạy
 */
function detectKhgdSectionCategory(c0, c1, c2) {
  if (c1 && isCourseCode(c1)) return null;
  if (/^\d+(\.\d+)?$/.test(c0)) return null;
  const combined = `${c0} ${c1} ${c2}`.trim().toLowerCase();
  if (!combined) return null;
  if (
    combined.includes('các học phần bắt buộc') ||
    combined.includes('học phần bắt buộc') ||
    /^(?:[a-z]|[ivx]+)[.)]\s*(?:các\s*)?học phần bắt buộc/i.test(combined) ||
    /\(compulsory courses\)/i.test(combined)
  ) {
    return 'BB';
  }
  if (
    combined.includes('các học phần tự chọn') ||
    combined.includes('học phần tự chọn') ||
    /^(?:[a-z]|[ivx]+)[.)]\s*(?:các\s*)?học phần tự chọn/i.test(combined) ||
    /\(elective courses\)/i.test(combined)
  ) {
    return 'TC';
  }
  return null;
}

/**
 * Tìm dòng bắt đầu dữ liệu, vị trí cột Học phần điều kiện & Ghi chú, đồng thời trích xuất metadata đầu file
 */
export function detectTableConfig(rows) {
  let startRow = -1;
  let headerRow = -1;
  let codeCol = 1;
  let nameCol = 2;
  let credCol = 3;
  let prereqCol = 4; // Mặc định cột E
  let noteCol = 5; // Mặc định cột F

  for (let r = 0; r < rows.length; r++) {
    const c0 = clean(rows[r][0]);
    const c1 = clean(rows[r][1]);
    if (
      c0 === 'STT' ||
      c1 === 'Mã môn học' ||
      c1 === 'Mã học phần' ||
      /^Mã\s*học\s*phần$/i.test(c1)
    ) {
      headerRow = r;
      for (let col = 0; col < rows[r].length; col++) {
        const headerText = clean(rows[r][col]).toLowerCase();
        if (
          headerText.includes('điều kiện') ||
          headerText.includes('tiên quyết') ||
          headerText.includes('song hành') ||
          headerText.includes('prerequisite') ||
          headerText.includes('condition')
        ) {
          prereqCol = col;
        } else if (
          headerText.includes('ghi chú') ||
          headerText.includes('note')
        ) {
          noteCol = col;
        }
      }
      let next = r + 1;
      while (next < rows.length) {
        const sub0 = clean(rows[next][0]);
        if (sub0 === '(No.)' || sub0 === 'No.') next++;
        else break;
      }
      startRow = next;
      break;
    }
  }

  if (startRow === -1) {
    for (let r = 0; r < rows.length; r++) {
      if (isCourseCode(rows[r][1])) {
        startRow = r;
        headerRow = Math.max(0, r - 1);
        break;
      }
    }
  }

  const finalStartRow = startRow === -1 ? 38 : startRow;

  let declaredTotalCredits = null;
  let programTitle = '';
  let legendNoteText = '';
  let generalNoteText = '';
  const legendDefinedTypes = new Set();

  for (let r = 0; r < Math.min(finalStartRow, rows.length); r++) {
    const lineJoined = rows[r].map((c) => clean(c)).filter(Boolean).join(' ');
    if (!lineJoined) continue;

    if (
      !programTitle &&
      (/chương trình đào tạo/i.test(lineJoined) ||
        /kế hoạch giảng dạy/i.test(lineJoined))
    ) {
      programTitle = lineJoined;
    }

    const credMatch = lineJoined.match(/(\d+)\s*tín\s*chỉ/i);
    if (credMatch && declaredTotalCredits == null) {
      declaredTotalCredits = Number(credMatch[1]);
    }

    if (/ghi\s*chú\s*\(notes?\)\s*:/i.test(lineJoined) || /\bTQ\s*:/i.test(lineJoined)) {
      legendNoteText = lineJoined;
      for (const token of ['TQ', 'HT', 'SHT', 'SH', 'KN', 'BB', 'TC']) {
        const reg = new RegExp(`\\b${token}\\s*:`, 'i');
        if (reg.test(lineJoined)) {
          legendDefinedTypes.add(token);
        }
      }
    } else if (/^ghi\s*chú/i.test(clean(rows[r][0]))) {
      generalNoteText = clean(rows[r][1]);
    }
  }

  return {
    headerRow: headerRow === -1 ? Math.max(0, finalStartRow - 2) : headerRow,
    startRow: finalStartRow,
    codeCol,
    nameCol,
    credCol,
    prereqCol,
    noteCol,
    meta: {
      programTitle,
      declaredTotalCredits,
      legendNoteText,
      legendDefinedTypes,
      generalNoteText,
    },
  };
}

function isSeparatorRow(c0, c1, c2) {
  if (clean(c1)) return false;
  return Boolean((!clean(c0) && !clean(c2)) || (clean(c0) && !clean(c1) && clean(c2)));
}

/**
 * Phân tích chuỗi học phần điều kiện / tiên quyết
 */
export function parsePrereq(raw) {
  const s = clean(raw);
  if (!s || s === '--' || s === '-' || /^không$/i.test(s) || /^none$/i.test(s)) {
    return [];
  }

  const parts = s
    .split(/[,;\n]+/)
    .map((p) => clean(p))
    .filter(Boolean);

  const items = [];
  for (const part of parts) {
    const m = part.match(/^([A-Za-z0-9_]+)\s*(?:\(\s*([^)]+?)\s*\))?$/);
    if (m) {
      const code = m[1].toUpperCase();
      const type = m[2] ? m[2].replace(/\s+/g, '').toUpperCase() : '';
      const norm = type ? `${code}(${type})` : code;
      items.push({ code, type, norm, raw: part });
    } else {
      const norm = part.replace(/\s+/g, '').toUpperCase();
      items.push({ code: norm, type: '', norm, raw: part });
    }
  }
  return items;
}

/**
 * So sánh học phần điều kiện giữa 2 nguồn
 */
export function comparePrereq(origRaw, prettyRaw) {
  const origItems = parsePrereq(origRaw);
  const prettyItems = parsePrereq(prettyRaw);

  if (origItems.length === 0 && prettyItems.length === 0) {
    return {
      match: true,
      origItems,
      prettyItems,
      missing: [],
      extra: [],
      wrongType: [],
      shtVsShOnly: false,
      summary: '',
    };
  }

  const origByCode = new Map();
  for (const item of origItems) origByCode.set(item.code, item);

  const prettyByCode = new Map();
  for (const item of prettyItems) prettyByCode.set(item.code, item);

  const missing = [];
  const extra = [];
  const wrongType = [];

  for (const [code, oItem] of origByCode.entries()) {
    if (!prettyByCode.has(code)) {
      missing.push(oItem.norm);
    } else {
      const pItem = prettyByCode.get(code);
      if (oItem.type !== pItem.type) {
        wrongType.push({
          code,
          origType: oItem.type || 'Không ghi',
          prettyType: pItem.type || 'Không ghi',
          origNorm: oItem.norm,
          prettyNorm: pItem.norm,
        });
      }
    }
  }

  for (const [code, pItem] of prettyByCode.entries()) {
    if (!origByCode.has(code)) {
      extra.push(pItem.norm);
    }
  }

  const match = missing.length === 0 && extra.length === 0 && wrongType.length === 0;
  const shtVsShOnly =
    missing.length === 0 &&
    extra.length === 0 &&
    wrongType.length > 0 &&
    wrongType.every(
      (w) =>
        (w.origType === 'SHT' && w.prettyType === 'SH') ||
        (w.origType === 'SH' && w.prettyType === 'SHT')
    );

  const summaryParts = [];
  if (missing.length > 0) summaryParts.push(`Thiếu: ${missing.join(', ')}`);
  if (extra.length > 0) summaryParts.push(`Thừa: ${extra.join(', ')}`);
  if (wrongType.length > 0) {
    summaryParts.push(
      `Lệch loại ĐK: ${wrongType
        .map((w) => `${w.code} (Gốc: ${w.origType} → Đẹp: ${w.prettyType})`)
        .join(', ')}`
    );
  }

  return {
    match,
    origItems,
    prettyItems,
    missing,
    extra,
    wrongType,
    shtVsShOnly,
    summary: summaryParts.join(' | '),
  };
}

/**
 * Đọc cấu trúc bảng từ 1 sheet (lưu cả thông tin từng dòng xuất hiện cho KHGD nhiều học kỳ)
 */
export function parseSheetData(rows, mode) {
  const codeToGroup = new Map();
  const codeToGroups = new Map();
  const codeToCategory = new Map(); // 'BB' | 'TC'
  const groups = new Map();
  const lookup = new Map();
  const occurrences = []; // Mỗi dòng môn học thực tế trên Excel (quan trọng cho KHGD khi 1 môn xuất hiện ở cả HK3 & HK4)
  const orderedCodes = [];
  const khgdMisplacedSubgroups = [];
  const parentTagWarnings = [];

  let currentGroup = null;
  let currentCategory = null;
  let currentKhgdSubgroup = '';

  const config = detectTableConfig(rows);
  const { startRow, prereqCol, noteCol, meta } = config;

  for (let r = startRow; r < rows.length; r++) {
    const row = rows[r];
    const c0 = clean(row[0]);
    const c1 = clean(row[1]);
    const c2 = clean(row[2]);
    const c3 = row[3];
    const cPrereq = clean(row[prereqCol]);
    const cNote = row[noteCol] != null ? clean(row[noteCol]) : '';

    let gid = parseGroupId(c0, mode);
    let gText = c0;
    let groupDeclaredCol = 3; // Mặc định cột D
    if (!gid && !c0) {
      const altGid = parseGroupId(c1, mode);
      if (altGid) {
        gid = altGid;
        gText = c1;
      }
    }

    if (gid) {
      const decCredits =
        c3 !== '' && !isNaN(Number(c3)) && Number(c3) >= 0 ? Number(c3) : null;
      currentGroup = gid;
      currentKhgdSubgroup = '';

      if (mode === 'ctdt') {
        const catInfo = inferGroupCategory(gText, gid);
        if (catInfo.category !== 'PARENT') {
          currentCategory = catInfo.category;
        } else {
          if (catInfo.explicitTag === 'TC') {
            parentTagWarnings.push({
              gid,
              row: r + 1,
              rawTitle: gText,
              detail: `Nhóm lớn "${gid}" trên File gốc MyBK đang gắn nhãn [TC] (Tự chọn), trong khi chứa các nhóm con Bắt buộc [BB] bên trong.`,
            });
          }
        }

        if (groups.has(gid)) {
          if (decCredits != null && groups.get(gid).creditsDeclared == null) {
            groups.get(gid).creditsDeclared = decCredits;
          }
        } else {
          groups.set(gid, {
            gid,
            excelRow: r + 1,
            declaredCol: groupDeclaredCol,
            label: cleanGroupLabel(gText),
            rawLabel: gText,
            category: catInfo.category,
            explicitTag: catInfo.explicitTag,
            creditsSum: 0,
            courseCount: 0,
            creditsDeclared: decCredits,
            courses: [],
          });
        }
      } else {
        currentCategory = null;
        if (groups.has(gid)) {
          if (decCredits != null && groups.get(gid).creditsDeclared == null) {
            groups.get(gid).creditsDeclared = decCredits;
          }
        } else {
          groups.set(gid, {
            gid,
            excelRow: r + 1,
            declaredCol: groupDeclaredCol,
            label: cleanGroupLabel(gText),
            rawLabel: gText,
            category: 'HK',
            creditsSum: 0,
            compulsoryCreditsSum: 0,
            electiveCreditsSum: 0,
            courseCount: 0,
            creditsDeclared: decCredits,
            courses: [],
          });
        }
      }
      continue;
    }

    if (mode === 'khgd') {
      const secCat = detectKhgdSectionCategory(c0, c1, c2);
      if (secCat) {
        currentCategory = secCat;
        currentKhgdSubgroup = '';
        continue;
      }

      if (/^\d+$/.test(c0) && !c1 && c2) {
        currentKhgdSubgroup = c2;
        if (
          currentCategory === 'TC' &&
          (/bắt buộc|compulsory/i.test(c2) || /^cơ sở ngành$/i.test(c2))
        ) {
          khgdMisplacedSubgroups.push({
            semester: currentGroup || '—',
            row: r + 1,
            subgroupName: c2,
            currentSection: 'Các học phần tự chọn (TC)',
            detail: `Mục "${c2}" mang tính chất Bắt buộc nhưng đang bị xếp dưới phần "Các học phần tự chọn" ở ${
              currentGroup || 'KHGD'
            }.`,
          });
        }
        continue;
      }
    }

    if (c1 && isCourseCode(c1)) {
      const credits = c3 !== '' && !isNaN(Number(c3)) ? Number(c3) : null;
      const rawVi = row[2] != null ? String(row[2]) : '';
      let engName = '';
      let rawEn = '';
      let engExcelRow = null;

      if (r + 1 < rows.length) {
        const nextRow = rows[r + 1];
        if (!clean(nextRow[0]) && !clean(nextRow[1]) && clean(nextRow[2])) {
          engName = clean(nextRow[2]);
          rawEn = String(nextRow[2]);
          engExcelRow = r + 2;
        }
      }

      const occ = {
        code: c1,
        excelRow: r + 1,
        engExcelRow,
        name: c2,
        rawVi,
        credits,
        rawCredits: c3,
        eng: engName,
        rawEn,
        prereq: cPrereq,
        prereqItems: parsePrereq(cPrereq),
        note: cNote,
        group: currentGroup || '',
        category: currentCategory || '',
        khgdSubgroup: currentKhgdSubgroup,
      };
      occurrences.push(occ);

      if (!lookup.has(c1)) {
        orderedCodes.push(c1);
      }

      const prevEntry = lookup.get(c1);
      lookup.set(c1, {
        ...occ,
        excelRow: prevEntry ? prevEntry.excelRow : r + 1,
        allExcelRows: prevEntry ? [...prevEntry.allExcelRows, r + 1] : [r + 1],
        engExcelRow: prevEntry?.engExcelRow || engExcelRow,
        group: prevEntry?.group || currentGroup || '',
        category: currentCategory || prevEntry?.category || '',
      });

      if (currentCategory && !codeToCategory.has(c1)) {
        codeToCategory.set(c1, currentCategory);
      }

      if (currentGroup) {
        if (!codeToGroup.has(c1)) {
          codeToGroup.set(c1, currentGroup);
        }
        if (!codeToGroups.has(c1)) {
          codeToGroups.set(c1, new Set());
        }
        codeToGroups.get(c1).add(currentGroup);

        const grpObj = groups.get(currentGroup);
        if (grpObj) {
          grpObj.creditsSum += credits || 0;
          if (currentCategory === 'BB') {
            grpObj.compulsoryCreditsSum =
              (grpObj.compulsoryCreditsSum || 0) + (credits || 0);
          } else if (currentCategory === 'TC') {
            grpObj.electiveCreditsSum =
              (grpObj.electiveCreditsSum || 0) + (credits || 0);
          }
          grpObj.courseCount++;
          grpObj.courses.push(c1);
        }
      }
    }
  }

  return {
    config,
    meta,
    codeToGroup,
    codeToGroups,
    codeToCategory,
    groups,
    lookup,
    occurrences,
    orderedCodes,
    khgdMisplacedSubgroups,
    parentTagWarnings,
  };
}

/**
 * Đối chiếu từng dòng của 1 sheet Gốc MyBK với 1 sheet Bản đẹp (CTĐT hoặc KHGD)
 */
function compareRows(origRows, origParsed, prettyParsed, mode, fileLabel) {
  const {
    lookup: prettyLookup,
    codeToGroups: prettyCodeToGroups,
    codeToCategory: prettyCodeToCategory,
  } = prettyParsed;

  const rowResults = [];
  let lastCourseCode = null;
  let currentGroup = null;
  let currentCategory = null;

  const { startRow, prereqCol, noteCol } = origParsed.config;

  for (let r = startRow; r < origRows.length; r++) {
    const row = origRows[r];
    const c0 = clean(row[0]);
    const c1 = clean(row[1]);
    const c2 = clean(row[2]);
    const c3 = row[3];
    const cPrereq = clean(row[prereqCol]);
    const cNote = row[noteCol] != null ? clean(row[noteCol]) : '';

    let gid = parseGroupId(c0, mode);
    let gText = c0;
    if (!gid && !c0) {
      const altGid = parseGroupId(c1, mode);
      if (altGid) {
        gid = altGid;
        gText = c1;
      }
    }

    if (gid) {
      currentGroup = gid;
      if (mode === 'ctdt') {
        const catInfo = inferGroupCategory(gText, gid);
        if (catInfo.category !== 'PARENT') currentCategory = catInfo.category;
      } else {
        currentCategory = null;
      }

      rowResults.push({
        row: r + 1,
        type: 'group',
        groupId: gid,
        category: currentCategory,
        name: cleanGroupLabel(gText),
        code: '',
        credits: '',
        prereq: '',
        status: 'group',
        fileLabel,
      });
      continue;
    }

    if (mode === 'khgd') {
      const secCat = detectKhgdSectionCategory(c0, c1, c2);
      if (secCat) {
        currentCategory = secCat;
        rowResults.push({
          row: r + 1,
          type: 'separator',
          name: `${c0 || c1 || c2} [${secCat}]`,
          code: '',
          credits: '',
          prereq: '',
          status: 'skip',
          fileLabel,
        });
        continue;
      }
    }

    if (!c1 && !c2) continue;

    if (isSeparatorRow(c0, c1, c2)) {
      if (c2 && !/^\d+(\.\d+)?$/.test(c0)) {
        rowResults.push({
          row: r + 1,
          type: 'separator',
          name: c2,
          code: '',
          credits: '',
          prereq: '',
          status: 'skip',
          fileLabel,
        });
      }
      continue;
    }

    if (c1 && isCourseCode(c1)) {
      lastCourseCode = c1;
      const pEntry = prettyLookup.get(c1);
      const origGrp = currentGroup || origParsed.codeToGroup.get(c1) || '';
      const pGroupsSet = prettyCodeToGroups.get(c1);
      const prettyGrpList = pGroupsSet ? Array.from(pGroupsSet) : [];
      const prettyGrp = prettyGrpList.includes(origGrp)
        ? origGrp
        : prettyGrpList.join(', ');
      const groupMismatch = Boolean(
        origGrp && prettyGrpList.length > 0 && !prettyGrpList.includes(origGrp)
      );

      const origCat = currentCategory || origParsed.codeToCategory.get(c1) || '';
      const prettyCat = prettyCodeToCategory.get(c1) || '';
      const categoryMismatch = Boolean(origCat && prettyCat && origCat !== prettyCat);

      const prereqDiff = comparePrereq(cPrereq, pEntry ? pEntry.prereq : '');

      if (!pEntry) {
        rowResults.push({
          row: r + 1,
          type: 'vi',
          code: c1,
          name: c2,
          credits: c3,
          prereq: cPrereq,
          prereqItems: parsePrereq(cPrereq),
          note: cNote,
          prereqDiff,
          prereqMismatch: false,
          origCat,
          prettyCat,
          categoryMismatch: false,
          status: 'miss',
          label: `${fileLabel} bị thiếu môn này`,
          origGrp,
          prettyGrp,
          groupMismatch,
          fileLabel,
        });
      } else {
        const nameOk = clean(pEntry.name) === clean(c2);
        const credNum = c3 !== '' && !isNaN(Number(c3)) ? Number(c3) : null;
        const credOk = credNum === pEntry.credits;
        const prereqOk = prereqDiff.match;

        if (nameOk && credOk && prereqOk) {
          rowResults.push({
            row: r + 1,
            type: 'vi',
            code: c1,
            name: c2,
            credits: c3,
            prereq: cPrereq,
            prereqItems: parsePrereq(cPrereq),
            note: cNote,
            prereqDiff,
            prereqMismatch: false,
            origCat,
            prettyCat,
            categoryMismatch,
            status: 'match',
            label: 'Khớp 100%',
            origGrp,
            prettyGrp,
            groupMismatch,
            pEntry,
            fileLabel,
          });
        } else if (nameOk && credOk && !prereqOk) {
          rowResults.push({
            row: r + 1,
            type: 'vi',
            code: c1,
            name: c2,
            credits: c3,
            prereq: cPrereq,
            prereqItems: parsePrereq(cPrereq),
            note: cNote,
            prereqDiff,
            prereqMismatch: true,
            origCat,
            prettyCat,
            categoryMismatch,
            status: 'wrong-prereq',
            label: `Sai ĐK tiên quyết ở ${fileLabel} (${prereqDiff.summary})`,
            origGrp,
            prettyGrp,
            groupMismatch,
            pEntry,
            fileLabel,
          });
        } else {
          const errParts = [];
          if (!nameOk) errParts.push('Sai tên TV');
          if (!credOk) errParts.push('Sai TC');
          if (!prereqOk) errParts.push('Sai ĐK tiên quyết');

          rowResults.push({
            row: r + 1,
            type: 'vi',
            code: c1,
            name: c2,
            credits: c3,
            prereq: cPrereq,
            prereqItems: parsePrereq(cPrereq),
            note: cNote,
            prereqDiff,
            prereqMismatch: !prereqOk,
            origCat,
            prettyCat,
            categoryMismatch,
            status: 'wrong',
            label: `Sai ở ${fileLabel} (${errParts.join(' + ')})`,
            origGrp,
            prettyGrp,
            groupMismatch,
            pEntry,
            fileLabel,
          });
        }
      }
      continue;
    }

    if (!c0 && !c1 && c2 && lastCourseCode) {
      const pEntry = prettyLookup.get(lastCourseCode);
      if (!pEntry) {
        rowResults.push({
          row: r + 1,
          type: 'en',
          code: '',
          parentCode: lastCourseCode,
          name: c2,
          credits: '',
          prereq: '',
          status: 'na',
          label: '#N/A (dòng tiếng Việt phía trên bị thiếu)',
          fileLabel,
        });
      } else {
        const engOk = clean(pEntry.eng) === clean(c2);
        rowResults.push({
          row: r + 1,
          type: 'en',
          code: '',
          parentCode: lastCourseCode,
          name: c2,
          credits: '',
          prereq: '',
          status: engOk ? 'match' : 'wrong-en',
          label: engOk ? 'Khớp 100%' : `Sai tên tiếng Anh ở ${fileLabel}`,
          pEntry,
          fileLabel,
        });
      }
    }
  }

  return rowResults;
}

/**
 * Rà soát nội bộ trên chính File 1 (Gốc MyBK) để tách riêng những điểm bất cập từ hệ thống MyBK (Dùng để báo Cô)
 */
export function runMybkInternalAudit(origCtdtRows, origKhgdRows) {
  const ctdtParsed = origCtdtRows ? parseSheetData(origCtdtRows, 'ctdt') : null;
  const khgdParsed = origKhgdRows ? parseSheetData(origKhgdRows, 'khgd') : null;

  const bbInTcIssues = [];
  const missingInKhgd = [];
  const missingInCtdt = [];
  const prereqCrossIssues = [];
  const mybkTypos = [];
  const parentTagWarnings = ctdtParsed?.parentTagWarnings || [];
  const khgdMisplacedSubgroups = khgdParsed?.khgdMisplacedSubgroups || [];

  if (ctdtParsed && khgdParsed) {
    for (const [code, cEntry] of ctdtParsed.lookup.entries()) {
      const kEntry = khgdParsed.lookup.get(code);
      if (!kEntry) {
        missingInKhgd.push({
          code,
          name: cEntry.name,
          credits: cEntry.credits,
          ctdtGroup: cEntry.group,
          ctdtCategory: cEntry.category,
        });
        continue;
      }

      if (cEntry.category && kEntry.category && cEntry.category !== kEntry.category) {
        const kSemesters = Array.from(khgdParsed.codeToGroups.get(code) || []).join(', ');
        bbInTcIssues.push({
          code,
          name: cEntry.name,
          credits: cEntry.credits,
          ctdtGroup: cEntry.group,
          ctdtCategory: cEntry.category,
          khgdSemester: kSemesters,
          khgdCategory: kEntry.category,
          khgdSubgroup: kEntry.khgdSubgroup || '',
        });
      }

      const pCmp = comparePrereq(cEntry.prereq, kEntry.prereq);
      if (!pCmp.match) {
        prereqCrossIssues.push({
          code,
          name: cEntry.name,
          ctdtPrereq: cEntry.prereq,
          khgdPrereq: kEntry.prereq,
          summary: pCmp.summary,
        });
      }
    }

    for (const [code, kEntry] of khgdParsed.lookup.entries()) {
      if (!ctdtParsed.lookup.has(code)) {
        missingInCtdt.push({
          code,
          name: kEntry.name,
          credits: kEntry.credits,
          khgdSemester: Array.from(khgdParsed.codeToGroups.get(code) || []).join(', '),
        });
      }
    }
  }

  // Quét lỗi chính tả có sẵn ngay từ File gốc MyBK
  const baseLookup = ctdtParsed?.lookup || khgdParsed?.lookup;
  if (baseLookup) {
    for (const [code, entry] of baseLookup.entries()) {
      const sp = checkCourseSpelling({
        code,
        rawVi: entry.rawVi,
        rawEn: entry.rawEn,
        sourceLabel: 'File gốc MyBK',
      });
      // Lọc các lỗi chính tả thực sự (typo) & khoảng trắng kép trên MyBK để báo Cô gọn gàng
      sp.filter((x) => x.severity === 'typo' || x.severity === 'format').forEach((item) =>
        mybkTypos.push({ ...item, courseName: entry.name })
      );
    }
  }

  // Danh sách các mục báo cáo Cô có cấu trúc để hiển thị checkbox chọn lọc
  const reportItems = [];

  khgdMisplacedSubgroups.forEach((sg, idx) => {
    reportItems.push({
      id: `mybk_sg_${idx}`,
      category: 'Cấu trúc KHGD trên MyBK',
      text: `[Sheet KHGD - ${sg.semester}] Mục "${sg.subgroupName}" là học phần Bắt buộc nhưng trên file xuất MyBK đang nằm dưới mục "Các học phần tự chọn" (làm tổng tín chỉ bắt buộc khai báo ở dòng tiêu đề ${sg.semester} chưa cộng nhóm này).`,
    });
  });

  if (bbInTcIssues.length > 0 && khgdMisplacedSubgroups.length === 0) {
    reportItems.push({
      id: 'mybk_bb_tc',
      category: 'Lệch Bắt buộc/Tự chọn trên MyBK',
      text: `Các môn [${bbInTcIssues
        .map((x) => `${x.code} - ${x.name}`)
        .join('; ')}] thuộc nhóm Bắt buộc [BB] ở sheet Khối kiến thức nhưng đang nằm ở mục Tự chọn [TC] bên sheet Kế hoạch giảng dạy của MyBK.`,
    });
  }

  parentTagWarnings.forEach((pw, idx) => {
    reportItems.push({
      id: `mybk_pt_${idx}`,
      category: 'Nhãn nhóm trên MyBK',
      text: `[Sheet Khối kiến thức] Tiêu đề nhóm lớn "${pw.rawTitle}" đang gắn ký hiệu [TC], dù bên trong bao gồm cả nhóm con Bắt buộc [BB].`,
    });
  });

  prereqCrossIssues.forEach((pc, idx) => {
    reportItems.push({
      id: `mybk_pr_${idx}`,
      category: 'Lệch Tiên quyết giữa 2 sheet MyBK',
      text: `Môn ${pc.code} (${pc.name}): Học phần điều kiện giữa 2 sheet của MyBK không khớp (Khối kiến thức: "${
        pc.ctdtPrereq || 'Trống'
      }" ≠ KHGD: "${pc.khgdPrereq || 'Trống'}").`,
    });
  });

  missingInKhgd.forEach((mk, idx) => {
    reportItems.push({
      id: `mybk_mk_${idx}`,
      category: 'Khác biệt danh mục giữa 2 sheet MyBK',
      text: `Học phần/Điều kiện ${mk.code} (${mk.name} - Nhóm ${mk.ctdtGroup}) có trong sheet Khối kiến thức nhưng không xuất hiện bên sheet Kế hoạch giảng dạy của MyBK.`,
    });
  });

  missingInCtdt.forEach((mc, idx) => {
    reportItems.push({
      id: `mybk_mc_${idx}`,
      category: 'Khác biệt danh mục giữa 2 sheet MyBK',
      text: `Môn ${mc.code} (${mc.name} - ${mc.khgdSemester}) có ở sheet Kế hoạch giảng dạy nhưng không có bên sheet Khối kiến thức của MyBK.`,
    });
  });

  mybkTypos.forEach((tp, idx) => {
    reportItems.push({
      id: `mybk_tp_${idx}`,
      category: 'Chính tả tên môn gốc trên MyBK',
      text: `Môn ${tp.code} (${tp.courseName}): Tên gốc trên hệ thống MyBK có lỗi "${tp.detail}" (Hiện tại: "${tp.text}") — xin ý kiến Cô có chỉnh lại trên bản đẹp / cập nhật hệ thống không ạ.`,
    });
  });

  return {
    ctdtParsed,
    khgdParsed,
    bbInTcIssues,
    missingInKhgd,
    missingInCtdt,
    prereqCrossIssues,
    parentTagWarnings,
    khgdMisplacedSubgroups,
    mybkTypos,
    reportItems,
    totalIssues: reportItems.length,
  };
}

/**
 * Chạy đối chiếu hoàn chỉnh cho 1 cặp (Sheet Gốc MyBK <-> File Bản đẹp tương ứng: CTĐT hoặc KHGD)
 */
export function runSingleModeComparison({
  origRows,
  prettyRows,
  mode,
  fileKey, // 'ctdt' | 'khgd'
  fileLabel, // 'Bản đẹp CTĐT' | 'Bản đẹp KHGD'
  mybkAudit = null,
}) {
  const origParsed = parseSheetData(origRows, mode);
  const prettyParsed = parseSheetData(prettyRows, mode);
  const rowResults = compareRows(origRows, origParsed, prettyParsed, mode, fileLabel);

  const ctdtRef = mybkAudit?.ctdtParsed || (mode === 'ctdt' ? origParsed : null);
  const khgdRef = mybkAudit?.khgdParsed || (mode === 'khgd' ? origParsed : null);

  // B1: Phân nhóm & BB/TC trên file bản đẹp này
  const groupMismatches = [];
  const categoryMismatches = [];

  for (const [code, origSet] of origParsed.codeToGroups.entries()) {
    const prettySet = prettyParsed.codeToGroups.get(code);
    const info = origParsed.lookup.get(code) || prettyParsed.lookup.get(code) || {};
    const origCat = origParsed.codeToCategory.get(code) || '';
    const prettyCat = prettyParsed.codeToCategory.get(code) || '';

    if (prettySet) {
      const origArr = Array.from(origSet);
      const prettyArr = Array.from(prettySet);
      const isSameSet =
        origArr.length === prettyArr.length && origArr.every((g) => prettySet.has(g));
      if (!isSameSet) {
        groupMismatches.push({
          code,
          name: info.name || '',
          origGrp: origArr.join(', '),
          prettyGrp: prettyArr.join(', '),
          origCat,
          prettyCat,
          fileKey,
          fileLabel,
        });
      }
    }

    if (origCat && prettyCat && origCat !== prettyCat) {
      categoryMismatches.push({
        code,
        name: info.name || '',
        origGrp: Array.from(origSet).join(', '),
        prettyGrp: Array.from(prettySet || []).join(', '),
        origCat,
        prettyCat,
        fileKey,
        fileLabel,
      });
    }
  }

  // Ngoài ra nếu đây là Bản đẹp KHGD, kiểm tra xem Bản đẹp KHGD có bị đặt môn [BB] của CTĐT vào mục [TC] không!
  const khgdPrettyBbInTc = [];
  if (mode === 'khgd' && ctdtRef) {
    for (const [code, pEntry] of prettyParsed.lookup.entries()) {
      const ctdtCat = ctdtRef.codeToCategory.get(code);
      const prettyKhgdCat = prettyParsed.codeToCategory.get(code);
      if (ctdtCat && prettyKhgdCat && ctdtCat !== prettyKhgdCat) {
        khgdPrettyBbInTc.push({
          code,
          name: pEntry.name,
          credits: pEntry.credits,
          ctdtGroup: ctdtRef.codeToGroup.get(code) || '',
          ctdtCategory: ctdtCat,
          khgdSemester: Array.from(prettyParsed.codeToGroups.get(code) || []).join(', '),
          khgdCategory: prettyKhgdCat,
          fileKey,
          fileLabel,
        });
      }
    }
  }

  const allGroupIds = [
    ...new Set([...origParsed.groups.keys(), ...prettyParsed.groups.keys()]),
  ].sort((a, b) => a.localeCompare(b, 'vi'));

  const groupOverview = allGroupIds.map((gid) => {
    const oGrp = origParsed.groups.get(gid);
    const pGrp = prettyParsed.groups.get(gid);
    return {
      gid,
      label: pGrp?.label || oGrp?.label || gid,
      origCategory: oGrp?.category || '—',
      prettyCategory: pGrp?.category || '—',
      origCount: oGrp?.courseCount ?? 0,
      prettyCount: pGrp?.courseCount ?? 0,
      fileLabel,
    };
  });

  // B2: Sót môn, Đủ tín chỉ, Nội dung & Chính tả
  const missingInPretty = [];
  const extraInPretty = [];
  const orderDifferences = [];

  for (const [code, oEntry] of origParsed.lookup.entries()) {
    if (!prettyParsed.lookup.has(code)) {
      missingInPretty.push({ ...oEntry, fileKey, fileLabel });
    }
  }

  for (const [code, pEntry] of prettyParsed.lookup.entries()) {
    if (!origParsed.lookup.has(code)) {
      extraInPretty.push({ ...pEntry, fileKey, fileLabel });
    }
  }

  for (const gid of allGroupIds) {
    const oCourses = origParsed.groups.get(gid)?.courses || [];
    const pCourses = prettyParsed.groups.get(gid)?.courses || [];
    const commonOrig = oCourses.filter((c) => pCourses.includes(c));
    const commonPretty = pCourses.filter((c) => oCourses.includes(c));
    if (
      commonOrig.length > 1 &&
      commonOrig.length === commonPretty.length &&
      commonOrig.some((c, i) => c !== commonPretty[i])
    ) {
      orderDifferences.push({
        gid,
        origOrder: commonOrig,
        prettyOrder: commonPretty,
        fileLabel,
      });
    }
  }

  const groupCredits = allGroupIds.map((gid) => {
    const oGrp = origParsed.groups.get(gid);
    const pGrp = prettyParsed.groups.get(gid);
    const origSum = oGrp ? oGrp.creditsSum : null;
    const prettyDec = pGrp?.creditsDeclared ?? null;
    const prettySum = pGrp ? pGrp.creditsSum : null;
    const prettyRef = prettyDec ?? prettySum;
    const label = pGrp?.label || oGrp?.label || gid;
    const category =
      pGrp?.category && pGrp.category !== 'PARENT'
        ? pGrp.category
        : oGrp?.category || '';

    let effectiveOrigSum = origSum;
    let effectivePrettySum = prettySum;
    const isParent = gid === 'A' || gid === 'C';
    if (isParent) {
      const childIds = allGroupIds.filter((id) => id.startsWith(`${gid}.`));
      if (childIds.length > 0) {
        effectiveOrigSum = childIds.reduce(
          (acc, cid) => acc + (origParsed.groups.get(cid)?.creditsSum || 0),
          0
        );
        effectivePrettySum = childIds.reduce(
          (acc, cid) => acc + (prettyParsed.groups.get(cid)?.creditsSum || 0),
          0
        );
      }
    }

    let match = false;
    let note = '';
    if (isParent) {
      const childIds = allGroupIds.filter((id) => id.startsWith(`${gid}.`));
      const childDeclaredSum = childIds.reduce(
        (acc, cid) =>
          acc +
          (prettyParsed.groups.get(cid)?.creditsDeclared ??
            prettyParsed.groups.get(cid)?.creditsSum ??
            0),
        0
      );
      match = prettyDec != null ? prettyDec === childDeclaredSum : true;
      note =
        prettyDec != null
          ? `Tổng khai báo nhóm con = ${childDeclaredSum} TC`
          : 'Nhóm cha';
    } else if (category === 'TC') {
      const samePoolSum = effectiveOrigSum === effectivePrettySum;
      const validQuota =
        prettyDec == null || (prettyDec > 0 && prettyDec <= effectivePrettySum);
      match = samePoolSum && validQuota;
      note =
        prettyDec != null
          ? `Tự chọn: chọn ${prettyDec} / ${effectivePrettySum} TC`
          : 'Nhóm tự chọn';
    } else if (mode === 'khgd') {
      // Bỏ kiểm tra lệch tín chỉ khai báo dòng tiêu đề Học kỳ trên KHGD (chỉ đối chiếu tổng môn mở trong kỳ nếu cần tham khảo)
      match = true;
      note = 'Không kiểm TC khai báo đầu HK';
    } else {
      match =
        effectiveOrigSum != null &&
        prettyRef != null &&
        effectiveOrigSum === prettyRef;
      note = category === 'BB' ? 'Bắt buộc (tổng môn = TC yêu cầu)' : '';
    }

    return {
      gid,
      isParent,
      category,
      label,
      origSum: effectiveOrigSum,
      prettyDec,
      prettySum: effectivePrettySum,
      prettyRef,
      match,
      note,
      origCount: oGrp?.courseCount,
      prettyCount: pGrp?.courseCount,
      excelRow: pGrp?.excelRow,
      fileLabel,
    };
  });

  const declaredProgramCredits =
    prettyParsed.meta.declaredTotalCredits ||
    origParsed.meta.declaredTotalCredits ||
    60;

  let calculatedProgramCredits = null;
  if (mode === 'ctdt') {
    const leafGroupIds = allGroupIds.filter((gid) => {
      const hasChildren = allGroupIds.some((other) => other.startsWith(`${gid}.`));
      return !hasChildren;
    });
    calculatedProgramCredits = leafGroupIds.reduce((sum, gid) => {
      const pGrp = prettyParsed.groups.get(gid);
      const oGrp = origParsed.groups.get(gid);
      const val = pGrp?.creditsDeclared ?? pGrp?.creditsSum ?? oGrp?.creditsSum ?? 0;
      return sum + val;
    }, 0);
  }

  // Quét chính tả và phân biệt rõ: Lỗi phát sinh ở riêng Bản đẹp hay lỗi có sẵn từ Gốc MyBK!
  const spellingIssues = [];
  for (const [code, pEntry] of prettyParsed.lookup.entries()) {
    const pIssues = checkCourseSpelling({
      code,
      rawVi: pEntry.rawVi,
      rawEn: pEntry.rawEn,
      sourceLabel: fileLabel,
    });
    const oEntry = origParsed.lookup.get(code);
    const oIssues = oEntry
      ? checkCourseSpelling({
          code,
          rawVi: oEntry.rawVi,
          rawEn: oEntry.rawEn,
          sourceLabel: 'Gốc MyBK',
        })
      : [];

    for (const pi of pIssues) {
      const alsoInMybk = oIssues.some(
        (oi) => oi.lang === pi.lang && oi.severity === pi.severity && oi.detail === pi.detail
      );
      spellingIssues.push({
        ...pi,
        fileKey,
        fileLabel,
        alsoInMybk,
        originBadge: alsoInMybk
          ? `Có từ Gốc MyBK & giữ ở ${fileLabel}`
          : `Sai riêng ở ${fileLabel}`,
      });
    }
  }

  // B3: Tiên quyết, Môn bắt buộc, Ghi chú & Điều kiện tốt nghiệp
  const prereqMismatches = [];
  const allPrereqCourses = [];
  const seenCodes = new Set();

  const khgdSemesterNumMap = new Map();
  const semSource = mode === 'khgd' ? prettyParsed : khgdRef;
  if (semSource) {
    for (const [c, gSet] of semSource.codeToGroups.entries()) {
      const nums = Array.from(gSet)
        .map((hk) => {
          const m = hk.match(/\d+/);
          return m ? Number(m[0]) : null;
        })
        .filter((n) => n != null);
      if (nums.length > 0) {
        khgdSemesterNumMap.set(c, Math.min(...nums));
      }
    }
  }

  const allProgramCourseCodes = new Set([
    ...origParsed.lookup.keys(),
    ...prettyParsed.lookup.keys(),
    ...(ctdtRef ? ctdtRef.lookup.keys() : []),
  ]);

  const usedPrereqTypesPretty = new Set();
  const usedPrereqTypesOrig = new Set();

  for (const r of rowResults) {
    if (r.type !== 'vi' || seenCodes.has(r.code)) continue;
    seenCodes.add(r.code);

    (r.prereqItems || []).forEach((it) => {
      if (it.type) usedPrereqTypesOrig.add(it.type);
    });
    (r.pEntry?.prereqItems || []).forEach((it) => {
      if (it.type) usedPrereqTypesPretty.add(it.type);
    });

    const hasOrigPrereq = r.prereqItems && r.prereqItems.length > 0;
    const hasPrettyPrereq =
      r.pEntry && r.pEntry.prereqItems && r.pEntry.prereqItems.length > 0;

    if (hasOrigPrereq || hasPrettyPrereq || r.prereqMismatch) {
      const combinedItems =
        r.prereqItems && r.prereqItems.length > 0
          ? r.prereqItems
          : r.pEntry?.prereqItems || [];

      const prereqAnalysis = combinedItems.map((it) => {
        const inProgram = allProgramCourseCodes.has(it.code);
        let semesterCheck = null;
        if (inProgram && khgdSemesterNumMap.has(r.code) && khgdSemesterNumMap.has(it.code)) {
          const courseSem = khgdSemesterNumMap.get(r.code);
          const reqSem = khgdSemesterNumMap.get(it.code);
          const isCoReq = it.type === 'SH' || it.type === 'SHT';
          const ok = isCoReq ? reqSem <= courseSem : reqSem < courseSem;
          semesterCheck = {
            ok,
            courseSem: `HK${courseSem}`,
            reqSem: `HK${reqSem}`,
            msg: ok
              ? `Hợp lệ (${it.code}: HK${reqSem} → ${r.code}: HK${courseSem})`
              : `Lệch thứ tự HK (${it.code} ở HK${reqSem} nhưng ${r.code} ở HK${courseSem})`,
          };
        }
        return {
          ...it,
          inProgram,
          scopeLabel: inProgram
            ? 'Môn trong khung CTĐT'
            : 'Môn ĐK bậc ĐH / ngoài khung',
          semesterCheck,
        };
      });

      const enriched = { ...r, prereqAnalysis, fileKey, fileLabel };
      if (r.prereqMismatch) prereqMismatches.push(enriched);
      allPrereqCourses.push(enriched);
    }
  }

  const compulsoryChecklist = [];
  const specialRequirements = [];

  for (const r of rowResults) {
    if (r.type !== 'vi') continue;
    const ctdtCat = ctdtRef?.codeToCategory.get(r.code) || r.origCat;
    const grp = ctdtRef?.codeToGroup.get(r.code) || r.origGrp;

    const isGradOrLangReq =
      grp === 'E' ||
      grp === 'D' ||
      grp === 'A.3' ||
      r.code.includes('ENG') ||
      /ngoại ngữ|anh văn|tốt nghiệp|thực tập|đề án|luận văn/i.test(r.name);

    if (isGradOrLangReq && !specialRequirements.some((x) => x.code === r.code)) {
      const inKhgd = khgdRef ? khgdRef.lookup.has(r.code) : null;
      specialRequirements.push({
        code: r.code,
        name: r.name,
        eng: r.pEntry?.eng || origParsed.lookup.get(r.code)?.eng || '',
        credits: r.credits,
        group: grp,
        origPrereq: r.prereq,
        prettyPrereq: r.pEntry?.prereq || '',
        inPretty: Boolean(r.pEntry),
        inKhgd,
        status: r.status,
        prereqMismatch: r.prereqMismatch,
        prereqSummary: r.prereqDiff?.summary || '',
        fileLabel,
      });
    }

    if (ctdtCat === 'BB' && !compulsoryChecklist.some((x) => x.code === r.code)) {
      compulsoryChecklist.push({
        code: r.code,
        name: r.name,
        credits: r.credits,
        group: grp,
        inPretty: Boolean(r.pEntry),
        status: r.status,
        prereq: r.prereq,
        prettyPrereq: r.pEntry?.prereq || '',
        prereqMismatch: r.prereqMismatch,
        note: r.note || '',
        fileLabel,
      });
    }
  }

  const legendWarnings = [];
  for (const t of usedPrereqTypesOrig) {
    if (
      prettyParsed.meta.legendDefinedTypes.size > 0 &&
      !prettyParsed.meta.legendDefinedTypes.has(t)
    ) {
      legendWarnings.push({
        fileLabel,
        text: `File gốc MyBK dùng ký hiệu "${t}" (${
          t === 'SHT' ? 'Song hành/Trước' : t === 'HT' ? 'Học trước' : t
        }), nhưng dòng Ghi chú đầu bảng của ${fileLabel} chỉ giải thích: ${Array.from(
          prettyParsed.meta.legendDefinedTypes
        ).join(', ')}.`,
      });
    }
  }
  for (const t of usedPrereqTypesPretty) {
    if (
      prettyParsed.meta.legendDefinedTypes.size > 0 &&
      !prettyParsed.meta.legendDefinedTypes.has(t)
    ) {
      legendWarnings.push({
        fileLabel,
        text: `Trong bảng của ${fileLabel} có dùng ký hiệu "${t}" nhưng chưa giải thích trên dòng Ghi chú (notes) đầu bảng.`,
      });
    }
  }

  // B4: Ánh xạ chi tiết đến TỪNG Ô (Cell-level) của từng dòng trong File Bản đẹp để tô chữ đỏ / highlight ô đỏ!
  const prettyRowMap = new Map();
  const spellingByCode = new Map();
  for (const sp of spellingIssues) {
    if (!spellingByCode.has(sp.code)) spellingByCode.set(sp.code, []);
    spellingByCode.get(sp.code).push(sp);
  }

  const rowResultByCode = new Map();
  const enResultByCode = new Map();
  for (const r of rowResults) {
    if (r.type === 'vi' && !rowResultByCode.has(r.code)) {
      rowResultByCode.set(r.code, r);
    } else if (r.type === 'en' && r.parentCode && !enResultByCode.has(r.parentCode)) {
      enResultByCode.set(r.parentCode, r);
    }
  }

  for (const occ of prettyParsed.occurrences) {
    const { code, excelRow, engExcelRow } = occ;
    const rRes = rowResultByCode.get(code);
    const enRes = enResultByCode.get(code);
    const spList = spellingByCode.get(code) || [];
    const khgdBbWarn = khgdPrettyBbInTc.find((x) => x.code === code);

    const autoNotes = [];
    let severity = 'ok'; // 'ok' | 'warn' | 'error'

    // Đánh dấu chính xác Ô NÀO bị lỗi trên dòng này để exporter tô đỏ đúng ô đó!
    const cellErrors = {
      codeCol: false, // Lỗi mã môn thừa / sai nhóm / sai học kỳ / lệch BB-TC
      viNameCol: false, // Lỗi tên Tiếng Việt / chính tả TV
      enNameCol: false, // Lỗi tên Tiếng Anh / chính tả TA
      credCol: false, // Lỗi số tín chỉ
      prereqCol: false, // Lỗi thiếu/sai học phần tiên quyết
      missingPrereqText: '', // Nội dung tiên quyết bị thiếu để điền chữ đỏ gợi ý vào ô trống
    };

    if (!rRes) {
      severity = 'error';
      cellErrors.codeCol = true;
      autoNotes.push(
        `[Thừa/Sai mã môn ở ${fileLabel}] Môn này có trong ${fileLabel} nhưng KHÔNG có trong File gốc MyBK`
      );
    } else {
      // Kiểm tra học kỳ / nhóm cụ thể của dòng này
      const origGroupsSet = origParsed.codeToGroups.get(code);
      if (occ.group && origGroupsSet && !origGroupsSet.has(occ.group)) {
        severity = 'error';
        cellErrors.codeCol = true;
        autoNotes.push(
          `[B1-Sai ${mode === 'khgd' ? 'Học kỳ' : 'Nhóm'}] Ở ${fileLabel} đặt tại ${
            occ.group
          }, nhưng Gốc MyBK là ${Array.from(origGroupsSet).join(', ')}`
        );
      }
      if (rRes.categoryMismatch) {
        severity = 'error';
        cellErrors.codeCol = true;
        autoNotes.push(
          `[B1-Lệch BB/TC] Gốc MyBK là [${rRes.origCat}] nhưng ${fileLabel} đang xếp ở [${rRes.prettyCat}]`
        );
      }
      if (khgdBbWarn && !rRes.categoryMismatch) {
        if (severity === 'ok') severity = 'warn';
        cellErrors.codeCol = true;
        autoNotes.push(
          `[B1-Chéo CTĐT] Môn [${khgdBbWarn.ctdtCategory}] ở CTĐT nhưng đang nằm mục [${khgdBbWarn.khgdCategory}] trong ${fileLabel}`
        );
      }

      if (clean(occ.name) !== clean(rRes.name)) {
        severity = 'error';
        cellErrors.viNameCol = true;
        autoNotes.push(`[B2-Sai tên TV] Gốc MyBK: "${rRes.name}"`);
      }

      const origCred =
        rRes.credits !== '' && !isNaN(Number(rRes.credits))
          ? Number(rRes.credits)
          : null;
      if (origCred !== occ.credits) {
        severity = 'error';
        cellErrors.credCol = true;
        autoNotes.push(
          `[B2-Sai Tín chỉ] Gốc MyBK: ${rRes.credits} TC ≠ ${fileLabel}: ${
            occ.credits ?? 'Trống'
          } TC`
        );
      }

      if (enRes && clean(occ.eng) !== clean(enRes.name)) {
        severity = 'error';
        cellErrors.enNameCol = true;
        autoNotes.push(`[B2-Sai tên TA] Gốc MyBK: "${enRes.name}"`);
      }

      const occPrereqCmp = comparePrereq(rRes.prereq, occ.prereq);
      if (!occPrereqCmp.match) {
        severity = 'error';
        cellErrors.prereqCol = true;
        if (!clean(occ.prereq) && clean(rRes.prereq)) {
          cellErrors.missingPrereqText = rRes.prereq;
        }
        autoNotes.push(
          `[B3-Tiên quyết] ${occPrereqCmp.summary} (Gốc MyBK: ${
            rRes.prereq || 'Trống'
          })`
        );
      }
    }

    for (const sp of spList) {
      if (sp.severity === 'typo') {
        severity = 'error';
        if (sp.lang === 'VI') cellErrors.viNameCol = true;
        if (sp.lang === 'EN') cellErrors.enNameCol = true;
        autoNotes.push(
          `[B2-Chính tả ${sp.lang}${sp.alsoInMybk ? ' (Từ gốc MyBK)' : ''}] ${sp.detail}`
        );
      } else if (sp.severity === 'format') {
        if (severity === 'ok') severity = 'warn';
        autoNotes.push(`[B2-Khoảng trắng ${sp.lang}] ${sp.detail}`);
      }
    }

    prettyRowMap.set(excelRow, {
      excelRow,
      engExcelRow,
      code,
      name: occ.name,
      eng: occ.eng,
      credits: occ.credits,
      group: occ.group,
      category: occ.category,
      prereq: occ.prereq,
      origPrereq: rRes?.prereq || '',
      severity,
      cellErrors,
      statusText:
        severity === 'ok'
          ? '✓ Khớp chuẩn'
          : severity === 'error'
          ? `✗ Lỗi ở ${fileLabel}`
          : '⚠ Lưu ý rà soát',
      autoNoteText:
        autoNotes.length > 0 ? autoNotes.join(' ; ') : 'Khớp 100% với File gốc MyBK',
      autoNotes,
      fileKey,
      fileLabel,
    });
  }

  const validRows = rowResults.filter((r) => r.type === 'vi' || r.type === 'en');
  const typoCount = spellingIssues.filter((s) => s.severity === 'typo').length;
  const prettyOnlyTypoCount = spellingIssues.filter(
    (s) => s.severity === 'typo' && !s.alsoInMybk
  ).length;
  const creditErrCount = groupCredits.filter(
    (g) => g.origSum != null && g.prettyRef != null && !g.match
  ).length;
  const contentErrCount = validRows.filter(
    (r) => r.status === 'wrong' || r.status === 'wrong-en'
  ).length;

  const totalFileErrors =
    groupMismatches.length +
    categoryMismatches.length +
    missingInPretty.length +
    extraInPretty.length +
    contentErrCount +
    creditErrCount +
    prereqMismatches.length +
    typoCount;

  const stats = {
    total: validRows.length,
    match: validRows.filter((r) => r.status === 'match').length,
    miss: missingInPretty.length,
    extra: extraInPretty.length,
    wrong: contentErrCount,
    wrongPrereq: prereqMismatches.length,
    groupMismatch: groupMismatches.length,
    categoryMismatch: categoryMismatches.length + khgdPrettyBbInTc.length,
    creditMismatch: creditErrCount,
    typoCount,
    prettyOnlyTypoCount,
    totalFileErrors,
    b1Issues:
      groupMismatches.length + categoryMismatches.length + khgdPrettyBbInTc.length,
    b2Issues:
      missingInPretty.length +
      extraInPretty.length +
      contentErrCount +
      creditErrCount +
      typoCount,
    b3Issues:
      prereqMismatches.length +
      legendWarnings.length +
      specialRequirements.filter((s) => !s.inPretty || s.prereqMismatch).length,
  };

  return {
    mode,
    fileKey,
    fileLabel,
    origParsed,
    prettyParsed,
    prettyRows,
    rowResults,
    // B1
    groupMismatches,
    categoryMismatches,
    khgdPrettyBbInTc,
    groupOverview,
    // B2
    missingInPretty,
    extraInPretty,
    orderDifferences,
    groupCredits,
    declaredProgramCredits,
    calculatedProgramCredits,
    spellingIssues,
    // B3
    prereqMismatches,
    allPrereqCourses,
    compulsoryChecklist,
    specialRequirements,
    legendWarnings,
    // B4
    prettyRowMap,
    prettyCoursesList: Array.from(prettyRowMap.values()),
    stats,
  };
}

/**
 * Đối chiếu chéo giữa 2 File Bản Đẹp (CTĐT Bản đẹp <-> KHGD Bản đẹp) khi người dùng tải lên đủ cả 3 file
 */
export function runPrettyCrossCheck(ctdtRes, khgdRes) {
  if (!ctdtRes || !khgdRes) return [];
  const diffs = [];

  for (const [code, cEntry] of ctdtRes.prettyParsed.lookup.entries()) {
    const kEntry = khgdRes.prettyParsed.lookup.get(code);
    if (!kEntry) {
      if (cEntry.group !== 'E') {
        diffs.push({
          code,
          name: cEntry.name,
          type: 'Thiếu ở Bản đẹp KHGD',
          detail: `Môn có trong Bản đẹp CTĐT (Nhóm ${cEntry.group}) nhưng chưa có trong Bản đẹp KHGD.`,
          targetFile: 'Bản đẹp KHGD',
        });
      }
      continue;
    }

    if (clean(cEntry.name) !== clean(kEntry.name)) {
      diffs.push({
        code,
        name: cEntry.name,
        type: 'Lệch tên TV giữa 2 Bản đẹp',
        detail: `CTĐT đẹp: "${cEntry.name}" ≠ KHGD đẹp: "${kEntry.name}"`,
        targetFile: 'Cả 2 Bản đẹp',
      });
    }
    if (clean(cEntry.eng) !== clean(kEntry.eng)) {
      diffs.push({
        code,
        name: cEntry.name,
        type: 'Lệch tên TA giữa 2 Bản đẹp',
        detail: `CTĐT đẹp: "${cEntry.eng}" ≠ KHGD đẹp: "${kEntry.eng}"`,
        targetFile: 'Cả 2 Bản đẹp',
      });
    }
    if (cEntry.credits !== kEntry.credits) {
      diffs.push({
        code,
        name: cEntry.name,
        type: 'Lệch tín chỉ giữa 2 Bản đẹp',
        detail: `CTĐT đẹp: ${cEntry.credits} TC ≠ KHGD đẹp: ${kEntry.credits} TC`,
        targetFile: 'Cả 2 Bản đẹp',
      });
    }
    const pCmp = comparePrereq(cEntry.prereq, kEntry.prereq);
    if (!pCmp.match) {
      diffs.push({
        code,
        name: cEntry.name,
        type: 'Lệch Tiên quyết giữa 2 Bản đẹp',
        detail: `CTĐT đẹp: "${cEntry.prereq || 'Trống'}" ≠ KHGD đẹp: "${
          kEntry.prereq || 'Trống'
        }" (${pCmp.summary})`,
        targetFile: 'Cả 2 Bản đẹp',
      });
    }
  }

  for (const [code, kEntry] of khgdRes.prettyParsed.lookup.entries()) {
    if (!ctdtRes.prettyParsed.lookup.has(code)) {
      diffs.push({
        code,
        name: kEntry.name,
        type: 'Thiếu ở Bản đẹp CTĐT',
        detail: `Môn có trong Bản đẹp KHGD (${kEntry.group}) nhưng không có trong Bản đẹp CTĐT.`,
        targetFile: 'Bản đẹp CTĐT',
      });
    }
  }

  return diffs;
}
