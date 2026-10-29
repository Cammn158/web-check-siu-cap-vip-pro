/**
 * Bộ kiểm tra chính tả & quy chuẩn định dạng tên môn học (Tiếng Anh & Tiếng Việt)
 */

// Từ điển các lỗi chính tả tiếng Anh học thuật/kỹ thuật thường gặp trong CTĐT
const EN_TYPO_DICT = {
  fundametals: 'Fundamentals',
  fundametal: 'Fundamental',
  experimential: 'Experimental',
  expermental: 'Experimental',
  mechannics: 'Mechanics',
  mecahnics: 'Mechanics',
  avanced: 'Advanced',
  advaced: 'Advanced',
  advaned: 'Advanced',
  enginering: 'Engineering',
  enginnering: 'Engineering',
  engineeering: 'Engineering',
  manufaturing: 'Manufacturing',
  manufactuing: 'Manufacturing',
  technolgy: 'Technology',
  techonology: 'Technology',
  analysys: 'Analysis',
  analisis: 'Analysis',
  simualtion: 'Simulation',
  simulaton: 'Simulation',
  optmization: 'Optimization',
  optimizaton: 'Optimization',
  inteligence: 'Intelligence',
  artifical: 'Artificial',
  mangement: 'Management',
  managment: 'Management',
  enviroment: 'Environment',
  enviromental: 'Environmental',
  maintainance: 'Maintenance',
  maintenace: 'Maintenance',
  electonic: 'Electronic',
  electonics: 'Electronics',
  eletrical: 'Electrical',
  machinary: 'Machinery',
  equipments: 'Equipment (không đếm được)',
  reseach: 'Research',
  reasearch: 'Research',
  developement: 'Development',
  aplication: 'Application',
  aplications: 'Applications',
  controll: 'Control',
  proceses: 'Processes',
  proccess: 'Process',
  proccessing: 'Processing',
  materals: 'Materials',
  hydralic: 'Hydraulic',
  hydraulics: null,
  numercial: 'Numerical',
  vibation: 'Vibration',
  vibations: 'Vibrations',
  thermodinamics: 'Thermodynamics',
  automatical: 'Automatic',
  automative: 'Automotive',
  stategy: 'Strategy',
  statictics: 'Statistics',
  bussiness: 'Business',
  corperate: 'Corporate',
  responsability: 'Responsibility',
  methodolgy: 'Methodology',
  scienctific: 'Scientific',
  entrepreneurhip: 'Entrepreneurship',
  enterpreneurship: 'Entrepreneurship',
  intergration: 'Integration',
  transmision: 'Transmission',
  distibution: 'Distribution',
  protetion: 'Protection',
  dyanmics: 'Dynamics',
  kinamatics: 'Kinematics',
  composit: 'Composite',
  metallugy: 'Metallurgy',
  metalurgy: 'Metallurgy',
  reliablity: 'Reliability',
  measurment: 'Measurement',
  construcion: 'Construction',
};

// Từ nối / giới từ tiếng Anh được phép viết thường trong Title Case
const EN_MINOR_WORDS = new Set([
  'a',
  'an',
  'the',
  'and',
  'or',
  'nor',
  'but',
  'for',
  'yet',
  'so',
  'at',
  'by',
  'in',
  'of',
  'on',
  'to',
  'up',
  'as',
  'vs',
  'via',
  'with',
  'from',
  'into',
  'onto',
  'upon',
  'over',
  'under',
]);

// Từ viết tắt kỹ thuật được phép viết IN HOA
const EN_ACRONYMS = new Set([
  'CAD',
  'CAM',
  'CAE',
  'CNC',
  'FACTS',
  'HVDC',
  'SCADA',
  'DSM',
  'AI',
  'IOT',
  'PLC',
  'FEM',
  'CFD',
  'NCKH',
  'IC',
  'VLSI',
  'FPGA',
  'MEMS',
  'PID',
  'DC',
  'AC',
  'II',
  'III',
  'IV',
]);

const VI_TYPO_PATTERNS = [
  { regex: /\s+[,.;:!?]/, msg: 'Dư khoảng trắng trước dấu câu' },
  { regex: /[,.;:](?=[^\s\d])/, msg: 'Thiếu khoảng trắng sau dấu câu' },
  { regex: /\bkỉ\s+thuật\b/i, msg: 'Chính tả tiếng Việt: "kỉ thuật" → "kỹ thuật"' },
  { regex: /\bchuẩn\s+đoán\b/i, msg: 'Chính tả tiếng Việt: "chuẩn đoán" → "chẩn đoán"' },
  { regex: /\bxử\s+lí\b/i, msg: 'Không đồng nhất chính tả: "xử lí" / "xử lý"' },
  { regex: /\bvật\s+lí\b/i, msg: 'Không đồng nhất chính tả: "vật lí" / "vật lý"' },
  { regex: /\bnguyên\s+lí\b/i, msg: 'Không đồng nhất chính tả: "nguyên lí" / "nguyên lý"' },
];

export function toSuggestedTitleCase(str) {
  if (!str) return '';
  const words = str.trim().replace(/\s+/g, ' ').split(' ');
  return words
    .map((word, index) => {
      // Giữ nguyên các cụm có dấu gạch chéo như CAD/CAM
      if (word.includes('/')) {
        return word
          .split('/')
          .map((part) => {
            const up = part.toUpperCase();
            if (EN_ACRONYMS.has(up)) return up;
            return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
          })
          .join('/');
      }
      const cleanWord = word.replace(/[^A-Za-z]/g, '');
      const upper = cleanWord.toUpperCase();
      if (EN_ACRONYMS.has(upper)) {
        return word.toUpperCase();
      }
      const lower = word.toLowerCase();
      if (index > 0 && index < words.length - 1 && EN_MINOR_WORDS.has(lower)) {
        return lower;
      }
      return lower.charAt(0).toUpperCase() + lower.slice(1);
    })
    .join(' ');
}

/**
 * Kiểm tra chính tả & quy chuẩn định dạng cho 1 môn học
 * @param {object} params
 * @param {string} params.code - Mã môn học
 * @param {string} params.rawVi - Tên tiếng Việt gốc (chưa gộp khoảng trắng)
 * @param {string} params.rawEn - Tên tiếng Anh gốc (chưa gộp khoảng trắng)
 * @param {string} params.sourceLabel - Nguồn ('File bản đẹp' hoặc 'File gốc')
 */
export function checkCourseSpelling({ code, rawVi, rawEn, sourceLabel = 'Bản đẹp' }) {
  const issues = [];

  // 1. Kiểm tra tên tiếng Việt
  if (rawVi != null && String(rawVi).length > 0) {
    const viStr = String(rawVi);
    const trimmedVi = viStr.trim();

    if (/\s{2,}/.test(trimmedVi)) {
      issues.push({
        code,
        lang: 'VI',
        severity: 'format',
        source: sourceLabel,
        text: trimmedVi,
        detail: 'Dư khoảng trắng kép giữa các từ trong tên tiếng Việt',
        suggestion: trimmedVi.replace(/\s+/g, ' '),
      });
    }

    for (const pat of VI_TYPO_PATTERNS) {
      if (pat.regex.test(trimmedVi)) {
        issues.push({
          code,
          lang: 'VI',
          severity: 'typo',
          source: sourceLabel,
          text: trimmedVi,
          detail: pat.msg,
          suggestion: '',
        });
      }
    }
  }

  // 2. Kiểm tra tên tiếng Anh
  if (rawEn != null && String(rawEn).trim().length > 0) {
    const enStr = String(rawEn);
    const trimmedEn = enStr.trim();

    // Kiểm tra khoảng trắng kép ở giữa câu
    if (/\s{2,}/.test(trimmedEn)) {
      issues.push({
        code,
        lang: 'EN',
        severity: 'format',
        source: sourceLabel,
        text: trimmedEn,
        detail: 'Dư khoảng trắng kép giữa các từ trong tên tiếng Anh',
        suggestion: trimmedEn.replace(/\s+/g, ' '),
      });
    }

    // Quét từng từ tiếng Anh qua từ điển lỗi chính tả
    const normalizedEn = trimmedEn.replace(/\s+/g, ' ');
    const tokens = normalizedEn.split(/[\s/,()-]+/).filter(Boolean);
    const typoMatches = [];

    for (const tok of tokens) {
      const cleanTok = tok.replace(/[^A-Za-z]/g, '').toLowerCase();
      if (!cleanTok) continue;
      if (EN_TYPO_DICT[cleanTok]) {
        typoMatches.push(`"${tok}" → "${EN_TYPO_DICT[cleanTok]}"`);
      } else if (/([a-z])\1\1/i.test(cleanTok)) {
        // 3 ký tự giống nhau liên tiếp (VD: engineeering)
        typoMatches.push(`Từ nghi ngờ lặp ký tự: "${tok}"`);
      }
    }

    if (typoMatches.length > 0) {
      issues.push({
        code,
        lang: 'EN',
        severity: 'typo',
        source: sourceLabel,
        text: normalizedEn,
        detail: `Sai chính tả tiếng Anh: ${typoMatches.join('; ')}`,
        suggestion: '',
      });
    }

    // Kiểm tra viết IN HOA TOÀN BỘ (ALL CAPS) không đồng nhất Title Case
    const lettersOnly = normalizedEn.replace(/[^A-Za-z]/g, '');
    if (
      lettersOnly.length >= 10 &&
      lettersOnly === lettersOnly.toUpperCase() &&
      !EN_ACRONYMS.has(lettersOnly)
    ) {
      issues.push({
        code,
        lang: 'EN',
        severity: 'style',
        source: sourceLabel,
        text: normalizedEn,
        detail: 'Viết IN HOA TOÀN BỘ (không đồng nhất chuẩn Title Case với các môn khác)',
        suggestion: toSuggestedTitleCase(normalizedEn),
      });
    } else if (lettersOnly.length >= 6) {
      // Kiểm tra từ chính bị viết thường chữ cái đầu trong Title Case
      const words = normalizedEn.split(' ');
      const lowerMajorWords = [];
      words.forEach((w, idx) => {
        const cleanW = w.replace(/[^A-Za-z]/g, '');
        if (!cleanW || cleanW.length <= 2) return;
        const lowerW = cleanW.toLowerCase();
        if (idx > 0 && EN_MINOR_WORDS.has(lowerW)) return;
        if (/^[a-z]/.test(w)) {
          lowerMajorWords.push(w);
        }
      });

      if (lowerMajorWords.length > 0) {
        issues.push({
          code,
          lang: 'EN',
          severity: 'style',
          source: sourceLabel,
          text: normalizedEn,
          detail: `Chữ hoa/thường chưa đồng nhất Title Case (từ viết thường: "${lowerMajorWords.join(', ')}")`,
          suggestion: toSuggestedTitleCase(normalizedEn),
        });
      }
    }
  }

  return issues;
}
