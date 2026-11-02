import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'framer-motion';
import { Toaster, toast } from 'sonner';
import {
  Building2,
  FileSpreadsheet,
  CalendarRange,
  Search,
  X,
  Play,
  Loader2,
  AlertTriangle,
  CheckCircle2,
  XCircle,
  Sparkles,
  Download,
  Copy,
  Check,
  Send,
  Eye,
  FileWarning,
  Layers,
  ListChecks,
  BookOpenCheck,
  SpellCheck,
  ShieldAlert,
  GraduationCap,
  ChevronLeft,
  ChevronRight,
  RotateCcw,
  Plus,
  Sun,
  Moon,
  Palette,
  SlidersHorizontal,
  FolderKanban,
  ArrowDownToLine,
  FileText,
  ChevronUp,
  ChevronDown,
  Zap,
  Info,
} from 'lucide-react';
import {
  clean,
  sheetToRows,
  runMybkInternalAudit,
  runSingleModeComparison,
  runPrettyCrossCheck,
} from './utils/parser.js';
import { exportAnnotatedPrettyCopy } from './utils/exporter.js';
import './styles/theme.css';

const DEFAULT_STEP_ORDER = ['b1', 'b2', 'b3', 'b4'];

const STEP_META = {
  b1: {
    id: 'b1',
    code: 'B1',
    title: 'Phân nhóm & Bắt buộc / Tự chọn',
    desc: 'Kiểm tra môn BB / TC thuộc đúng nhóm & học kỳ trên từng file',
    icon: Layers,
  },
  b2: {
    id: 'b2',
    code: 'B2',
    title: 'Sót môn, Tín chỉ, Nội dung & Chính tả',
    desc: 'Tách rõ sót/sai ở CTĐT đẹp, KHGD đẹp hay lỗi gốc từ MyBK',
    icon: ListChecks,
  },
  b3: {
    id: 'b3',
    code: 'B3',
    title: 'Tiên quyết (Quan trọng), Bắt buộc & ĐK Tốt nghiệp',
    desc: 'Rà tiên quyết, môn BB, ghi chú & Đ/kiện ngoại ngữ tốt nghiệp SĐH',
    priority: true,
    icon: ShieldAlert,
  },
  b4: {
    id: 'b4',
    code: 'B4',
    title: 'Xuất Bản Sao Tô Đỏ (CTĐT / KHGD)',
    desc: 'Tạo bản sao File đẹp với chữ đỏ / highlight ô đỏ để tự check & sửa',
    icon: FileWarning,
  },
};

const GROUP_COLORS_CTDT = {
  A: '#2563eb',
  'A.1': '#2563eb',
  'A.2': '#7c3aed',
  'A.3': '#0d9488',
  B: '#059669',
  C: '#d97706',
  'C.1': '#d97706',
  'C.2': '#ea580c',
  D: '#db2777',
  E: '#64748b',
};

const GROUP_COLORS_KHGD = {
  HK1: '#059669',
  HK2: '#2563eb',
  HK3: '#d97706',
  HK4: '#ea580c',
  HK5: '#db2777',
};

function getGroupColor(gid, mode) {
  return (mode === 'khgd' ? GROUP_COLORS_KHGD : GROUP_COLORS_CTDT)[gid] || '#64748b';
}

const STATUS_BADGES = {
  match: { cls: 'b-match', icon: '✓', short: 'Khớp 100%' },
  miss: { cls: 'b-miss', icon: '✗', short: 'Thiếu môn' },
  wrong: { cls: 'b-wrong', icon: '!', short: 'Sai thông tin' },
  'wrong-en': { cls: 'b-wrong-en', icon: '~', short: 'Sai tiếng Anh' },
  'wrong-prereq': { cls: 'b-wrong-prereq', icon: '⚡', short: 'Sai ĐK tiên quyết' },
  na: { cls: 'b-na', icon: '?', short: '#N/A' },
  skip: { cls: 'b-skip', icon: '–', short: 'Bỏ qua' },
};

const GROUP_COL_TITLE = {
  ctdt: 'Nhóm (A/B/C...)',
  khgd: 'Học kỳ (HK)',
};

function CategoryBadge({ cat }) {
  if (!cat || cat === '—' || cat === 'HK') {
    return <span style={{ color: 'var(--text2)', fontSize: 11 }}>—</span>;
  }
  if (cat === 'PARENT') {
    return <span className="badge b-skip">Nhóm cha</span>;
  }
  return (
    <span className={`badge ${cat === 'BB' ? 'b-bb' : 'b-tc'}`}>
      {cat === 'BB' ? 'BB • Bắt buộc' : 'TC • Tự chọn'}
    </span>
  );
}

function PrereqPills({ items, emptyText = '—' }) {
  if (!items || items.length === 0) {
    return <span style={{ color: 'var(--text2)', fontSize: 11 }}>{emptyText}</span>;
  }
  return (
    <div className="prereq-list">
      {items.map((it, idx) => (
        <span
          key={idx}
          className={`prereq-pill ${it.type ? `t-${it.type}` : ''}`}
          title={it.type || 'Điều kiện'}
        >
          <span>{it.code}</span>
          {it.type && <span style={{ opacity: 0.85 }}>({it.type})</span>}
        </span>
      ))}
    </div>
  );
}

export default function App() {
  // Giao diện Theme mới (Mặc định: 'light' Studio để lột xác hoàn toàn so với theme tối cũ, hỗ trợ đổi sang 'emerald' và 'dark')
  const [theme, setTheme] = useState(() => {
    return localStorage.getItem('web_check_theme') || 'light';
  });
  const [searchQuery, setSearchQuery] = useState('');
  const [collapseUploads, setCollapseUploads] = useState(false);

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('web_check_theme', theme);
  }, [theme]);

  // 3 FILE ĐẦU VÀO: 1 File gốc MyBK + 2 File Bản đẹp (CTĐT & KHGD)
  const [fileMybk, setFileMybk] = useState(null);
  const [fileCtdtPretty, setFileCtdtPretty] = useState(null);
  const [fileKhgdPretty, setFileKhgdPretty] = useState(null);

  const [dragMybk, setDragMybk] = useState(false);
  const [dragCtdt, setDragCtdt] = useState(false);
  const [dragKhgd, setDragKhgd] = useState(false);

  // Kết quả phân tích tách riêng từng file
  const [reviewData, setReviewData] = useState(null); // { ctdtResult, khgdResult, mybkAudit, prettyCrossDiffs }
  const [activeFileScope, setActiveFileScope] = useState('ctdt'); // 'ctdt' | 'khgd'
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  // Thứ tự 4 bước (có thể sắp xếp lại)
  const [stepOrder, setStepOrder] = useState(DEFAULT_STEP_ORDER);
  const [activeStep, setActiveStep] = useState('b1');

  // Bộ lọc bên trong các bước
  const [b2SubTab, setB2SubTab] = useState('rows');
  const [rowFilter, setRowFilter] = useState('all');
  const [spellFilter, setSpellFilter] = useState('all');
  const [prereqFilter, setPrereqFilter] = useState('all');
  const [b4Filter, setB4Filter] = useState('all');

  // Dữ liệu Check & Note TÁCH RIÊNG cho Bản đẹp CTĐT và Bản đẹp KHGD
  const [notesByFile, setNotesByFile] = useState({ ctdt: {}, khgd: {} });
  const [checkedByFile, setCheckedByFile] = useState({ ctdt: {}, khgd: {} });
  const [generalNoteByFile, setGeneralNoteByFile] = useState({ ctdt: '', khgd: '' });
  const [outNameByFile, setOutNameByFile] = useState({ ctdt: '', khgd: '' });

  // Tuỳ chọn kiểu tô đỏ khi xuất Bản sao Excel ở B4
  const [markStyle, setMarkStyle] = useState('both'); // 'both' | 'red_text' | 'highlight'
  const [fillMissingWithRedHint, setFillMissingWithRedHint] = useState(true);
  const [includeSummarySheet, setIncludeSummarySheet] = useState(true);
  const [exportingKey, setExportingKey] = useState(null);

  // Khung Note File MyBK để báo riêng với Cô
  const [mybkSelectedIds, setMybkSelectedIds] = useState({});
  const [mybkCustomItems, setMybkCustomItems] = useState([]);
  const [newMybkCustomInput, setNewMybkCustomInput] = useState('');
  const [profDraftEdited, setProfDraftEdited] = useState(null); // null = tự động sinh từ danh sách tick
  const [copiedProf, setCopiedProf] = useState(false);
  const profPanelRef = useRef(null);

  const moveStep = (stepId, dir, e) => {
    e.stopPropagation();
    const idx = stepOrder.indexOf(stepId);
    const targetIdx = idx + dir;
    if (idx === -1 || targetIdx < 0 || targetIdx >= stepOrder.length) return;
    const next = [...stepOrder];
    [next[idx], next[targetIdx]] = [next[targetIdx], next[idx]];
    setStepOrder(next);
  };

  const onPickFile = (slot) => (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    applyFileToSlot(slot, f);
  };

  const onDropFile = (slot, setDrag) => (e) => {
    e.preventDefault();
    setDrag(false);
    const f = e.dataTransfer.files?.[0];
    if (!f) return;
    applyFileToSlot(slot, f);
  };

  const applyFileToSlot = (slot, f) => {
    const base = f.name.replace(/\.(xlsx|xls|csv)$/i, '');
    if (slot === 'mybk') {
      setFileMybk(f);
      toast.success(`Đã nạp File 1 (Gốc MyBK): ${f.name}`);
    } else if (slot === 'ctdt') {
      setFileCtdtPretty(f);
      setOutNameByFile((prev) => ({
        ...prev,
        ctdt: `${base}_check_note_do.xlsx`,
      }));
      toast.success(`Đã nạp File 2 (Bản đẹp CTĐT): ${f.name}`);
    } else if (slot === 'khgd') {
      setFileKhgdPretty(f);
      setOutNameByFile((prev) => ({
        ...prev,
        khgd: `${base}_check_note_do.xlsx`,
      }));
      toast.success(`Đã nạp File 3 (Bản đẹp KHGD): ${f.name}`);
    }
  };

  const handleRun = async () => {
    if (!fileMybk || (!fileCtdtPretty && !fileKhgdPretty)) return;
    setLoading(true);
    setError(null);
    setReviewData(null);

    try {
      const readWorkbook = (file) =>
        new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (ev) => {
            try {
              resolve(XLSX.read(ev.target.result, { type: 'array' }));
            } catch (err) {
              reject(err);
            }
          };
          reader.onerror = reject;
          reader.readAsArrayBuffer(file);
        });

      const [wbMybk, wbCtdt, wbKhgd] = await Promise.all([
        readWorkbook(fileMybk),
        fileCtdtPretty ? readWorkbook(fileCtdtPretty) : Promise.resolve(null),
        fileKhgdPretty ? readWorkbook(fileKhgdPretty) : Promise.resolve(null),
      ]);

      const findSheet = (wb, keywords) => {
        for (const kw of keywords) {
          const found = wb.SheetNames.find((s) =>
            s.toLowerCase().includes(kw.toLowerCase())
          );
          if (found) return found;
        }
        return wb.SheetNames[0];
      };

      // Tìm 2 sheet gốc trong File 1 (MyBK)
      const origCtdtSheet =
        wbMybk.SheetNames.find((s) => /kh.*i|ki.*n|bloc|knowledge|subject/i.test(s)) ||
        wbMybk.SheetNames[0];
      const origKhgdSheet =
        wbMybk.SheetNames.find(
          (s) => /k.*ho|gi|plan|schedule/i.test(s) && !/kh.*i|ki.*n|bloc/i.test(s)
        ) || (wbMybk.SheetNames.length >= 2 ? wbMybk.SheetNames[1] : null);

      const origCtdtRows = origCtdtSheet ? sheetToRows(wbMybk, origCtdtSheet) : null;
      const origKhgdRows =
        origKhgdSheet && origKhgdSheet !== origCtdtSheet
          ? sheetToRows(wbMybk, origKhgdSheet)
          : !wbCtdt && wbKhgd
          ? sheetToRows(wbMybk, wbMybk.SheetNames[0])
          : null;

      // 1. Chạy kiểm toán nội bộ trên chính File 1 (Gốc MyBK) -> Phục vụ Khung Note báo Cô
      const mybkAudit = runMybkInternalAudit(origCtdtRows, origKhgdRows);

      // 2. Nếu có File 2 (Bản đẹp CTĐT) -> Đối chiếu CTĐT
      let ctdtResult = null;
      if (wbCtdt) {
        if (!origCtdtRows) {
          throw new Error(
            `Không tìm thấy sheet Khối kiến thức trong File gốc MyBK (${wbMybk.SheetNames.join(
              ', '
            )}).`
          );
        }
        const ctdtPrettySheet = findSheet(wbCtdt, ['Sheet1', 'sheet1', 'ctđt', 'ctdt', 'kh']);
        const ctdtPrettyRows = sheetToRows(wbCtdt, ctdtPrettySheet);
        ctdtResult = {
          ...runSingleModeComparison({
            origRows: origCtdtRows,
            prettyRows: ctdtPrettyRows,
            mode: 'ctdt',
            fileKey: 'ctdt',
            fileLabel: 'Bản đẹp CTĐT',
            mybkAudit,
          }),
          prettySheetName: ctdtPrettySheet,
          origSheetName: origCtdtSheet,
        };
      }

      // 3. Nếu có File 3 (Bản đẹp KHGD) -> Đối chiếu KHGD
      let khgdResult = null;
      if (wbKhgd) {
        const effectiveOrigKhgd = origKhgdRows || sheetToRows(wbMybk, wbMybk.SheetNames[0]);
        if (!effectiveOrigKhgd) {
          throw new Error(
            `Không tìm thấy sheet Kế hoạch giảng dạy trong File gốc MyBK (${wbMybk.SheetNames.join(
              ', '
            )}).`
          );
        }
        const khgdPrettySheet = findSheet(wbKhgd, [
          'Trang',
          'trang',
          'khgd',
          'kế hoạch',
          'Sheet1',
          'sheet',
        ]);
        const khgdPrettyRows = sheetToRows(wbKhgd, khgdPrettySheet);
        khgdResult = {
          ...runSingleModeComparison({
            origRows: effectiveOrigKhgd,
            prettyRows: khgdPrettyRows,
            mode: 'khgd',
            fileKey: 'khgd',
            fileLabel: 'Bản đẹp KHGD',
            mybkAudit,
          }),
          prettySheetName: khgdPrettySheet,
          origSheetName: origKhgdSheet || wbMybk.SheetNames[0],
        };
      }

      // 4. Nếu tải lên cả 2 Bản đẹp (CTĐT + KHGD) -> Đối chiếu chéo giữa 2 Bản đẹp
      const prettyCrossDiffs = runPrettyCrossCheck(ctdtResult, khgdResult);

      // Khởi tạo tick chọn sẵn các mục lỗi MyBK để báo Cô
      const initSelected = {};
      mybkAudit.reportItems.forEach((it) => {
        initSelected[it.id] = true;
      });
      setMybkSelectedIds(initSelected);
      setProfDraftEdited(null);

      const defaultScope = ctdtResult ? 'ctdt' : 'khgd';
      setActiveFileScope(defaultScope);
      setReviewData({
        ctdtResult,
        khgdResult,
        mybkAudit,
        prettyCrossDiffs,
      });
      setActiveStep(stepOrder[0]);
      setRowFilter('all');
      const targetRes = defaultScope === 'ctdt' ? ctdtResult : khgdResult;
      setPrereqFilter(targetRes?.prereqMismatches.length > 0 ? 'mismatch' : 'all');

      toast.success('Hoàn tất phân tích & đối chiếu dữ liệu!', {
        description: `Đã tách riêng lỗi Bản đẹp và ${mybkAudit.totalIssues} mục cần báo Cô trên MyBK.`,
      });
    } catch (err) {
      setError(err.message);
      toast.error(`Lỗi phân tích: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  // Result đang hiển thị theo tab File Bản đẹp đang chọn ('ctdt' hoặc 'khgd')
  const currentRes =
    activeFileScope === 'ctdt'
      ? reviewData?.ctdtResult || reviewData?.khgdResult
      : reviewData?.khgdResult || reviewData?.ctdtResult;

  const currentScopeKey = currentRes?.fileKey || 'ctdt';
  const currentNotes = notesByFile[currentScopeKey] || {};
  const currentChecked = checkedByFile[currentScopeKey] || {};

  const updateNoteForFile = (fKey, code, val) => {
    setNotesByFile((prev) => ({
      ...prev,
      [fKey]: { ...(prev[fKey] || {}), [code]: val },
    }));
  };

  const toggleCheckForFile = (fKey, code) => {
    setCheckedByFile((prev) => ({
      ...prev,
      [fKey]: { ...(prev[fKey] || {}), [code]: !prev[fKey]?.[code] },
    }));
  };

  const addCustomMybkNote = (text) => {
    const cleanTxt = text.trim();
    if (!cleanTxt) return;
    const id = `custom_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    setMybkCustomItems((prev) => [...prev, { id, category: 'Ghi chú thêm', text: cleanTxt }]);
    setMybkSelectedIds((prev) => ({ ...prev, [id]: true }));
    setProfDraftEdited(null);
    toast.success('Đã thêm mục ghi chú vào Khung Note Báo Cô!');
  };

  const scrollToProfPanel = () => {
    profPanelRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  // Sinh văn bản báo Cô từ các mục MyBK đã chọn
  const buildProfessorMessage = () => {
    if (!reviewData?.mybkAudit) return '';
    const programName =
      reviewData.mybkAudit.ctdtParsed?.meta?.programTitle ||
      reviewData.mybkAudit.khgdParsed?.meta?.programTitle ||
      'Chương trình đào tạo';

    const allItems = [
      ...reviewData.mybkAudit.reportItems,
      ...mybkCustomItems,
    ].filter((it) => mybkSelectedIds[it.id]);

    if (allItems.length === 0) {
      return `Dạ em chào Cô ạ,\nEm đã rà soát file gốc xuất từ hệ thống MyBK (${programName}), hiện tại chưa có ghi chú bất cập nào trên file hệ thống MyBK ạ.`;
    }

    const lines = [
      `Dạ em chào Cô ạ,`,
      `Trong quá trình đối chiếu và làm bản đẹp cho "${programName}", em phát hiện một số điểm trên FILE GỐC HỆ THỐNG MYBK cần xin ý kiến Cô kiểm tra / xác nhận ạ:`,
      '',
    ];
    allItems.forEach((it, idx) => {
      lines.push(`${idx + 1}. [${it.category}] ${it.text}`);
    });
    lines.push(
      '',
      'Nhờ Cô xem giúp em các mục trên để em đồng bộ chính xác sang file bản đẹp CTĐT và KHGD ạ. Em cảm ơn Cô!'
    );
    return lines.join('\n');
  };

  const finalProfText =
    profDraftEdited !== null ? profDraftEdited : buildProfessorMessage();

  const handleCopyProfNote = async () => {
    try {
      await navigator.clipboard.writeText(finalProfText);
      setCopiedProf(true);
      toast.success('Đã sao chép nội dung báo Cô vào Clipboard!');
      setTimeout(() => setCopiedProf(false), 2500);
    } catch {
      toast.error('Không thể copy tự động, vui lòng bôi đen văn bản.');
    }
  };

  const handleDownloadProfTxt = () => {
    const blob = new Blob(['\uFEFF' + finalProfText], {
      type: 'text/plain;charset=utf-8;',
    });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'Note_File_MyBK_Bao_Co.txt';
    link.click();
    toast.success('Đã tải xuống file Note_File_MyBK_Bao_Co.txt');
  };

  const handleDownloadPrettyCopyFor = async (targetKey) => {
    const targetRes =
      targetKey === 'ctdt' ? reviewData?.ctdtResult : reviewData?.khgdResult;
    const targetFile = targetKey === 'ctdt' ? fileCtdtPretty : fileKhgdPretty;
    if (!targetRes || !targetFile) return;

    setExportingKey(targetKey);
    try {
      await exportAnnotatedPrettyCopy({
        filePretty: targetFile,
        prettyRows: targetRes.prettyRows,
        prettySheetName: targetRes.prettySheetName,
        comparisonResult: targetRes,
        userNotes: notesByFile[targetKey] || {},
        checkedCourses: checkedByFile[targetKey] || {},
        generalUserNote: generalNoteByFile[targetKey] || '',
        customFileName: outNameByFile[targetKey] || '',
        options: {
          markStyle,
          fillMissingWithRedHint,
          includeSummarySheet,
        },
      });
      toast.success(`Đã xuất bản sao tô đỏ cho ${targetRes.fileLabel}!`);
    } catch (err) {
      setError(`Lỗi khi xuất bản sao ${targetRes.fileLabel}: ${err.message}`);
      toast.error(`Lỗi xuất file: ${err.message}`);
    } finally {
      setExportingKey(null);
    }
  };

  const handleExportCsv = () => {
    if (!currentRes) return;
    const rows = [
      [
        'File đang kiểm',
        'Dòng',
        'Loại',
        'Mã môn',
        'Tên môn học',
        'TC',
        'Phân loại (BB/TC)',
        'Học phần ĐK (Gốc MyBK)',
        `Học phần ĐK (${currentRes.fileLabel})`,
        'Chi tiết lệch ĐK',
        'Kết quả',
        'Nhóm/HK (Gốc)',
        `Nhóm/HK (${currentRes.fileLabel})`,
        'Ghi chú sửa Bản đẹp',
      ],
    ];

    currentRes.rowResults
      .filter((r) => r.type === 'vi' || r.type === 'en')
      .forEach((r) => {
        const codeKey = r.code || r.parentCode || '';
        rows.push([
          currentRes.fileLabel,
          r.row,
          r.type,
          r.code,
          r.name,
          r.credits,
          r.origCat || '',
          r.prereq || '',
          r.pEntry?.prereq || '',
          r.prereqDiff?.summary || '',
          r.label,
          r.origGrp || '',
          r.prettyGrp || '',
          codeKey ? currentNotes[codeKey] || '' : '',
        ]);
      });

    const csvContent = rows
      .map((r) => r.map((cell) => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(','))
      .join('\n');

    const link = document.createElement('a');
    link.href = URL.createObjectURL(
      new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' })
    );
    link.download = `ra_soat_${currentScopeKey}.csv`;
    link.click();
    toast.success(`Đã xuất báo cáo CSV cho ${currentRes.fileLabel}`);
  };

  // Bộ lọc tìm kiếm nhanh toàn cục (không làm thay đổi logic bộ lọc gốc khi để trống)
  const normSearch = searchQuery.trim().toLowerCase();
  const matchesQuery = (code, name, extra = '') => {
    if (!normSearch) return true;
    return (
      String(code || '').toLowerCase().includes(normSearch) ||
      String(name || '').toLowerCase().includes(normSearch) ||
      String(extra || '').toLowerCase().includes(normSearch)
    );
  };

  const filteredRows = currentRes?.rowResults.filter((r) => {
    if (r.type === 'group' || r.type === 'separator') {
      return rowFilter === 'all' && !normSearch;
    }
    if (!matchesQuery(r.code || r.parentCode, r.name, r.pEntry?.name)) return false;
    if (rowFilter === 'all') return true;
    if (rowFilter === 'wrong-group') return r.groupMismatch;
    if (rowFilter === 'wrong-prereq') return r.prereqMismatch;
    return r.status === rowFilter;
  });

  const filteredSpelling = currentRes?.spellingIssues.filter((sp) => {
    if (!matchesQuery(sp.code, sp.text, sp.detail)) return false;
    if (spellFilter === 'all') return true;
    if (spellFilter === 'pretty_only') return !sp.alsoInMybk;
    if (spellFilter === 'from_mybk') return sp.alsoInMybk;
    return sp.severity === spellFilter;
  });

  const filteredPrereqCourses = currentRes?.allPrereqCourses.filter((r) => {
    if (!matchesQuery(r.code, r.name, r.prereq)) return false;
    if (prereqFilter === 'mismatch') return r.prereqMismatch;
    if (prereqFilter === 'tq-only') {
      const hasOrigTQ = r.prereqItems?.some((i) => i.type === 'TQ');
      const hasPrettyTQ = r.pEntry?.prereqItems?.some((i) => i.type === 'TQ');
      return hasOrigTQ || hasPrettyTQ;
    }
    return true;
  });

  const filteredB4Rows = currentRes?.prettyCoursesList.filter((item) => {
    if (!matchesQuery(item.code, item.name, item.eng)) return false;
    if (b4Filter === 'issues') return item.severity !== 'ok';
    if (b4Filter === 'noted') {
      return Boolean(
        currentNotes[item.code]?.trim() || currentChecked[item.code]
      );
    }
    return true;
  });

  const stats = currentRes?.stats;
  const resMode = currentRes?.mode || 'ctdt';
  const currentStepIdx = stepOrder.indexOf(activeStep);
  const notedCount = Object.values(currentNotes).filter((v) => v && v.trim()).length;
  const checkedCount = Object.values(currentChecked).filter(Boolean).length;

  const getStepBadge = (stepId) => {
    if (!stats) return null;
    if (stepId === 'b1') {
      return stats.b1Issues > 0 ? (
        <span className="tb tb-warn">
          <AlertTriangle size={11} /> {stats.b1Issues} lệch
        </span>
      ) : (
        <span className="tb tb-ok">
          <CheckCircle2 size={11} /> Đúng nhóm
        </span>
      );
    }
    if (stepId === 'b2') {
      return stats.b2Issues > 0 ? (
        <span className="tb tb-red">
          <XCircle size={11} /> {stats.b2Issues} lỗi
        </span>
      ) : (
        <span className="tb tb-ok">
          <CheckCircle2 size={11} /> Khớp
        </span>
      );
    }
    if (stepId === 'b3') {
      return stats.b3Issues > 0 ? (
        <span className="tb tb-pink">
          <Zap size={11} /> {stats.b3Issues} cần xử lý
        </span>
      ) : (
        <span className="tb tb-ok">
          <CheckCircle2 size={11} /> Chuẩn ĐK
        </span>
      );
    }
    if (stepId === 'b4') {
      return (
        <span className="tb tb-red">
          <FileWarning size={11} />
          {stats.totalFileErrors > 0
            ? `${stats.totalFileErrors} ô tô đỏ`
            : 'Xuất bản sao'}
        </span>
      );
    }
    return null;
  };

  return (
    <>
      <Toaster
        position="top-right"
        richColors
        closeButton
        theme={theme === 'dark' ? 'dark' : 'light'}
      />

      {/* =====================================================================
          TOP STICKY COMMAND NAVBAR
      ===================================================================== */}
      <header className="top-navbar">
        <div className="top-navbar-inner">
          <div className="brand-cluster">
            <div className="brand-logo">
              <BookOpenCheck size={22} strokeWidth={2.2} />
            </div>
            <div className="brand-title-wrap">
              <div className="brand-eyebrow">
                <Sparkles size={11} /> CURRICULUM AUDIT STUDIO • V3.5
              </div>
              <div className="brand-title">
                web check siu cấp vip pro
                {reviewData && (
                  <span className="tb tb-blue" style={{ fontSize: 10 }}>
                    Đã nạp kết quả
                  </span>
                )}
              </div>
            </div>
          </div>

          <div className="nav-actions">
            {/* Thanh tìm kiếm nhanh Mã môn / Tên môn */}
            {reviewData && (
              <div className="global-search-wrap">
                <Search size={14} className="global-search-icon" />
                <input
                  type="text"
                  className="global-search-input"
                  placeholder="Tìm nhanh mã môn, tên môn..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                />
                {searchQuery && (
                  <button
                    className="global-search-clear"
                    onClick={() => setSearchQuery('')}
                    title="Xoá tìm kiếm"
                  >
                    <X size={13} />
                  </button>
                )}
              </div>
            )}

            {/* Nút thu gọn / mở rộng khung tải file khi đã có kết quả */}
            {reviewData && (
              <button
                className="fh-btn"
                onClick={() => setCollapseUploads((prev) => !prev)}
                title="Thu gọn hoặc mở rộng 3 ô chọn file Excel"
              >
                {collapseUploads ? (
                  <>
                    <ChevronDown size={14} /> Mở khung 3 File
                  </>
                ) : (
                  <>
                    <ChevronUp size={14} /> Thu gọn khung File
                  </>
                )}
              </button>
            )}

            {/* Bộ chuyển đổi giao diện (Theme Switcher) */}
            <div className="theme-switcher" role="group" aria-label="Chọn giao diện">
              <button
                className={`theme-btn ${theme === 'light' ? 'active' : ''}`}
                onClick={() => setTheme('light')}
                title="Giao diện Sáng Studio (Mặc định mới)"
              >
                <Sun size={13} /> Sáng Studio
              </button>
              <button
                className={`theme-btn ${theme === 'emerald' ? 'active' : ''}`}
                onClick={() => setTheme('emerald')}
                title="Giao diện Excel Emerald"
              >
                <Palette size={13} /> Emerald
              </button>
              <button
                className={`theme-btn ${theme === 'dark' ? 'active' : ''}`}
                onClick={() => setTheme('dark')}
                title="Giao diện Tối Obsidian Glass"
              >
                <Moon size={13} /> Tối Glass
              </button>
            </div>
          </div>
        </div>
      </header>

      <div className="app">
        {/* =====================================================================
            WORKSPACE INTRO BANNER
        ===================================================================== */}
        <div className="workspace-banner">
          <div className="workspace-banner-left">
            <div className="workspace-banner-title">
              <FolderKanban size={18} style={{ color: 'var(--accent)' }} />
              Hệ thống đối chiếu đồng thời 3 File Chương trình Đào tạo
            </div>
            <p className="workspace-banner-desc">
              Hỗ trợ tải cùng lúc <strong>3 file</strong> (1 File gốc MyBK + Bản đẹp CTĐT +
              Bản đẹp KHGD). Tự động <strong>tách riêng lỗi nằm ở file nào</strong>: File bản
              đẹp bị thiếu/sai sẽ xuất bản sao với <strong>chữ đỏ / highlight ô đỏ</strong>{' '}
              ở B4; lỗi gốc từ hệ thống MyBK được tổng hợp vào{' '}
              <strong>Khung Note riêng để báo Cô</strong>.
            </p>
          </div>
          {reviewData && (
            <div className="workspace-banner-right">
              <button className="fh-btn btn-orange" onClick={scrollToProfPanel}>
                <Send size={13} /> Khung Note Báo Cô (
                {Object.values(mybkSelectedIds).filter(Boolean).length})
              </button>
            </div>
          )}
        </div>

        {/* =====================================================================
            KHỐI TẢI LÊN 3 FILE (1 MYBK + 2 BẢN ĐẸP: CTĐT & KHGD)
        ===================================================================== */}
        <AnimatePresence initial={false}>
          {!collapseUploads && (
            <motion.div
              className="upload-section-wrap"
              initial={{ opacity: 0, height: 0 }}
              animate={{ opacity: 1, height: 'auto' }}
              exit={{ opacity: 0, height: 0 }}
              transition={{ duration: 0.22 }}
            >
              <div className="upload-grid-3">
                {/* FILE 1: GỐC MYBK */}
                <div
                  className={`upload-card card-mybk ${fileMybk ? 'has-file' : ''} ${
                    dragMybk ? 'drag-over' : ''
                  }`}
                  onDragOver={(e) => {
                    e.preventDefault();
                    setDragMybk(true);
                  }}
                  onDragLeave={() => setDragMybk(false)}
                  onDrop={onDropFile('mybk', setDragMybk)}
                >
                  {fileMybk && (
                    <button
                      className="clear-file-btn"
                      onClick={(e) => {
                        e.stopPropagation();
                        setFileMybk(null);
                        toast.info('Đã gỡ File 1 (Gốc MyBK)');
                      }}
                    >
                      <X size={12} /> Gỡ
                    </button>
                  )}
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={onPickFile('mybk')}
                  />
                  <div>
                    <div className="upload-icon-box">
                      <Building2 size={22} />
                    </div>
                    <div className="upload-label" style={{ color: 'var(--orange)' }}>
                      File 1 — Gốc MyBK (Bắt buộc)
                    </div>
                    <div className="upload-title">
                      {fileMybk ? fileMybk.name : 'Chọn File tải từ MyBK'}
                    </div>
                    <div className="upload-desc">
                      Gồm 2 sheet gốc: Khối kiến thức & Kế hoạch giảng dạy. Các bất cập từ
                      file này sẽ gom vào Khung Note báo Cô.
                    </div>
                  </div>
                  {fileMybk && (
                    <div
                      className="upload-filename"
                      style={{
                        color: 'var(--orange)',
                        background: 'var(--orange-soft)',
                      }}
                    >
                      <CheckCircle2 size={13} /> {fileMybk.name}
                    </div>
                  )}
                </div>

                {/* FILE 2: BẢN ĐẸP CTĐT */}
                <div
                  className={`upload-card card-ctdt ${fileCtdtPretty ? 'has-file' : ''} ${
                    dragCtdt ? 'drag-over' : ''
                  }`}
                  onDragOver={(e) => {
                    e.preventDefault();
                    setDragCtdt(true);
                  }}
                  onDragLeave={() => setDragCtdt(false)}
                  onDrop={onDropFile('ctdt', setDragCtdt)}
                >
                  {fileCtdtPretty && (
                    <button
                      className="clear-file-btn"
                      onClick={(e) => {
                        e.stopPropagation();
                        setFileCtdtPretty(null);
                        toast.info('Đã gỡ File 2 (Bản đẹp CTĐT)');
                      }}
                    >
                      <X size={12} /> Gỡ
                    </button>
                  )}
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={onPickFile('ctdt')}
                  />
                  <div>
                    <div className="upload-icon-box">
                      <FileSpreadsheet size={22} />
                    </div>
                    <div className="upload-label" style={{ color: 'var(--accent)' }}>
                      File 2 — Bản đẹp CTĐT (Khối kiến thức)
                    </div>
                    <div className="upload-title">
                      {fileCtdtPretty ? fileCtdtPretty.name : 'Chọn File CTDT_... bản đẹp'}
                    </div>
                    <div className="upload-desc">
                      Dò với sheet Khối kiến thức MyBK. Nếu thiếu/sai sẽ xuất bản sao tô đỏ
                      ô lỗi ở B4.
                    </div>
                  </div>
                  {fileCtdtPretty && (
                    <div
                      className="upload-filename"
                      style={{
                        color: 'var(--accent)',
                        background: 'var(--accent-soft)',
                      }}
                    >
                      <CheckCircle2 size={13} /> {fileCtdtPretty.name}
                    </div>
                  )}
                </div>

                {/* FILE 3: BẢN ĐẸP KHGD */}
                <div
                  className={`upload-card card-khgd ${fileKhgdPretty ? 'has-file' : ''} ${
                    dragKhgd ? 'drag-over' : ''
                  }`}
                  onDragOver={(e) => {
                    e.preventDefault();
                    setDragKhgd(true);
                  }}
                  onDragLeave={() => setDragKhgd(false)}
                  onDrop={onDropFile('khgd', setDragKhgd)}
                >
                  {fileKhgdPretty && (
                    <button
                      className="clear-file-btn"
                      onClick={(e) => {
                        e.stopPropagation();
                        setFileKhgdPretty(null);
                        toast.info('Đã gỡ File 3 (Bản đẹp KHGD)');
                      }}
                    >
                      <X size={12} /> Gỡ
                    </button>
                  )}
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={onPickFile('khgd')}
                  />
                  <div>
                    <div className="upload-icon-box">
                      <CalendarRange size={22} />
                    </div>
                    <div className="upload-label" style={{ color: 'var(--teal)' }}>
                      File 3 — Bản đẹp KHGD (Kế hoạch giảng dạy)
                    </div>
                    <div className="upload-title">
                      {fileKhgdPretty ? fileKhgdPretty.name : 'Chọn File KHGD_... bản đẹp'}
                    </div>
                    <div className="upload-desc">
                      Dò với sheet KHGD MyBK & đối chiếu chéo với CTĐT đẹp. Xuất bản sao
                      KHGD tô đỏ ô lỗi ở B4.
                    </div>
                  </div>
                  {fileKhgdPretty && (
                    <div
                      className="upload-filename"
                      style={{
                        color: 'var(--teal)',
                        background: 'var(--teal-soft)',
                      }}
                    >
                      <CheckCircle2 size={13} /> {fileKhgdPretty.name}
                    </div>
                  )}
                </div>
              </div>

              <button
                className="run-btn"
                onClick={handleRun}
                disabled={!fileMybk || (!fileCtdtPretty && !fileKhgdPretty) || loading}
              >
                {loading ? (
                  <>
                    <Loader2 size={18} className="spin-icon" />
                    Đang phân tích & tách lỗi theo từng file...
                  </>
                ) : fileCtdtPretty && fileKhgdPretty ? (
                  <>
                    <Play size={17} fill="currentColor" />
                    Bắt đầu rà soát cả 3 File (MyBK + Bản đẹp CTĐT + Bản đẹp KHGD)
                  </>
                ) : (
                  <>
                    <Play size={17} fill="currentColor" />
                    Bắt đầu rà soát các File đã chọn
                  </>
                )}
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {error && (
          <div className="error-box">
            <AlertTriangle size={18} style={{ flexShrink: 0, marginTop: 2 }} />
            <div>{error}</div>
          </div>
        )}

        {/* =====================================================================
            BẢNG PHÂN TÁCH LỖI RÕ RÀNG THEO TỪNG FILE: SAI Ở FILE NÀO?
        ===================================================================== */}
        {reviewData && (
          <>
            <div className="file-split-grid">
              {/* THẺ 1: TÌNH TRẠNG BẢN ĐẸP CTĐT */}
              <div
                className={`file-health-card ${
                  activeFileScope === 'ctdt' && reviewData.ctdtResult
                    ? 'active-scope'
                    : ''
                }`}
              >
                <div>
                  <div className="fh-top">
                    <span className="fh-tag fo-ctdt">
                      <FileSpreadsheet size={13} /> File 2 • Bản đẹp CTĐT
                    </span>
                    {reviewData.ctdtResult ? (
                      reviewData.ctdtResult.stats.totalFileErrors > 0 ? (
                        <span className="tb tb-red">
                          {reviewData.ctdtResult.stats.totalFileErrors} lỗi cần sửa
                        </span>
                      ) : (
                        <span className="tb tb-ok">✓ Khớp chuẩn</span>
                      )
                    ) : (
                      <span className="tb b-skip">Chưa tải lên</span>
                    )}
                  </div>
                  <div className="fh-title" title={fileCtdtPretty?.name}>
                    {fileCtdtPretty ? fileCtdtPretty.name : 'Chưa chọn Bản đẹp CTĐT'}
                  </div>
                  {reviewData.ctdtResult ? (
                    <>
                      <div className="fh-kpi-grid">
                        <div className="fh-kpi-box">
                          <span className="fh-kpi-label">Sót / Thừa môn</span>
                          <span className="fh-kpi-val">
                            {reviewData.ctdtResult.stats.miss +
                              reviewData.ctdtResult.stats.extra}
                          </span>
                        </div>
                        <div className="fh-kpi-box">
                          <span className="fh-kpi-label">Lệch Nhóm / BB-TC</span>
                          <span className="fh-kpi-val">
                            {reviewData.ctdtResult.stats.groupMismatch +
                              reviewData.ctdtResult.stats.categoryMismatch}
                          </span>
                        </div>
                        <div className="fh-kpi-box">
                          <span className="fh-kpi-label">Sai Tên/TC & Chính tả</span>
                          <span className="fh-kpi-val">
                            {reviewData.ctdtResult.stats.wrong +
                              reviewData.ctdtResult.stats.creditMismatch}{' '}
                            <span style={{ fontSize: 11, color: 'var(--text2)' }}>
                              ({reviewData.ctdtResult.stats.typoCount} CT)
                            </span>
                          </span>
                        </div>
                        <div className="fh-kpi-box">
                          <span className="fh-kpi-label">Thiếu / Sai Tiên quyết</span>
                          <span
                            className="fh-kpi-val"
                            style={{ color: 'var(--pink)' }}
                          >
                            {reviewData.ctdtResult.stats.wrongPrereq}
                          </span>
                        </div>
                      </div>
                    </>
                  ) : (
                    <div className="fh-breakdown">
                      Bạn có thể tải thêm file Bản đẹp CTĐT ở ô phía trên để rà soát đồng
                      thời.
                    </div>
                  )}
                </div>
                {reviewData.ctdtResult && (
                  <div className="fh-actions">
                    <button
                      className="fh-btn"
                      onClick={() => setActiveFileScope('ctdt')}
                    >
                      <Eye size={13} /> Xem chi tiết CTĐT đẹp
                    </button>
                    <button
                      className="fh-btn btn-red"
                      onClick={() => {
                        setActiveFileScope('ctdt');
                        setActiveStep('b4');
                      }}
                    >
                      <FileWarning size={13} /> Tạo bản sao CTĐT tô đỏ
                    </button>
                  </div>
                )}
              </div>

              {/* THẺ 2: TÌNH TRẠNG BẢN ĐẸP KHGD */}
              <div
                className={`file-health-card ${
                  activeFileScope === 'khgd' && reviewData.khgdResult
                    ? 'active-scope'
                    : ''
                }`}
              >
                <div>
                  <div className="fh-top">
                    <span className="fh-tag fo-khgd">
                      <CalendarRange size={13} /> File 3 • Bản đẹp KHGD
                    </span>
                    {reviewData.khgdResult ? (
                      reviewData.khgdResult.stats.totalFileErrors > 0 ? (
                        <span className="tb tb-red">
                          {reviewData.khgdResult.stats.totalFileErrors} lỗi cần sửa
                        </span>
                      ) : (
                        <span className="tb tb-ok">✓ Khớp chuẩn</span>
                      )
                    ) : (
                      <span className="tb b-skip">Chưa tải lên</span>
                    )}
                  </div>
                  <div className="fh-title" title={fileKhgdPretty?.name}>
                    {fileKhgdPretty ? fileKhgdPretty.name : 'Chưa chọn Bản đẹp KHGD'}
                  </div>
                  {reviewData.khgdResult ? (
                    <>
                      <div className="fh-kpi-grid">
                        <div className="fh-kpi-box">
                          <span className="fh-kpi-label">Sót / Thừa môn</span>
                          <span className="fh-kpi-val">
                            {reviewData.khgdResult.stats.miss +
                              reviewData.khgdResult.stats.extra}
                          </span>
                        </div>
                        <div className="fh-kpi-box">
                          <span className="fh-kpi-label">Sai HK / BB-TC</span>
                          <span className="fh-kpi-val">
                            {reviewData.khgdResult.stats.groupMismatch +
                              reviewData.khgdResult.stats.categoryMismatch}
                          </span>
                        </div>
                        <div className="fh-kpi-box">
                          <span className="fh-kpi-label">Sai Tên/TC & Chính tả</span>
                          <span className="fh-kpi-val">
                            {reviewData.khgdResult.stats.wrong +
                              reviewData.khgdResult.stats.creditMismatch}{' '}
                            <span style={{ fontSize: 11, color: 'var(--text2)' }}>
                              ({reviewData.khgdResult.stats.typoCount} CT)
                            </span>
                          </span>
                        </div>
                        <div className="fh-kpi-box">
                          <span className="fh-kpi-label">Thiếu / Sai Tiên quyết</span>
                          <span
                            className="fh-kpi-val"
                            style={{ color: 'var(--pink)' }}
                          >
                            {reviewData.khgdResult.stats.wrongPrereq}
                          </span>
                        </div>
                      </div>
                    </>
                  ) : (
                    <div className="fh-breakdown">
                      Tải thêm file Bản đẹp KHGD vào ô số 3 phía trên nếu bạn muốn dò luôn
                      KHGD cùng lúc.
                    </div>
                  )}
                </div>
                {reviewData.khgdResult && (
                  <div className="fh-actions">
                    <button
                      className="fh-btn"
                      onClick={() => setActiveFileScope('khgd')}
                    >
                      <Eye size={13} /> Xem chi tiết KHGD đẹp
                    </button>
                    <button
                      className="fh-btn btn-red"
                      onClick={() => {
                        setActiveFileScope('khgd');
                        setActiveStep('b4');
                      }}
                    >
                      <FileWarning size={13} /> Tạo bản sao KHGD tô đỏ
                    </button>
                  </div>
                )}
              </div>

              {/* THẺ 3: LỖI GỐC TỪ FILE MYBK -> BÁO RIÊNG VỚI CÔ */}
              <div className="file-health-card">
                <div>
                  <div className="fh-top">
                    <span className="fh-tag fo-mybk">
                      <Building2 size={13} /> File 1 • Gốc MyBK (Báo Cô)
                    </span>
                    <span className="tb tb-orange">
                      {reviewData.mybkAudit.totalIssues + mybkCustomItems.length} mục báo
                      Cô
                    </span>
                  </div>
                  <div className="fh-title" title={fileMybk?.name}>
                    {fileMybk?.name}
                  </div>
                  <div className="fh-kpi-grid">
                    <div className="fh-kpi-box">
                      <span className="fh-kpi-label">Lệch BB/TC 2 sheet</span>
                      <span className="fh-kpi-val" style={{ color: 'var(--orange)' }}>
                        {reviewData.mybkAudit.khgdMisplacedSubgroups.length} nhóm (
                        {reviewData.mybkAudit.bbInTcIssues.length} môn)
                      </span>
                    </div>
                    <div className="fh-kpi-box">
                      <span className="fh-kpi-label">Chính tả gốc MyBK</span>
                      <span className="fh-kpi-val">
                        {reviewData.mybkAudit.mybkTypos.length} lỗi
                      </span>
                    </div>
                    <div className="fh-kpi-box" style={{ gridColumn: 'span 2' }}>
                      <span className="fh-kpi-label">
                        Khác biệt danh mục / điều kiện giữa 2 sheet MyBK
                      </span>
                      <span className="fh-kpi-val">
                        {reviewData.mybkAudit.missingInKhgd.length +
                          reviewData.mybkAudit.prereqCrossIssues.length}{' '}
                        mục
                      </span>
                    </div>
                  </div>
                </div>
                <div className="fh-actions">
                  <button className="fh-btn btn-orange" onClick={scrollToProfPanel}>
                    <Send size={13} /> Mở Khung Note Báo Riêng Với Cô ↓
                  </button>
                </div>
              </div>
            </div>

            {/* CẢNH BÁO NẾU 2 FILE BẢN ĐẸP (CTĐT ĐẸP & KHGD ĐẸP) LỆCH NHAU */}
            {reviewData.prettyCrossDiffs.length > 0 && (
              <div className="alert-banner alert-warn">
                <AlertTriangle
                  size={18}
                  style={{ color: 'var(--yellow)', flexShrink: 0, marginTop: 2 }}
                />
                <div style={{ width: '100%' }}>
                  <strong>
                    Phát hiện {reviewData.prettyCrossDiffs.length} điểm chưa đồng bộ giữa
                    2 File Bản Đẹp (Bản đẹp CTĐT ↔ Bản đẹp KHGD):
                  </strong>
                  <div style={{ marginTop: 6, fontSize: 12 }}>
                    {reviewData.prettyCrossDiffs.slice(0, 6).map((d, i) => (
                      <div key={i} style={{ marginBottom: 4 }}>
                        • <code className="code-cell">{d.code}</code> ({d.name}) —{' '}
                        <span className="file-origin-tag fo-err">{d.type}</span>:{' '}
                        {d.detail}
                      </div>
                    ))}
                    {reviewData.prettyCrossDiffs.length > 6 && (
                      <div>
                        ... và {reviewData.prettyCrossDiffs.length - 6} mục khác.
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}

            {/* THANH CHUYỂN ĐỔI FILE BẢN ĐẸP ĐANG XEM TRONG CÁC BƯỚC B1 -> B4 */}
            <div className="scope-bar">
              <div className="scope-btns">
                <span
                  style={{
                    fontSize: 12,
                    color: 'var(--text2)',
                    fontWeight: 700,
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: 6,
                  }}
                >
                  <SlidersHorizontal size={14} /> Đang xem kết quả rà soát của:
                </span>
                {reviewData.ctdtResult && (
                  <button
                    className={`scope-btn ${
                      currentScopeKey === 'ctdt' ? 'active-ctdt' : ''
                    }`}
                    onClick={() => setActiveFileScope('ctdt')}
                  >
                    <FileSpreadsheet size={15} />
                    Bản đẹp CTĐT ({fileCtdtPretty?.name})
                    <span
                      className={`tb ${
                        reviewData.ctdtResult.stats.totalFileErrors > 0
                          ? 'tb-red'
                          : 'tb-ok'
                      }`}
                    >
                      {reviewData.ctdtResult.stats.totalFileErrors} lỗi
                    </span>
                  </button>
                )}
                {reviewData.khgdResult && (
                  <button
                    className={`scope-btn ${
                      currentScopeKey === 'khgd' ? 'active-khgd' : ''
                    }`}
                    onClick={() => setActiveFileScope('khgd')}
                  >
                    <CalendarRange size={15} />
                    Bản đẹp KHGD ({fileKhgdPretty?.name})
                    <span
                      className={`tb ${
                        reviewData.khgdResult.stats.totalFileErrors > 0
                          ? 'tb-red'
                          : 'tb-ok'
                      }`}
                    >
                      {reviewData.khgdResult.stats.totalFileErrors} lỗi
                    </span>
                  </button>
                )}
              </div>
              <button className="fh-btn btn-orange" onClick={scrollToProfPanel}>
                <Send size={13} /> Khung Note File MyBK báo Cô (
                {Object.values(mybkSelectedIds).filter(Boolean).length})
              </button>
            </div>
          </>
        )}

        {/* THANH QUY TRÌNH 4 BƯỚC (CÓ THỂ ĐỔI THỨ TỰ) */}
        <div className="stepper-header">
          <div className="stepper-hint">
            <Info size={14} style={{ color: 'var(--accent)' }} />
            <span>
              <strong>Quy trình 4 bước rà soát:</strong> Có thể nhấn <code>◀</code>{' '}
              <code>▶</code> để đổi thứ tự các bước theo thói quen làm việc của bạn.
            </span>
          </div>
          {stepOrder.join(',') !== DEFAULT_STEP_ORDER.join(',') && (
            <button
              className="reset-order-btn"
              onClick={() => setStepOrder(DEFAULT_STEP_ORDER)}
            >
              <RotateCcw size={12} /> Đặt lại mặc định (B1 → B4)
            </button>
          )}
        </div>

        <div className="stepper-grid">
          {stepOrder.map((stepId, idx) => {
            const meta = STEP_META[stepId];
            const StepIcon = meta.icon;
            const isActive = currentRes ? activeStep === stepId : idx === 0;
            return (
              <div
                key={stepId}
                className={`step-card ${meta.priority ? 'priority-step' : ''} ${
                  isActive ? 'active' : ''
                }`}
                onClick={() => currentRes && setActiveStep(stepId)}
              >
                <div>
                  <div className="step-top">
                    <span className="step-num-tag">
                      <StepIcon size={12} />
                      Bước {idx + 1} ({meta.code}) {meta.priority ? '★' : ''}
                    </span>
                    <div className="step-reorder" title="Đổi thứ tự bước này">
                      <button
                        className="reorder-btn"
                        disabled={idx === 0}
                        onClick={(e) => moveStep(stepId, -1, e)}
                      >
                        ◀
                      </button>
                      <button
                        className="reorder-btn"
                        disabled={idx === stepOrder.length - 1}
                        onClick={(e) => moveStep(stepId, 1, e)}
                      >
                        ▶
                      </button>
                    </div>
                  </div>
                  <div className="step-title">{meta.title}</div>
                  <div className="step-desc">{meta.desc}</div>
                </div>
                <div className="step-footer">
                  <span style={{ fontSize: 11, color: 'var(--text2)', fontWeight: 500 }}>
                    {currentRes ? currentRes.fileLabel : 'Chờ chạy'}
                  </span>
                  {getStepBadge(stepId)}
                </div>
              </div>
            );
          })}
        </div>

        {/* =====================================================================
            CHI TIẾT TỪNG BƯỚC CHO FILE BẢN ĐẸP ĐANG CHỌN (CTĐT ĐẸP HOẶC KHGD ĐẸP)
        ===================================================================== */}
        {currentRes && (
          <motion.div
            key={`${currentScopeKey}-${activeStep}`}
            initial={{ opacity: 0, y: 8 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.2 }}
          >
            {/* =================================================================
                BƯỚC B1: PHÂN NHÓM & BẮT BUỘC / TỰ CHỌN
            ================================================================= */}
            {activeStep === 'b1' && (
              <div>
                {/* 1. Lỗi phân nhóm / học kỳ / BB-TC ngay trên File Bản đẹp đang xem */}
                <div className="section-card">
                  <div className="section-head">
                    <div className="section-head-left">
                      <Layers size={16} style={{ color: 'var(--accent)' }} />
                      <span className="section-head-title">
                        Kiểm tra Phân nhóm & Bắt buộc/Tự chọn trên{' '}
                        <span
                          className={`file-origin-tag ${
                            currentScopeKey === 'ctdt' ? 'fo-ctdt' : 'fo-khgd'
                          }`}
                        >
                          {currentRes.fileLabel}
                        </span>
                      </span>
                    </div>
                    {currentRes.groupMismatches.length +
                      currentRes.categoryMismatches.length +
                      currentRes.khgdPrettyBbInTc.length >
                    0 ? (
                      <span className="tb tb-red">
                        {currentRes.groupMismatches.length +
                          currentRes.categoryMismatches.length +
                          currentRes.khgdPrettyBbInTc.length}{' '}
                        lỗi ở {currentRes.fileLabel}
                      </span>
                    ) : (
                      <span className="tb tb-ok">
                        ✓ Đúng nhóm & BB/TC trên {currentRes.fileLabel}
                      </span>
                    )}
                  </div>

                  {currentRes.groupMismatches.length === 0 &&
                  currentRes.categoryMismatches.length === 0 &&
                  currentRes.khgdPrettyBbInTc.length === 0 ? (
                    <div className="empty-state" style={{ padding: '28px 16px', border: 'none' }}>
                      <p>
                        ✅ Tất cả các môn trong <strong>{currentRes.fileLabel}</strong>{' '}
                        đều thuộc đúng {resMode === 'khgd' ? 'Học kỳ' : 'Nhóm kiến thức'}{' '}
                        và đúng mục Bắt buộc / Tự chọn.
                      </p>
                    </div>
                  ) : (
                    <div className="table-scroll">
                      <table>
                        <thead>
                          <tr>
                            <th>Sai ở File</th>
                            <th>Mã môn</th>
                            <th>Tên môn học</th>
                            <th>Chuẩn đối chiếu</th>
                            <th>Hiện tại trên {currentRes.fileLabel}</th>
                            <th>Ghi Note sửa Bản đẹp / Báo Cô</th>
                          </tr>
                        </thead>
                        <tbody>
                          {currentRes.groupMismatches
                            .filter((m) => matchesQuery(m.code, m.name))
                            .map((m, idx) => (
                              <tr key={`gm-${idx}`}>
                                <td>
                                  <span className="file-origin-tag fo-err">
                                    Sai ở: {currentRes.fileLabel}
                                  </span>
                                </td>
                                <td className="code-cell">{m.code}</td>
                                <td className="row-vi">{m.name}</td>
                                <td>
                                  <b>{m.origGrp}</b> <CategoryBadge cat={m.origCat} />
                                </td>
                                <td>
                                  <b style={{ color: 'var(--red)' }}>{m.prettyGrp}</b>{' '}
                                  <CategoryBadge cat={m.prettyCat} />
                                </td>
                                <td>
                                  <input
                                    className={`note-input ${
                                      currentNotes[m.code] ? 'has-val' : ''
                                    }`}
                                    placeholder={`Note sửa trên ${currentRes.fileLabel}...`}
                                    value={currentNotes[m.code] || ''}
                                    onChange={(e) =>
                                      updateNoteForFile(
                                        currentScopeKey,
                                        m.code,
                                        e.target.value
                                      )
                                    }
                                  />
                                </td>
                              </tr>
                            ))}
                          {currentRes.categoryMismatches
                            .filter((m) => matchesQuery(m.code, m.name))
                            .map((m, idx) => (
                              <tr key={`cm-${idx}`}>
                                <td>
                                  <span className="file-origin-tag fo-err">
                                    Sai ở: {currentRes.fileLabel}
                                  </span>
                                </td>
                                <td className="code-cell">{m.code}</td>
                                <td className="row-vi">{m.name}</td>
                                <td>
                                  {m.origGrp} <CategoryBadge cat={m.origCat} />
                                </td>
                                <td>
                                  {m.prettyGrp} <CategoryBadge cat={m.prettyCat} />
                                </td>
                                <td>
                                  <input
                                    className={`note-input ${
                                      currentNotes[m.code] ? 'has-val' : ''
                                    }`}
                                    placeholder="Note chuyển mục BB/TC..."
                                    value={currentNotes[m.code] || ''}
                                    onChange={(e) =>
                                      updateNoteForFile(
                                        currentScopeKey,
                                        m.code,
                                        e.target.value
                                      )
                                    }
                                  />
                                </td>
                              </tr>
                            ))}
                          {currentRes.khgdPrettyBbInTc
                            .filter((it) => matchesQuery(it.code, it.name))
                            .map((it, idx) => (
                              <tr key={`kbb-${idx}`}>
                                <td>
                                  <span className="file-origin-tag fo-khgd">
                                    Lưu ý ở: Bản đẹp KHGD
                                  </span>
                                </td>
                                <td className="code-cell">{it.code}</td>
                                <td className="row-vi">{it.name}</td>
                                <td>
                                  Khối kiến thức ({it.ctdtGroup}):{' '}
                                  <CategoryBadge cat={it.ctdtCategory} />
                                </td>
                                <td>
                                  {it.khgdSemester}: <CategoryBadge cat={it.khgdCategory} />
                                </td>
                                <td style={{ fontSize: 11.5, color: 'var(--yellow)', fontWeight: 500 }}>
                                  Môn [{it.ctdtCategory}] ở CTĐT nhưng nằm mục [
                                  {it.khgdCategory}] trên Bản đẹp KHGD
                                </td>
                              </tr>
                            ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>

                {/* 2. Tách riêng các điểm lệch BB/TC xuất phát từ chính FILE GỐC MYBK */}
                {reviewData.mybkAudit.bbInTcIssues.length > 0 && (
                  <div className="section-card">
                    <div className="section-head">
                      <div className="section-head-left">
                        <Building2 size={16} style={{ color: 'var(--orange)' }} />
                        <span className="section-head-title">
                          Phát hiện trên{' '}
                          <span className="file-origin-tag fo-mybk">
                            File 1 • Gốc MyBK
                          </span>
                          : Môn Bắt buộc [BB] ở Khối kiến thức nhưng MyBK xếp vào mục Tự
                          chọn [TC] bên KHGD
                        </span>
                      </div>
                      <button
                        className="fh-btn btn-orange"
                        onClick={scrollToProfPanel}
                      >
                        <Send size={13} /> Đã tự động đưa vào Khung Note Báo Cô ↓
                      </button>
                    </div>
                    <div className="table-scroll" style={{ maxHeight: 300 }}>
                      <table>
                        <thead>
                          <tr>
                            <th>Nguồn lỗi</th>
                            <th>Mã môn</th>
                            <th>Tên môn học</th>
                            <th>Sheet Khối kiến thức (MyBK)</th>
                            <th>Sheet KHGD (MyBK)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {reviewData.mybkAudit.bbInTcIssues
                            .filter((it) => matchesQuery(it.code, it.name))
                            .map((it, idx) => (
                              <tr key={idx}>
                                <td>
                                  <span className="file-origin-tag fo-mybk">
                                    Gốc MyBK (Báo Cô)
                                  </span>
                                </td>
                                <td className="code-cell">{it.code}</td>
                                <td className="row-vi">{it.name}</td>
                                <td>
                                  <b>{it.ctdtGroup}</b>{' '}
                                  <CategoryBadge cat={it.ctdtCategory} />
                                </td>
                                <td>
                                  <b>{it.khgdSemester}</b>{' '}
                                  <CategoryBadge cat={it.khgdCategory} />{' '}
                                  {it.khgdSubgroup && (
                                    <span style={{ fontSize: 11, color: 'var(--text2)' }}>
                                      ({it.khgdSubgroup})
                                    </span>
                                  )}
                                </td>
                              </tr>
                            ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* 3. Tổng quan nhóm */}
                <div className="section-card">
                  <div className="section-head">
                    <div className="section-head-left">
                      <FolderKanban size={16} style={{ color: 'var(--accent)' }} />
                      <span className="section-head-title">
                        Tổng quan các {resMode === 'khgd' ? 'Học kỳ' : 'Nhóm kiến thức'}{' '}
                        ({currentRes.fileLabel})
                      </span>
                    </div>
                  </div>
                  <div className="table-scroll">
                    <table>
                      <thead>
                        <tr>
                          <th>Mã</th>
                          <th>Tên Nhóm / Học kỳ</th>
                          <th>Phân loại (Gốc MyBK)</th>
                          <th>Phân loại ({currentRes.fileLabel})</th>
                          <th>Số môn (Gốc)</th>
                          <th>Số môn ({currentRes.fileLabel})</th>
                        </tr>
                      </thead>
                      <tbody>
                        {currentRes.groupOverview.map((g, idx) => (
                          <tr key={idx}>
                            <td>
                              <span
                                className="grp-id-badge"
                                style={{
                                  background: `${getGroupColor(g.gid, resMode)}22`,
                                  color: getGroupColor(g.gid, resMode),
                                }}
                              >
                                {g.gid}
                              </span>
                            </td>
                            <td className="row-vi">{g.label}</td>
                            <td>
                              <CategoryBadge cat={g.origCategory} />
                            </td>
                            <td>
                              <CategoryBadge cat={g.prettyCategory} />
                            </td>
                            <td style={{ fontFamily: 'var(--mono)', fontWeight: 600 }}>
                              {g.origCount}
                            </td>
                            <td style={{ fontFamily: 'var(--mono)', fontWeight: 600 }}>
                              {g.prettyCount}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* =================================================================
                BƯỚC B2: SÓT MÔN, ĐỦ TÍN CHỈ, NỘI DUNG & CHÍNH TẢ
            ================================================================= */}
            {activeStep === 'b2' && (
              <div>
                <div className="subtabs">
                  <button
                    className={`subtab-btn ${b2SubTab === 'rows' ? 'active' : ''}`}
                    onClick={() => setB2SubTab('rows')}
                  >
                    <ListChecks size={15} />
                    Đối chiếu từng dòng ({currentRes.fileLabel})
                    {stats.miss + stats.wrong > 0 ? (
                      <span className="tb tb-red">{stats.miss + stats.wrong}</span>
                    ) : (
                      <span className="tb tb-ok">✓</span>
                    )}
                  </button>
                  <button
                    className={`subtab-btn ${b2SubTab === 'missing' ? 'active' : ''}`}
                    onClick={() => setB2SubTab('missing')}
                  >
                    <Search size={15} />
                    Sót môn / Thừa môn ({currentRes.fileLabel})
                    {currentRes.missingInPretty.length +
                      currentRes.extraInPretty.length >
                    0 ? (
                      <span className="tb tb-red">
                        {currentRes.missingInPretty.length +
                          currentRes.extraInPretty.length}
                      </span>
                    ) : (
                      <span className="tb tb-ok">✓</span>
                    )}
                  </button>
                  <button
                    className={`subtab-btn ${b2SubTab === 'credits' ? 'active' : ''}`}
                    onClick={() => setB2SubTab('credits')}
                  >
                    <Layers size={15} />
                    Tín chỉ ({currentRes.fileLabel})
                    {stats.creditMismatch > 0 ? (
                      <span className="tb tb-red">{stats.creditMismatch} lệch</span>
                    ) : (
                      <span className="tb tb-ok">✓ Đủ TC</span>
                    )}
                  </button>
                  <button
                    className={`subtab-btn ${b2SubTab === 'spelling' ? 'active' : ''}`}
                    onClick={() => setB2SubTab('spelling')}
                  >
                    <SpellCheck size={15} />
                    Chính tả (Phân biệt lỗi Bản đẹp vs lỗi từ MyBK)
                    {currentRes.spellingIssues.length > 0 ? (
                      <span
                        className={`tb ${
                          stats.typoCount > 0 ? 'tb-red' : 'tb-warn'
                        }`}
                      >
                        {currentRes.spellingIssues.length}
                      </span>
                    ) : (
                      <span className="tb tb-ok">✓</span>
                    )}
                  </button>
                </div>

                {/* B2.1: ĐỐI CHIẾU TỪNG DÒNG */}
                {b2SubTab === 'rows' && (
                  <>
                    <div className="toolbar">
                      <div className="filter-row" style={{ marginBottom: 0 }}>
                        {[
                          { key: 'all', label: 'Tất cả' },
                          { key: 'match', label: 'Khớp', dot: '#059669' },
                          { key: 'miss', label: 'Thiếu môn', dot: '#e11d48' },
                          { key: 'wrong', label: 'Sai thông tin', dot: '#d97706' },
                          {
                            key: 'wrong-prereq',
                            label: 'Sai ĐK tiên quyết',
                            dot: '#db2777',
                          },
                          { key: 'wrong-en', label: 'Sai tiếng Anh', dot: '#2563eb' },
                          { key: 'wrong-group', label: 'Sai nhóm/HK', dot: '#ea580c' },
                          { key: 'na', label: '#N/A', dot: '#f43f5e' },
                        ].map((f) => (
                          <button
                            key={f.key}
                            className={`filter-tab ${rowFilter === f.key ? 'active' : ''}`}
                            onClick={() => setRowFilter(f.key)}
                          >
                            {f.dot && (
                              <span className="fdot" style={{ background: f.dot }} />
                            )}{' '}
                            {f.label}
                          </button>
                        ))}
                      </div>
                      <button className="export-btn" onClick={handleExportCsv}>
                        <ArrowDownToLine size={13} /> Xuất CSV ({currentRes.fileLabel})
                      </button>
                    </div>

                    <div className="table-wrap">
                      <div className="table-scroll">
                        <table>
                          <thead>
                            <tr>
                              <th>Dòng</th>
                              <th>Mã môn</th>
                              <th>Tên môn học (Đối chiếu Gốc ↔ {currentRes.fileLabel})</th>
                              <th>TC</th>
                              <th>Học phần ĐK / Tiên quyết</th>
                              <th>{GROUP_COL_TITLE[resMode]} (Gốc)</th>
                              <th>{GROUP_COL_TITLE[resMode]} (Đẹp)</th>
                              <th>Trạng thái & File lỗi</th>
                              <th>Note sửa {currentRes.fileLabel} / Báo Cô</th>
                            </tr>
                          </thead>
                          <tbody>
                            {!filteredRows?.length && (
                              <tr>
                                <td
                                  colSpan={9}
                                  style={{
                                    textAlign: 'center',
                                    padding: 28,
                                    color: 'var(--text2)',
                                  }}
                                >
                                  Không có dòng nào phù hợp.
                                </td>
                              </tr>
                            )}
                            {filteredRows?.map((row, idx) => {
                              if (row.type === 'group') {
                                return (
                                  <tr key={idx} className="row-group">
                                    <td
                                      style={{
                                        color: 'var(--text2)',
                                        fontFamily: 'var(--mono)',
                                        fontSize: 10,
                                      }}
                                    >
                                      {row.row}
                                    </td>
                                    <td>
                                      <span
                                        style={{
                                          fontFamily: 'var(--mono)',
                                          fontSize: 11,
                                          color: getGroupColor(row.groupId, resMode),
                                          fontWeight: 700,
                                        }}
                                      >
                                        {row.groupId}
                                      </span>
                                    </td>
                                    <td
                                      colSpan={7}
                                      style={{
                                        color: 'var(--accent)',
                                        fontSize: 12,
                                        fontWeight: 700,
                                      }}
                                    >
                                      {row.name}
                                    </td>
                                  </tr>
                                );
                              }

                              if (row.type === 'separator') {
                                return (
                                  <tr key={idx} style={{ opacity: 0.7 }}>
                                    <td
                                      style={{
                                        color: 'var(--text2)',
                                        fontFamily: 'var(--mono)',
                                        fontSize: 10,
                                      }}
                                    >
                                      {row.row}
                                    </td>
                                    <td />
                                    <td
                                      colSpan={7}
                                      style={{
                                        color: 'var(--text2)',
                                        fontSize: 11,
                                        fontStyle: 'italic',
                                      }}
                                    >
                                      {row.name}
                                    </td>
                                  </tr>
                                );
                              }

                              const badge =
                                STATUS_BADGES[row.status] || STATUS_BADGES.skip;
                              const isErr = row.status !== 'match' && row.status !== 'skip';

                              return (
                                <tr key={idx}>
                                  <td
                                    style={{
                                      color: 'var(--text2)',
                                      fontFamily: 'var(--mono)',
                                      fontSize: 10.5,
                                    }}
                                  >
                                    {row.row}
                                  </td>
                                  <td className="code-cell">{row.code}</td>
                                  <td>
                                    <div
                                      className={
                                        row.type === 'en' ? 'row-en' : 'row-vi'
                                      }
                                    >
                                      {row.name}
                                    </div>
                                    {row.status === 'wrong' &&
                                      row.pEntry &&
                                      clean(row.pEntry.name) !== clean(row.name) && (
                                        <div className="diff-hint">
                                          ↳ Trên {currentRes.fileLabel}:{' '}
                                          {row.pEntry.name}
                                        </div>
                                      )}
                                    {row.status === 'wrong-en' && row.pEntry && (
                                      <div className="diff-hint-blue">
                                        ↳ Trên {currentRes.fileLabel}: {row.pEntry.eng}
                                      </div>
                                    )}
                                  </td>
                                  <td className="cred-cell">
                                    {row.credits !== '' ? row.credits : '—'}
                                    {row.status === 'wrong' &&
                                      row.pEntry &&
                                      row.pEntry.credits != row.credits && (
                                        <div className="diff-hint">
                                          {row.pEntry.credits}
                                        </div>
                                      )}
                                  </td>
                                  <td>
                                    {row.type === 'vi' ? (
                                      <div>
                                        <PrereqPills items={row.prereqItems} />
                                        {row.prereqMismatch && (
                                          <div className="diff-hint-pink">
                                            ↳ {currentRes.fileLabel}:{' '}
                                            {row.pEntry?.prereq || 'Đang trống'}
                                            <div style={{ fontWeight: 700, marginTop: 2 }}>
                                              ⚡ {row.prereqDiff.summary}
                                            </div>
                                          </div>
                                        )}
                                      </div>
                                    ) : null}
                                  </td>
                                  <td>
                                    {row.origGrp ? (
                                      <span
                                        style={{
                                          fontFamily: 'var(--mono)',
                                          fontSize: 11,
                                          fontWeight: 700,
                                          color: getGroupColor(row.origGrp, resMode),
                                        }}
                                      >
                                        {row.origGrp}
                                      </span>
                                    ) : (
                                      '—'
                                    )}
                                  </td>
                                  <td>
                                    {row.prettyGrp ? (
                                      <span
                                        style={{
                                          fontFamily: 'var(--mono)',
                                          fontSize: 11,
                                          fontWeight: 700,
                                          color: row.groupMismatch
                                            ? 'var(--red)'
                                            : getGroupColor(row.prettyGrp, resMode),
                                        }}
                                      >
                                        {row.prettyGrp}
                                      </span>
                                    ) : (
                                      '—'
                                    )}
                                  </td>
                                  <td>
                                    <span className={`badge ${badge.cls}`}>
                                      {badge.icon} {badge.short}
                                    </span>
                                    {isErr && (
                                      <div style={{ marginTop: 4 }}>
                                        <span className="file-origin-tag fo-err">
                                          Sai ở: {currentRes.fileLabel}
                                        </span>
                                      </div>
                                    )}
                                  </td>
                                  <td>
                                    {row.type === 'vi' && row.code ? (
                                      <div
                                        style={{
                                          display: 'flex',
                                          gap: 5,
                                          alignItems: 'center',
                                        }}
                                      >
                                        <input
                                          className={`note-input ${
                                            currentNotes[row.code] ? 'has-val' : ''
                                          }`}
                                          placeholder={`Note đỏ vào ${currentRes.fileLabel}...`}
                                          value={currentNotes[row.code] || ''}
                                          onChange={(e) =>
                                            updateNoteForFile(
                                              currentScopeKey,
                                              row.code,
                                              e.target.value
                                            )
                                          }
                                        />
                                        <button
                                          className="export-btn"
                                          title="Đưa môn này vào Khung Note File MyBK để hỏi/báo Cô"
                                          onClick={() => {
                                            addCustomMybkNote(
                                              `Môn ${row.code} (${row.name}): Cần kiểm tra lại thông tin trên MyBK (${row.label}).`
                                            );
                                            scrollToProfPanel();
                                          }}
                                        >
                                          +Báo Cô
                                        </button>
                                      </div>
                                    ) : null}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </>
                )}

                {/* B2.2: SÓT MÔN / THỪA MÔN */}
                {b2SubTab === 'missing' && (
                  <div className="section-card">
                    <div className="section-head">
                      <div className="section-head-left">
                        <Search size={16} style={{ color: 'var(--accent)' }} />
                        <span className="section-head-title">
                          Rà Sót môn / Thừa môn trên{' '}
                          <span
                            className={`file-origin-tag ${
                              currentScopeKey === 'ctdt' ? 'fo-ctdt' : 'fo-khgd'
                            }`}
                          >
                            {currentRes.fileLabel}
                          </span>
                        </span>
                      </div>
                    </div>
                    {currentRes.missingInPretty.length === 0 &&
                    currentRes.extraInPretty.length === 0 ? (
                      <div className="empty-state" style={{ padding: '28px 16px', border: 'none' }}>
                        <p>
                          ✅ <strong>{currentRes.fileLabel}</strong> có đầy đủ 100% mã học
                          phần khớp với sheet tương ứng trên File gốc MyBK!
                        </p>
                      </div>
                    ) : (
                      <div className="table-scroll">
                        <table>
                          <thead>
                            <tr>
                              <th>Sai ở File</th>
                              <th>Tình trạng</th>
                              <th>Mã môn</th>
                              <th>Tên môn học</th>
                              <th>TC</th>
                              <th>Nhóm / HK</th>
                              <th>Xử lý ở B4</th>
                            </tr>
                          </thead>
                          <tbody>
                            {currentRes.missingInPretty
                              .filter((m) => matchesQuery(m.code, m.name))
                              .map((m, i) => (
                                <tr key={`m-${i}`}>
                                  <td>
                                    <span className="file-origin-tag fo-err">
                                      Thiếu ở: {currentRes.fileLabel}
                                    </span>
                                  </td>
                                  <td>
                                    <span className="badge b-miss">✗ Bị sót môn</span>
                                  </td>
                                  <td className="code-cell">{m.code}</td>
                                  <td className="row-vi">{m.name}</td>
                                  <td className="cred-cell">{m.credits}</td>
                                  <td>{m.group}</td>
                                  <td style={{ color: 'var(--red)', fontSize: 11.5, fontWeight: 500 }}>
                                    Sẽ chèn dòng cảnh báo ĐỎ cuối bản sao{' '}
                                    {currentRes.fileLabel}
                                  </td>
                                </tr>
                              ))}
                            {currentRes.extraInPretty
                              .filter((m) => matchesQuery(m.code, m.name))
                              .map((m, i) => (
                                <tr key={`e-${i}`}>
                                  <td>
                                    <span className="file-origin-tag fo-err">
                                      Thừa/Lạ ở: {currentRes.fileLabel}
                                    </span>
                                  </td>
                                  <td>
                                    <span className="badge b-wrong">
                                      ! Không có ở Gốc MyBK
                                    </span>
                                  </td>
                                  <td className="code-cell">{m.code}</td>
                                  <td className="row-vi">{m.name}</td>
                                  <td className="cred-cell">{m.credits}</td>
                                  <td>{m.group}</td>
                                  <td style={{ color: 'var(--yellow)', fontSize: 11.5, fontWeight: 500 }}>
                                    Sẽ tô ĐỎ ô Mã môn tại dòng #{m.excelRow} của bản sao
                                  </td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

                {/* B2.3: TÍN CHỈ */}
                {b2SubTab === 'credits' && (
                  <>
                    {currentRes.calculatedProgramCredits != null && (
                      <div
                        className={`alert-banner ${
                          currentRes.calculatedProgramCredits ===
                          currentRes.declaredProgramCredits
                            ? 'alert-ok'
                            : 'alert-warn'
                        }`}
                      >
                        <GraduationCap size={18} style={{ flexShrink: 0, marginTop: 2 }} />
                        <div>
                          <strong>
                            Tổng tín chỉ toàn chương trình trên {currentRes.fileLabel}:
                          </strong>{' '}
                          Quy định <strong>{currentRes.declaredProgramCredits} TC</strong>{' '}
                          — Tổng các khối khai báo ={' '}
                          <strong>{currentRes.calculatedProgramCredits} TC</strong>{' '}
                          {currentRes.calculatedProgramCredits ===
                          currentRes.declaredProgramCredits
                            ? '✓ (Đủ tín chỉ!)'
                            : '✗ (Lệch tổng tín chỉ — sẽ tô đỏ trên bản sao!)'}
                        </div>
                      </div>
                    )}

                    <div className="grp-table-wrap">
                      <div className="grp-header">
                        <div className="grp-header-title">
                          <Layers size={16} style={{ color: 'var(--accent)' }} />
                          Đối chiếu tín chỉ từng{' '}
                          {resMode === 'khgd' ? 'Học kỳ' : 'Khối kiến thức'} trên{' '}
                          {currentRes.fileLabel}
                        </div>
                      </div>
                      <div className="grp-grid grp-th">
                        <div>{resMode === 'khgd' ? 'Học kỳ' : 'Nhóm'}</div>
                        <div>Phân loại</div>
                        <div>TC gốc (tổng môn)</div>
                        <div>TC đẹp (khai báo)</div>
                        <div>TC đẹp (tổng môn)</div>
                        <div>Kết quả</div>
                      </div>
                      {currentRes.groupCredits.map((g, idx) => {
                        const isBad =
                          g.origSum != null && g.prettyRef != null && !g.match;
                        return (
                          <div
                            key={idx}
                            className="grp-grid grp-row-data"
                            style={isBad ? { background: 'var(--red-soft)' } : {}}
                          >
                            <div>
                              <span
                                className="grp-id-badge"
                                style={{
                                  background: `${getGroupColor(g.gid, resMode)}22`,
                                  color: getGroupColor(g.gid, resMode),
                                }}
                              >
                                {g.gid}
                              </span>
                              <div>
                                <div style={{ fontSize: 12, fontWeight: 600 }}>{g.label}</div>
                                {g.note && (
                                  <div
                                    style={{
                                      fontSize: 10.5,
                                      color: 'var(--text2)',
                                      marginTop: 2,
                                    }}
                                  >
                                    {g.note}
                                  </div>
                                )}
                              </div>
                            </div>
                            <div>
                              <CategoryBadge cat={g.category} />
                            </div>
                            <div className={isBad ? 'num-bad' : 'num-ok'}>
                              {g.origSum ?? '—'}
                            </div>
                            <div className={isBad ? 'num-bad' : 'num-ok'}>
                              {g.prettyDec ?? '—'}
                            </div>
                            <div className="num-neutral">{g.prettySum ?? '—'}</div>
                            <div>
                              {g.match ? (
                                <span className="badge b-match">✓ Hợp lệ</span>
                              ) : (
                                <span className="badge b-miss">
                                  ✗ Sai ở {currentRes.fileLabel}
                                </span>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </>
                )}

                {/* B2.4: CHÍNH TẢ (TÁCH RÕ LỖI TỪ GỐC MYBK HAY LỖI RIÊNG Ở BẢN ĐẸP) */}
                {b2SubTab === 'spelling' && (
                  <div className="section-card">
                    <div className="section-head">
                      <div className="section-head-left">
                        <SpellCheck size={16} style={{ color: 'var(--accent)' }} />
                        <span className="section-head-title">
                          Rà Chính tả & Định dạng trên {currentRes.fileLabel} (Tách rõ
                          nguồn lỗi)
                        </span>
                      </div>
                      <div className="filter-row" style={{ marginBottom: 0 }}>
                        {[
                          {
                            key: 'all',
                            label: `Tất cả (${currentRes.spellingIssues.length})`,
                          },
                          {
                            key: 'pretty_only',
                            label: `Sai riêng ở ${currentRes.fileLabel} (${
                              currentRes.spellingIssues.filter((x) => !x.alsoInMybk)
                                .length
                            })`,
                            dot: '#e11d48',
                          },
                          {
                            key: 'from_mybk',
                            label: `Lỗi có từ Gốc MyBK (${
                              currentRes.spellingIssues.filter((x) => x.alsoInMybk)
                                .length
                            })`,
                            dot: '#ea580c',
                          },
                          {
                            key: 'typo',
                            label: `Chỉ lỗi từ vựng (${stats.typoCount})`,
                            dot: '#db2777',
                          },
                        ].map((f) => (
                          <button
                            key={f.key}
                            className={`filter-tab ${
                              spellFilter === f.key ? 'active' : ''
                            }`}
                            onClick={() => setSpellFilter(f.key)}
                          >
                            {f.dot && (
                              <span className="fdot" style={{ background: f.dot }} />
                            )}{' '}
                            {f.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    {filteredSpelling.length === 0 ? (
                      <div className="empty-state" style={{ border: 'none' }}>
                        <div className="icon-wrap">
                          <Sparkles size={24} />
                        </div>
                        <p>Không có lỗi chính tả nào trong bộ lọc này.</p>
                      </div>
                    ) : (
                      <div className="table-scroll">
                        <table>
                          <thead>
                            <tr>
                              <th>Nguồn gốc lỗi</th>
                              <th>Mã môn</th>
                              <th>Nội dung trên {currentRes.fileLabel}</th>
                              <th>Chi tiết & Gợi ý sửa</th>
                              <th>Hành động</th>
                            </tr>
                          </thead>
                          <tbody>
                            {filteredSpelling.map((sp, idx) => (
                              <tr key={idx}>
                                <td>
                                  <span
                                    className={`file-origin-tag ${
                                      sp.alsoInMybk ? 'fo-mybk' : 'fo-err'
                                    }`}
                                  >
                                    {sp.originBadge}
                                  </span>
                                </td>
                                <td className="code-cell">{sp.code}</td>
                                <td
                                  style={{
                                    fontFamily: 'var(--mono)',
                                    fontSize: 11.5,
                                  }}
                                >
                                  [{sp.lang}] {sp.text}
                                </td>
                                <td>
                                  <div style={{ color: 'var(--red)', fontSize: 12, fontWeight: 500 }}>
                                    {sp.detail}
                                  </div>
                                  {sp.suggestion && (
                                    <div
                                      style={{
                                        fontSize: 11.5,
                                        color: 'var(--green)',
                                        marginTop: 3,
                                        fontWeight: 500,
                                      }}
                                    >
                                      ↳ Gợi ý: <code>{sp.suggestion}</code>
                                    </div>
                                  )}
                                </td>
                                <td>
                                  <div style={{ display: 'flex', gap: 6 }}>
                                    <button
                                      className="export-btn"
                                      onClick={() => {
                                        const prev = currentNotes[sp.code]
                                          ? `${currentNotes[sp.code]}; `
                                          : '';
                                        updateNoteForFile(
                                          currentScopeKey,
                                          sp.code,
                                          `${prev}${sp.detail}`
                                        );
                                        toast.success(`Đã thêm Note đỏ cho môn ${sp.code}`);
                                      }}
                                    >
                                      +Note đỏ Bản đẹp
                                    </button>
                                    {sp.alsoInMybk && (
                                      <button
                                        className="export-btn"
                                        onClick={scrollToProfPanel}
                                      >
                                        Xem ở Note báo Cô
                                      </button>
                                    )}
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* =================================================================
                BƯỚC B3: TIÊN QUYẾT (QUAN TRỌNG NHẤT), MÔN BẮT BUỘC, GHI CHÚ & ĐK TỐT NGHIỆP
            ================================================================= */}
            {activeStep === 'b3' && (
              <div>
                {currentRes.prereqMismatches.length > 0 && (
                  <div className="alert-banner alert-pink">
                    <Zap size={18} style={{ color: 'var(--pink)', flexShrink: 0, marginTop: 2 }} />
                    <div>
                      <strong>
                        Phát hiện {currentRes.prereqMismatches.length} môn bị thiếu / sai
                        Học phần Tiên quyết trên{' '}
                        <span style={{ textDecoration: 'underline' }}>
                          {currentRes.fileLabel}
                        </span>
                        !
                      </strong>{' '}
                      Các ô Tiên quyết bị thiếu/sai này sẽ tự động được{' '}
                      <strong>tô highlight đỏ & ghi chữ đỏ</strong> khi bạn xuất bản sao{' '}
                      {currentRes.fileLabel} ở Bước B4.
                    </div>
                  </div>
                )}

                {currentRes.legendWarnings.map((lw, idx) => (
                  <div key={idx} className="alert-banner alert-warn">
                    <AlertTriangle size={18} style={{ color: 'var(--yellow)', flexShrink: 0, marginTop: 2 }} />
                    <div>
                      <strong>
                        Lưu ý dòng Ghi chú (notes) trên {lw.fileLabel}:
                      </strong>{' '}
                      {lw.text}
                    </div>
                  </div>
                ))}

                {/* 3.1: TIÊN QUYẾT */}
                <div className="section-card">
                  <div className="section-head">
                    <div className="section-head-left">
                      <ShieldAlert size={16} style={{ color: 'var(--pink)' }} />
                      <span className="section-head-title">
                        1. Rà Tiên quyết / Học phần điều kiện trên{' '}
                        <span
                          className={`file-origin-tag ${
                            currentScopeKey === 'ctdt' ? 'fo-ctdt' : 'fo-khgd'
                          }`}
                        >
                          {currentRes.fileLabel}
                        </span>
                      </span>
                    </div>
                    <div className="filter-row" style={{ marginBottom: 0 }}>
                      {[
                        {
                          key: 'all',
                          label: `Tất cả môn có ĐK (${currentRes.allPrereqCourses.length})`,
                        },
                        {
                          key: 'mismatch',
                          label: `Thiếu/Sai ở ${currentRes.fileLabel} (${currentRes.prereqMismatches.length})`,
                          dot: '#e11d48',
                        },
                        {
                          key: 'tq-only',
                          label: 'Chỉ môn có Tiên quyết (TQ)',
                          dot: '#db2777',
                        },
                      ].map((f) => (
                        <button
                          key={f.key}
                          className={`filter-tab ${
                            prereqFilter === f.key ? 'active' : ''
                          }`}
                          onClick={() => setPrereqFilter(f.key)}
                        >
                          {f.dot && (
                            <span className="fdot" style={{ background: f.dot }} />
                          )}{' '}
                          {f.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  {filteredPrereqCourses.length === 0 ? (
                    <div className="empty-state" style={{ border: 'none' }}>
                      <div className="icon-wrap">
                        <CheckCircle2 size={24} />
                      </div>
                      <p>
                        Tất cả học phần điều kiện / tiên quyết trên{' '}
                        <strong>{currentRes.fileLabel}</strong> đều khớp chuẩn với File
                        gốc MyBK.
                      </p>
                    </div>
                  ) : (
                    <div className="table-scroll">
                      <table>
                        <thead>
                          <tr>
                            <th>Sai ở File</th>
                            <th>Mã môn</th>
                            <th>Tên môn học</th>
                            <th>ĐK File gốc MyBK</th>
                            <th>ĐK trên {currentRes.fileLabel}</th>
                            <th>Kiểm tra phạm vi & Học kỳ</th>
                            <th>Note đỏ vào {currentRes.fileLabel}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredPrereqCourses.map((item, idx) => (
                            <tr key={idx}>
                              <td>
                                {item.prereqMismatch ? (
                                  <span className="file-origin-tag fo-err">
                                    Thiếu/Sai ở: {currentRes.fileLabel}
                                  </span>
                                ) : (
                                  <span className="badge b-match">✓ Khớp</span>
                                )}
                              </td>
                              <td className="code-cell">{item.code}</td>
                              <td className="row-vi">{item.name}</td>
                              <td>
                                <PrereqPills
                                  items={item.prereqItems}
                                  emptyText="Không có"
                                />
                              </td>
                              <td>
                                {item.pEntry ? (
                                  <div>
                                    <PrereqPills
                                      items={item.pEntry.prereqItems}
                                      emptyText="ĐANG BỊ TRỐNG"
                                    />
                                    {item.prereqMismatch && (
                                      <div
                                        style={{
                                          color: 'var(--red)',
                                          fontSize: 11,
                                          fontWeight: 700,
                                          marginTop: 4,
                                        }}
                                      >
                                        ⚡ {item.prereqDiff.summary}
                                      </div>
                                    )}
                                  </div>
                                ) : (
                                  <span className="badge b-miss">✗ Thiếu môn</span>
                                )}
                              </td>
                              <td style={{ fontSize: 11.5 }}>
                                {(item.prereqAnalysis || []).map((pa, i) => (
                                  <div key={i} style={{ marginBottom: 3 }}>
                                    <code className="code-cell">{pa.code}</code>:{' '}
                                    {pa.scopeLabel}
                                    {pa.semesterCheck && (
                                      <div
                                        style={{
                                          color: pa.semesterCheck.ok
                                            ? 'var(--green)'
                                            : 'var(--red)',
                                          fontWeight: 600,
                                        }}
                                      >
                                        ↳ {pa.semesterCheck.msg}
                                      </div>
                                    )}
                                  </div>
                                ))}
                              </td>
                              <td>
                                <input
                                  className={`note-input ${
                                    currentNotes[item.code] ? 'has-val' : ''
                                  }`}
                                  placeholder={`Note tiên quyết trên ${currentRes.fileLabel}...`}
                                  value={currentNotes[item.code] || ''}
                                  onChange={(e) =>
                                    updateNoteForFile(
                                      currentScopeKey,
                                      item.code,
                                      e.target.value
                                    )
                                  }
                                />
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>

                {/* 3.2: ĐIỀU KIỆN HOÀN THÀNH CHƯƠNG TRÌNH (ENG_B2, ĐỀ ÁN, THỰC TẬP, NGOẠI NGỮ) */}
                <div className="section-card">
                  <div className="section-head">
                    <div className="section-head-left">
                      <GraduationCap size={16} style={{ color: 'var(--accent)' }} />
                      <span className="section-head-title">
                        2. Yêu cầu hoàn thành chương trình, Ngoại ngữ & Điều kiện tốt
                        nghiệp (ENG_B2 - Đ/kiện ngoại ngữ tốt nghiệp SĐH) trên{' '}
                        {currentRes.fileLabel}
                      </span>
                    </div>
                  </div>
                  <div className="table-scroll">
                    <table>
                      <thead>
                        <tr>
                          <th>Nhóm</th>
                          <th>Mã HP / ĐK</th>
                          <th>Tên học phần / Điều kiện</th>
                          <th>TC</th>
                          <th>Tiên quyết / Song hành</th>
                          <th>Trạng thái trên {currentRes.fileLabel}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {currentRes.specialRequirements
                          .filter((sr) => matchesQuery(sr.code, sr.name, sr.eng))
                          .map((sr, idx) => (
                            <tr key={idx}>
                              <td>
                                <b>{sr.group}</b>
                              </td>
                              <td className="code-cell">{sr.code}</td>
                              <td>
                                <div className="row-vi">{sr.name}</div>
                                {sr.eng && <div className="row-en">{sr.eng}</div>}
                              </td>
                              <td className="cred-cell">{sr.credits}</td>
                              <td style={{ fontSize: 11.5 }}>
                                Gốc: <code>{sr.origPrereq || 'Không'}</code>
                                {sr.prereqMismatch && (
                                  <div style={{ color: 'var(--red)', fontWeight: 700 }}>
                                    {currentRes.fileLabel}:{' '}
                                    <code>{sr.prettyPrereq || 'TRỐNG'}</code> (
                                    {sr.prereqSummary})
                                  </div>
                                )}
                              </td>
                              <td>
                                {!sr.inPretty ? (
                                  <span className="file-origin-tag fo-err">
                                    ✗ Bị thiếu ở {currentRes.fileLabel}
                                  </span>
                                ) : sr.prereqMismatch ? (
                                  <span className="file-origin-tag fo-err">
                                    ⚡ Thiếu ĐK trên {currentRes.fileLabel}
                                  </span>
                                ) : (
                                  <span className="badge b-match">
                                    ✓ Đầy đủ trên {currentRes.fileLabel}
                                  </span>
                                )}
                              </td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* =================================================================
                BƯỚC B4: TẠO BẢN SAO FILE ĐẸP (CTĐT & KHGD) TÔ CHỮ ĐỎ / HIGHLIGHT Ô ĐỎ
            ================================================================= */}
            {activeStep === 'b4' && (
              <div>
                <div className="b4-config-grid">
                  <div className="section-card" style={{ marginBottom: 0 }}>
                    <div className="section-head">
                      <div className="section-head-left">
                        <FileWarning size={16} style={{ color: 'var(--red)' }} />
                        <span className="section-head-title">
                          Tạo Bản Sao Tô Đỏ cho{' '}
                          <span
                            className={`file-origin-tag ${
                              currentScopeKey === 'ctdt' ? 'fo-ctdt' : 'fo-khgd'
                            }`}
                          >
                            {currentRes.fileLabel}
                          </span>
                        </span>
                      </div>
                    </div>
                    <div className="section-body">
                      <p
                        style={{
                          fontSize: 12.5,
                          color: 'var(--text2)',
                          lineHeight: 1.65,
                          marginBottom: 14,
                        }}
                      >
                        Bản sao giữ nguyên 100% bố cục trình bày của{' '}
                        <strong>
                          {currentScopeKey === 'ctdt'
                            ? fileCtdtPretty?.name
                            : fileKhgdPretty?.name}
                        </strong>
                        . Mọi ô bị thiếu hoặc bị sai (Tên môn, Tín chỉ, Tiên quyết bị bỏ
                        trống, Mã môn sai nhóm/HK, Sót môn) sẽ được{' '}
                        <strong style={{ color: 'var(--red)' }}>
                          đánh dấu màu đỏ trực tiếp ngay tại ô đó
                        </strong>{' '}
                        kèm cột Note màu đỏ bên phải bảng.
                      </p>

                      {/* Chọn kiểu tô đỏ */}
                      <div style={{ marginBottom: 14 }}>
                        <div
                          style={{
                            fontSize: 12,
                            color: 'var(--text2)',
                            marginBottom: 7,
                            fontWeight: 600,
                          }}
                        >
                          Kiểu đánh dấu ô bị sai / thiếu trên bản sao Excel:
                        </div>
                        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                          {[
                            {
                              key: 'both',
                              label: '🖍️ Highlight ô màu đỏ + Chữ màu đỏ (Khuyên dùng)',
                            },
                            { key: 'red_text', label: '🔴 Chỉ đổi chữ màu đỏ đậm' },
                            {
                              key: 'highlight',
                              label: '🟥 Chỉ tô highlight ô màu đỏ',
                            },
                          ].map((st) => (
                            <button
                              key={st.key}
                              className={`option-btn ${
                                markStyle === st.key ? 'active' : ''
                              }`}
                              onClick={() => setMarkStyle(st.key)}
                            >
                              {st.label}
                            </button>
                          ))}
                        </div>
                      </div>

                      <div
                        style={{
                          display: 'flex',
                          gap: 16,
                          flexWrap: 'wrap',
                          marginBottom: 14,
                          fontSize: 12,
                        }}
                      >
                        <label
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: 6,
                            cursor: 'pointer',
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={fillMissingWithRedHint}
                            onChange={(e) =>
                              setFillMissingWithRedHint(e.target.checked)
                            }
                          />
                          Điền chữ đỏ <code>[THIẾU: ...]</code> ngay vào ô đang bị bỏ
                          trống (VD: ô Tiên quyết bị thiếu)
                        </label>
                        <label
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: 6,
                            cursor: 'pointer',
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={includeSummarySheet}
                            onChange={(e) => setIncludeSummarySheet(e.target.checked)}
                          />
                          Kèm thêm sheet tổng hợp lỗi của file này
                        </label>
                      </div>

                      <div style={{ marginBottom: 14 }}>
                        <label
                          style={{
                            display: 'block',
                            fontSize: 11.5,
                            color: 'var(--text2)',
                            marginBottom: 5,
                            fontWeight: 600,
                          }}
                        >
                          Tên file bản sao ({currentRes.fileLabel}):
                        </label>
                        <input
                          className="note-input"
                          value={outNameByFile[currentScopeKey] || ''}
                          onChange={(e) =>
                            setOutNameByFile((prev) => ({
                              ...prev,
                              [currentScopeKey]: e.target.value,
                            }))
                          }
                        />
                      </div>

                      {/* Nút tải bản sao cho cả CTĐT và KHGD */}
                      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                        {reviewData.ctdtResult && (
                          <button
                            className="red-action-btn"
                            disabled={exportingKey === 'ctdt'}
                            onClick={() => handleDownloadPrettyCopyFor('ctdt')}
                          >
                            {exportingKey === 'ctdt' ? (
                              <>
                                <Loader2 size={15} className="spin-icon" /> Đang xuất CTĐT...
                              </>
                            ) : (
                              <>
                                <Download size={15} /> Tải Bản Sao CTĐT Đẹp Tô Đỏ (
                                {reviewData.ctdtResult.stats.totalFileErrors} lỗi)
                              </>
                            )}
                          </button>
                        )}
                        {reviewData.khgdResult && (
                          <button
                            className="red-action-btn"
                            style={{
                              background: 'linear-gradient(135deg,#e11d48,#be123c)',
                            }}
                            disabled={exportingKey === 'khgd'}
                            onClick={() => handleDownloadPrettyCopyFor('khgd')}
                          >
                            {exportingKey === 'khgd' ? (
                              <>
                                <Loader2 size={15} className="spin-icon" /> Đang xuất KHGD...
                              </>
                            ) : (
                              <>
                                <Download size={15} /> Tải Bản Sao KHGD Đẹp Tô Đỏ (
                                {reviewData.khgdResult.stats.totalFileErrors} lỗi)
                              </>
                            )}
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  <div className="section-card" style={{ marginBottom: 0 }}>
                    <div className="section-head">
                      <div className="section-head-left">
                        <FileText size={16} style={{ color: 'var(--accent)' }} />
                        <span className="section-head-title">
                          Ghi chú chung đính kèm đầu bản sao ({currentRes.fileLabel})
                        </span>
                      </div>
                    </div>
                    <div className="section-body">
                      <textarea
                        className="note-input"
                        rows={6}
                        style={{ resize: 'vertical' }}
                        placeholder={`Ghi chú riêng cho ${currentRes.fileLabel} (sẽ in chữ đỏ ở góc trên bản sao Excel)...`}
                        value={generalNoteByFile[currentScopeKey] || ''}
                        onChange={(e) =>
                          setGeneralNoteByFile((prev) => ({
                            ...prev,
                            [currentScopeKey]: e.target.value,
                          }))
                        }
                      />
                      <div
                        style={{
                          marginTop: 10,
                          fontSize: 12,
                          color: 'var(--text2)',
                          display: 'flex',
                          justifyContent: 'space-between',
                        }}
                      >
                        <span>
                          Đang xem: <strong>{currentRes.fileLabel}</strong>
                        </span>
                        <span>
                          Đã check: <b>{checkedCount}</b> • Đã note: <b>{notedCount}</b>
                        </span>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Bảng dòng của File đẹp đang chọn */}
                <div className="section-card">
                  <div className="section-head">
                    <div className="section-head-left">
                      <ListChecks size={16} style={{ color: 'var(--accent)' }} />
                      <span className="section-head-title">
                        Danh sách dòng trên {currentRes.fileLabel} — Các dòng có lỗi sẽ tự
                        động tô ĐỎ ô sai/thiếu
                      </span>
                    </div>
                    <div className="filter-row" style={{ marginBottom: 0 }}>
                      {[
                        {
                          key: 'all',
                          label: `Tất cả dòng (${currentRes.prettyCoursesList.length})`,
                        },
                        {
                          key: 'issues',
                          label: `Dòng bị tô đỏ / lưu ý (${
                            currentRes.prettyCoursesList.filter(
                              (x) => x.severity !== 'ok'
                            ).length
                          })`,
                          dot: '#e11d48',
                        },
                        {
                          key: 'noted',
                          label: `Đã check / có Note (${
                            currentRes.prettyCoursesList.filter(
                              (x) =>
                                currentNotes[x.code]?.trim() ||
                                currentChecked[x.code]
                            ).length
                          })`,
                          dot: '#059669',
                        },
                      ].map((f) => (
                        <button
                          key={f.key}
                          className={`filter-tab ${b4Filter === f.key ? 'active' : ''}`}
                          onClick={() => setB4Filter(f.key)}
                        >
                          {f.dot && (
                            <span className="fdot" style={{ background: f.dot }} />
                          )}{' '}
                          {f.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="table-scroll">
                    <table>
                      <thead>
                        <tr>
                          <th>Dòng Excel</th>
                          <th>Check</th>
                          <th>Mã môn</th>
                          <th>Tên học phần ({currentRes.fileLabel})</th>
                          <th>Ô sẽ tô đỏ trên Excel</th>
                          <th>Phát hiện tại {currentRes.fileLabel}</th>
                          <th>Note sửa (In chữ đỏ trên bản sao)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredB4Rows.map((item, idx) => {
                          const ce = item.cellErrors || {};
                          const redCols = [];
                          if (ce.codeCol) redCols.push('Ô Mã môn');
                          if (ce.viNameCol) redCols.push('Ô Tên TV');
                          if (ce.enNameCol) redCols.push('Ô Tên TA');
                          if (ce.credCol) redCols.push('Ô Tín chỉ');
                          if (ce.prereqCol) redCols.push('Ô Tiên quyết');

                          return (
                            <tr key={idx}>
                              <td
                                style={{
                                  fontFamily: 'var(--mono)',
                                  fontSize: 11,
                                  color: 'var(--text2)',
                                }}
                              >
                                #{item.excelRow}
                              </td>
                              <td style={{ textAlign: 'center' }}>
                                <input
                                  type="checkbox"
                                  checked={Boolean(currentChecked[item.code])}
                                  onChange={() =>
                                    toggleCheckForFile(currentScopeKey, item.code)
                                  }
                                  style={{ width: 15, height: 15, cursor: 'pointer' }}
                                />
                              </td>
                              <td className="code-cell">{item.code}</td>
                              <td>
                                <div className="row-vi">{item.name}</div>
                                {item.eng && <div className="row-en">{item.eng}</div>}
                              </td>
                              <td>
                                {redCols.length > 0 ? (
                                  <div
                                    style={{
                                      display: 'flex',
                                      gap: 4,
                                      flexWrap: 'wrap',
                                    }}
                                  >
                                    {redCols.map((rc, i) => (
                                      <span key={i} className="file-origin-tag fo-err">
                                        🖍️ {rc}
                                      </span>
                                    ))}
                                  </div>
                                ) : (
                                  <span style={{ color: 'var(--text2)', fontSize: 11 }}>
                                    —
                                  </span>
                                )}
                              </td>
                              <td style={{ maxWidth: 310 }}>
                                {item.severity === 'ok' ? (
                                  <span className="badge b-match">✓ Khớp chuẩn</span>
                                ) : (
                                  <div
                                    style={{
                                      fontSize: 11.5,
                                      color:
                                        item.severity === 'error'
                                          ? 'var(--red)'
                                          : 'var(--yellow)',
                                      lineHeight: 1.5,
                                      fontWeight: 500,
                                    }}
                                  >
                                    {item.autoNotes.map((n, i) => (
                                      <div key={i}>• {n}</div>
                                    ))}
                                  </div>
                                )}
                              </td>
                              <td style={{ minWidth: 220 }}>
                                <input
                                  className={`note-input ${
                                    currentNotes[item.code] ? 'has-val' : ''
                                  }`}
                                  placeholder="Gõ note chữ đỏ vào dòng này..."
                                  value={currentNotes[item.code] || ''}
                                  onChange={(e) =>
                                    updateNoteForFile(
                                      currentScopeKey,
                                      item.code,
                                      e.target.value
                                    )
                                  }
                                />
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* ĐIỀU HƯỚNG GIỮA CÁC BƯỚC */}
            <div className="step-nav-footer">
              <button
                className="option-btn"
                disabled={currentStepIdx <= 0}
                onClick={() =>
                  currentStepIdx > 0 && setActiveStep(stepOrder[currentStepIdx - 1])
                }
                style={{ opacity: currentStepIdx <= 0 ? 0.35 : 1 }}
              >
                <ChevronLeft size={15} /> Bước trước
              </button>
              <span style={{ fontSize: 12.5, color: 'var(--text2)' }}>
                Đang xem <strong>{currentRes.fileLabel}</strong> •{' '}
                <strong>
                  Bước {currentStepIdx + 1}/4 ({STEP_META[activeStep].code})
                </strong>
              </span>
              <button
                className="option-btn active"
                disabled={currentStepIdx >= stepOrder.length - 1}
                onClick={() =>
                  currentStepIdx < stepOrder.length - 1 &&
                  setActiveStep(stepOrder[currentStepIdx + 1])
                }
                style={{
                  opacity: currentStepIdx >= stepOrder.length - 1 ? 0.35 : 1,
                }}
              >
                Bước tiếp theo <ChevronRight size={15} />
              </button>
            </div>
          </motion.div>
        )}

        {/* =====================================================================
            KHUNG NOTE RIÊNG CHO FILE GỐC MYBK ĐỂ BÁO RIÊNG VỚI CÔ
        ===================================================================== */}
        {reviewData && (
          <div className="professor-panel" ref={profPanelRef}>
            <div className="prof-head">
              <div>
                <div
                  style={{
                    fontSize: 11,
                    fontFamily: 'var(--mono)',
                    color: 'var(--orange)',
                    textTransform: 'uppercase',
                    letterSpacing: 1,
                    fontWeight: 700,
                    display: 'flex',
                    alignItems: 'center',
                    gap: 6,
                  }}
                >
                  <Building2 size={14} /> Tách riêng bất cập từ File 1 — Hệ thống MyBK (
                  {fileMybk?.name})
                </div>
                <div
                  style={{
                    fontSize: 16,
                    fontWeight: 800,
                    marginTop: 3,
                    color: 'var(--text)',
                  }}
                >
                  Khung Tổng Hợp Note File MyBK Để Báo Riêng Với Cô
                </div>
              </div>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <button className="fh-btn btn-orange" onClick={handleCopyProfNote}>
                  {copiedProf ? (
                    <>
                      <Check size={14} /> Đã Copy vào Clipboard!
                    </>
                  ) : (
                    <>
                      <Copy size={14} /> Copy nội dung báo Cô
                    </>
                  )}
                </button>
                <button className="fh-btn" onClick={handleDownloadProfTxt}>
                  <Download size={14} /> Tải file .txt
                </button>
              </div>
            </div>

            <div className="prof-body">
              {/* Cột trái: Danh sách các điểm phát hiện trên File MyBK + ô thêm nhanh */}
              <div>
                <div
                  style={{
                    fontSize: 12.5,
                    fontWeight: 700,
                    marginBottom: 10,
                    color: 'var(--text)',
                  }}
                >
                  Tick chọn các điểm trên File gốc MyBK muốn đưa vào tin nhắn báo Cô:
                </div>

                <div style={{ maxHeight: 270, overflowY: 'auto', paddingRight: 4 }}>
                  {[...reviewData.mybkAudit.reportItems, ...mybkCustomItems].length ===
                    0 && (
                    <div
                      style={{
                        fontSize: 12.5,
                        color: 'var(--text2)',
                        padding: '14px 0',
                      }}
                    >
                      Không phát hiện mâu thuẫn nội bộ nào trên File gốc MyBK. Bạn vẫn có
                      thể tự thêm ghi chú bên dưới để báo Cô.
                    </div>
                  )}

                  {[...reviewData.mybkAudit.reportItems, ...mybkCustomItems].map(
                    (it) => (
                      <label
                        key={it.id}
                        className="prof-item"
                        style={{ cursor: 'pointer' }}
                      >
                        <input
                          type="checkbox"
                          checked={Boolean(mybkSelectedIds[it.id])}
                          onChange={() => {
                            setMybkSelectedIds((prev) => ({
                              ...prev,
                              [it.id]: !prev[it.id],
                            }));
                            setProfDraftEdited(null);
                          }}
                          style={{ marginTop: 3, width: 15, height: 15 }}
                        />
                        <div>
                          <span
                            className="file-origin-tag fo-mybk"
                            style={{ marginRight: 6 }}
                          >
                            {it.category}
                          </span>
                          <span>{it.text}</span>
                        </div>
                      </label>
                    )
                  )}
                </div>

                {/* Ô thêm ghi chú MyBK thủ công */}
                <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
                  <input
                    className="note-input"
                    placeholder="Thêm ghi chú khác về file MyBK để báo Cô..."
                    value={newMybkCustomInput}
                    onChange={(e) => setNewMybkCustomInput(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && newMybkCustomInput.trim()) {
                        addCustomMybkNote(newMybkCustomInput);
                        setNewMybkCustomInput('');
                      }
                    }}
                  />
                  <button
                    className="fh-btn btn-orange"
                    onClick={() => {
                      if (newMybkCustomInput.trim()) {
                        addCustomMybkNote(newMybkCustomInput);
                        setNewMybkCustomInput('');
                      }
                    }}
                  >
                    <Plus size={14} /> Thêm ý
                  </button>
                </div>
              </div>

              {/* Cột phải: Nội dung tin nhắn / báo cáo hoàn chỉnh để Copy gửi Cô */}
              <div>
                <div
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    marginBottom: 10,
                  }}
                >
                  <span
                    style={{ fontSize: 12.5, fontWeight: 700, color: 'var(--text)' }}
                  >
                    Nội dung tin nhắn / ghi chú gửi Cô (có thể sửa trực tiếp):
                  </span>
                  {profDraftEdited !== null && (
                    <button
                      className="reset-order-btn"
                      onClick={() => setProfDraftEdited(null)}
                    >
                      <RotateCcw size={12} /> Tạo lại tự động
                    </button>
                  )}
                </div>
                <textarea
                  className="prof-textarea"
                  value={finalProfText}
                  onChange={(e) => setProfDraftEdited(e.target.value)}
                />
              </div>
            </div>
          </div>
        )}

        {!reviewData && !loading && (
          <div className="empty-state">
            <div className="icon-wrap">
              <BookOpenCheck size={28} />
            </div>
            <p>
              Tải lên <strong>File 1 (Gốc MyBK)</strong> cùng với{' '}
              <strong>File 2 (Bản đẹp CTĐT)</strong> và/hoặc{' '}
              <strong>File 3 (Bản đẹp KHGD)</strong>,
              <br />
              rồi nhấn <strong>&quot;Bắt đầu rà soát&quot;</strong> để tách rõ lỗi từng
              file, xuất bản sao tô đỏ và lấy Note báo Cô.
            </p>
          </div>
        )}
      </div>
    </>
  );
}
