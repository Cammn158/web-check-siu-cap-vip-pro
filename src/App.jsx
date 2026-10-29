import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  clean,
  sheetToRows,
  runMybkInternalAudit,
  runSingleModeComparison,
  runPrettyCrossCheck,
} from './utils/parser.js';
import { exportAnnotatedPrettyCopy } from './utils/exporter.js';

const STYLES = `
  @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg:#0f1117; --surface:#181c27; --surface2:#1e2335; --border:#2a3050;
    --accent:#4f8ef7; --accent2:#7c5cfc; --green:#22c55e; --red:#ef4444;
    --yellow:#f59e0b; --orange:#f97316; --teal:#14b8a6; --pink:#ec4899;
    --gray:#6b7280; --text:#e2e8f0; --text2:#94a3b8;
    --font:'Be Vietnam Pro',sans-serif; --mono:'JetBrains Mono',monospace;
  }
  body{background:var(--bg);color:var(--text);font-family:var(--font);}
  .app{min-height:100vh;padding:28px 22px 64px;max-width:1320px;margin:0 auto;}

  .header{margin-bottom:24px;}
  .header-tag{display:inline-block;font-size:10px;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:var(--accent);border:1px solid var(--accent);padding:3px 10px;border-radius:4px;margin-bottom:10px;font-family:var(--mono);}
  .header h1{font-size:25px;font-weight:700;background:linear-gradient(135deg,#e2e8f0 30%,#4f8ef7);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;}
  .header p{color:var(--text2);margin-top:6px;font-size:13px;line-height:1.6;}

  /* 3 UPLOAD CARDS */
  .upload-grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:13px;margin-bottom:16px;}
  @media(max-width:900px){.upload-grid-3{grid-template-columns:1fr;}}
  .upload-card{background:var(--surface);border:1.5px dashed var(--border);border-radius:12px;padding:16px;cursor:pointer;transition:all .2s;position:relative;}
  .upload-card:hover,.upload-card.drag-over{border-color:var(--accent);background:rgba(79,142,247,.05);}
  .upload-card.has-file.card-mybk{border-style:solid;border-color:var(--orange);background:rgba(249,115,22,.06);}
  .upload-card.has-file.card-ctdt{border-style:solid;border-color:var(--accent);background:rgba(79,142,247,.07);}
  .upload-card.has-file.card-khgd{border-style:solid;border-color:var(--teal);background:rgba(20,184,166,.07);}
  .upload-card input[type="file"]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;z-index:1;}
  .clear-file-btn{position:absolute;top:10px;right:10px;z-index:3;background:var(--surface2);border:1px solid var(--border);color:var(--text2);border-radius:6px;padding:2px 7px;font-size:10.5px;cursor:pointer;}
  .clear-file-btn:hover{border-color:var(--red);color:var(--red);}
  .upload-icon{font-size:20px;margin-bottom:6px;}
  .upload-label{font-size:10px;font-weight:700;letter-spacing:1.3px;text-transform:uppercase;font-family:var(--mono);margin-bottom:3px;}
  .upload-title{font-size:13px;font-weight:600;margin-bottom:3px;word-break:break-word;}
  .upload-desc{font-size:11px;color:var(--text2);line-height:1.45;}
  .upload-filename{margin-top:8px;font-size:11px;font-family:var(--mono);padding:3px 8px;border-radius:5px;display:inline-block;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}

  .run-btn{width:100%;padding:13px;background:linear-gradient(135deg,var(--accent),var(--accent2));border:none;border-radius:10px;color:white;font-size:14px;font-weight:700;font-family:var(--font);cursor:pointer;transition:all .2s;margin-bottom:22px;}
  .run-btn:hover{opacity:.9;transform:translateY(-1px);}
  .run-btn:disabled{opacity:.35;cursor:not-allowed;transform:none;}

  /* BẢNG PHÂN TÁCH LỖI THEO 3 FILE (BIẾT NGAY SAI Ở FILE NÀO) */
  .file-split-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px;}
  @media(max-width:920px){.file-split-grid{grid-template-columns:1fr;}}
  .file-health-card{background:var(--surface);border:1.5px solid var(--border);border-radius:12px;padding:14px 16px;display:flex;flex-direction:column;justify-content:space-between;transition:all .18s;}
  .file-health-card.active-scope{box-shadow:0 0 0 2px var(--accent),0 6px 20px rgba(79,142,247,.14);}
  .fh-top{display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:8px;}
  .fh-tag{font-size:10px;font-weight:700;font-family:var(--mono);text-transform:uppercase;letter-spacing:1px;padding:2px 7px;border-radius:4px;}
  .fh-title{font-size:13.5px;font-weight:700;margin-bottom:4px;}
  .fh-breakdown{font-size:11.5px;color:var(--text2);line-height:1.55;margin-bottom:12px;}
  .fh-actions{display:flex;gap:7px;flex-wrap:wrap;}
  .fh-btn{padding:6px 11px;border-radius:7px;border:1px solid var(--border);background:var(--surface2);color:var(--text);font-size:11px;font-weight:600;font-family:var(--font);cursor:pointer;transition:all .15s;}
  .fh-btn:hover{border-color:var(--accent);color:var(--accent);}
  .fh-btn.btn-red{background:rgba(239,68,68,.12);border-color:rgba(239,68,68,.35);color:#fca5a5;}
  .fh-btn.btn-red:hover{background:rgba(239,68,68,.22);color:#fff;}
  .fh-btn.btn-orange{background:rgba(249,115,22,.13);border-color:rgba(249,115,22,.38);color:#fdba74;}
  .fh-btn.btn-orange:hover{background:rgba(249,115,22,.24);color:#fff;}

  /* THANH CHỌN FILE ĐẸP ĐANG XEM TRONG B1-B3 */
  .scope-bar{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:10px 14px;margin-bottom:16px;}
  .scope-btns{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
  .scope-btn{padding:7px 14px;border-radius:8px;border:1.5px solid var(--border);background:var(--surface2);color:var(--text2);font-size:12px;font-weight:600;font-family:var(--font);cursor:pointer;display:inline-flex;align-items:center;gap:7px;transition:all .15s;}
  .scope-btn:hover{color:var(--text);}
  .scope-btn.active-ctdt{border-color:var(--accent);background:rgba(79,142,247,.15);color:#93c5fd;}
  .scope-btn.active-khgd{border-color:var(--teal);background:rgba(20,184,166,.15);color:#5eead4;}

  /* STEPPER PIPELINE BAR */
  .stepper-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:9px;flex-wrap:wrap;gap:8px;}
  .stepper-hint{font-size:11.5px;color:var(--text2);}
  .reset-order-btn{background:none;border:1px solid var(--border);color:var(--text2);font-size:11px;padding:3px 9px;border-radius:6px;cursor:pointer;font-family:var(--font);}
  .reset-order-btn:hover{border-color:var(--accent);color:var(--accent);}

  .stepper-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:18px;}
  @media(max-width:920px){.stepper-grid{grid-template-columns:repeat(2,1fr);}}
  @media(max-width:520px){.stepper-grid{grid-template-columns:1fr;}}
  .step-card{background:var(--surface);border:1.5px solid var(--border);border-radius:12px;padding:13px 14px;cursor:pointer;transition:all .18s;display:flex;flex-direction:column;justify-content:space-between;}
  .step-card:hover{border-color:rgba(79,142,247,.55);transform:translateY(-1px);}
  .step-card.active{border-color:var(--accent);background:linear-gradient(180deg,rgba(79,142,247,.13),rgba(124,92,252,.06));}
  .step-card.priority-step.active{border-color:var(--pink);background:linear-gradient(180deg,rgba(236,72,153,.14),rgba(124,92,252,.06));}
  .step-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:7px;gap:6px;}
  .step-num-tag{font-family:var(--mono);font-size:10.5px;font-weight:700;padding:2px 7px;border-radius:5px;background:var(--surface2);color:var(--accent);border:1px solid var(--border);}
  .step-card.priority-step .step-num-tag{color:#f472b6;border-color:rgba(236,72,153,.35);}
  .step-reorder{display:flex;gap:3px;}
  .reorder-btn{background:var(--surface2);border:1px solid var(--border);color:var(--text2);font-size:10px;width:20px;height:20px;border-radius:4px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;padding:0;}
  .reorder-btn:hover:not(:disabled){border-color:var(--accent);color:var(--text);}
  .reorder-btn:disabled{opacity:.25;cursor:default;}
  .step-title{font-size:13px;font-weight:700;line-height:1.35;margin-bottom:4px;color:var(--text);}
  .step-desc{font-size:11px;color:var(--text2);line-height:1.45;margin-bottom:9px;}
  .step-footer{display:flex;justify-content:space-between;align-items:center;}

  .subtabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--border);}
  .subtab-btn{padding:7px 13px;border-radius:8px;border:1px solid var(--border);background:var(--surface);color:var(--text2);font-size:12px;font-weight:500;font-family:var(--font);cursor:pointer;display:inline-flex;align-items:center;gap:6px;transition:all .15s;}
  .subtab-btn:hover{border-color:var(--accent);color:var(--text);}
  .subtab-btn.active{background:rgba(79,142,247,.14);border-color:var(--accent);color:var(--accent);font-weight:600;}

  .tb{font-size:10px;font-weight:700;padding:1.5px 7px;border-radius:20px;font-family:var(--mono);}
  .tb-red{background:var(--red);color:white;}
  .tb-ok{background:rgba(34,197,94,.18);color:#4ade80;border:1px solid rgba(34,197,94,.3);}
  .tb-warn{background:var(--yellow);color:#111;}
  .tb-pink{background:var(--pink);color:white;}
  .tb-blue{background:rgba(79,142,247,.2);color:#93c5fd;}
  .tb-orange{background:rgba(249,115,22,.22);color:#fdba74;border:1px solid rgba(249,115,22,.35);}

  .file-origin-tag{display:inline-flex;align-items:center;gap:4px;font-size:10px;font-weight:700;padding:2px 7px;border-radius:5px;white-space:nowrap;font-family:var(--mono);}
  .fo-ctdt{background:rgba(79,142,247,.15);color:#93c5fd;border:1px solid rgba(79,142,247,.35);}
  .fo-khgd{background:rgba(20,184,166,.15);color:#5eead4;border:1px solid rgba(20,184,166,.35);}
  .fo-mybk{background:rgba(249,115,22,.16);color:#fdba74;border:1px solid rgba(249,115,22,.38);}
  .fo-err{background:rgba(239,68,68,.15);color:#fca5a5;border:1px solid rgba(239,68,68,.38);}

  .filter-row{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px;}
  .filter-tab{padding:4px 11px;border-radius:20px;border:1px solid var(--border);background:transparent;color:var(--text2);font-size:11px;font-family:var(--font);cursor:pointer;transition:all .15s;display:flex;align-items:center;gap:5px;}
  .filter-tab:hover{border-color:var(--accent);color:var(--accent);}
  .filter-tab.active{background:var(--accent);border-color:var(--accent);color:white;font-weight:600;}
  .fdot{width:6px;height:6px;border-radius:50%;display:inline-block;flex-shrink:0;}

  .table-wrap{background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:16px;}
  .table-scroll{overflow-x:auto;max-height:560px;overflow-y:auto;}
  table{width:100%;border-collapse:collapse;font-size:12.5px;}
  thead{background:var(--surface2);position:sticky;top:0;z-index:2;}
  th{padding:10px 12px;text-align:left;font-size:10px;font-weight:600;letter-spacing:.8px;text-transform:uppercase;color:var(--text2);border-bottom:1px solid var(--border);white-space:nowrap;}
  td{padding:8px 12px;border-bottom:1px solid rgba(42,48,80,.4);vertical-align:middle;}
  tr:last-child td{border-bottom:none;}
  tr:hover td{background:rgba(255,255,255,.018);}
  .row-vi{font-weight:500;color:var(--text);}
  .row-en{font-size:11px;color:var(--text2);font-style:italic;}
  .row-group td{background:rgba(79,142,247,.055);}
  .code-cell{font-family:var(--mono);font-size:11px;color:var(--accent);white-space:nowrap;font-weight:600;}
  .cred-cell{text-align:center;font-family:var(--mono);font-weight:600;}

  .badge{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:20px;font-size:10.5px;font-weight:600;white-space:nowrap;}
  .b-match{background:rgba(34,197,94,.1);color:#22c55e;border:1px solid rgba(34,197,94,.22);}
  .b-miss{background:rgba(239,68,68,.09);color:#ef4444;border:1px solid rgba(239,68,68,.2);}
  .b-wrong{background:rgba(245,158,11,.09);color:#f59e0b;border:1px solid rgba(245,158,11,.2);}
  .b-wrong-en{background:rgba(79,142,247,.09);color:#4f8ef7;border:1px solid rgba(79,142,247,.2);}
  .b-wrong-prereq{background:rgba(236,72,153,.12);color:#f472b6;border:1px solid rgba(236,72,153,.3);}
  .b-na{background:rgba(239,68,68,.07);color:#f87171;border:1px solid rgba(239,68,68,.15);}
  .b-skip{background:rgba(107,114,128,.09);color:#9ca3af;border:1px solid rgba(107,114,128,.22);}
  .b-bb{background:rgba(124,92,252,.14);color:#c4b5fd;border:1px solid rgba(124,92,252,.3);font-family:var(--mono);}
  .b-tc{background:rgba(20,184,166,.12);color:#5eead4;border:1px solid rgba(20,184,166,.28);font-family:var(--mono);}

  .prereq-list{display:flex;flex-wrap:wrap;gap:4px;align-items:center;}
  .prereq-pill{display:inline-flex;align-items:center;gap:3px;font-family:var(--mono);font-size:10.5px;padding:1.5px 6px;border-radius:4px;border:1px solid var(--border);background:var(--surface2);white-space:nowrap;}
  .prereq-pill.t-TQ{background:rgba(236,72,153,.13);color:#f472b6;border-color:rgba(236,72,153,.35);font-weight:600;}
  .prereq-pill.t-HT{background:rgba(245,158,11,.12);color:#fbbf24;border-color:rgba(245,158,11,.3);}
  .prereq-pill.t-SH,.prereq-pill.t-SHT{background:rgba(20,184,166,.12);color:#2dd4bf;border-color:rgba(20,184,166,.3);}
  .prereq-pill.t-KN{background:rgba(79,142,247,.1);color:#93c5fd;border:color:rgba(79,142,247,.25);}

  .grp-table-wrap{background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:18px;}
  .grp-header{padding:12px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;}
  .grp-header-title{font-size:13px;font-weight:600;}
  .grp-header-sub{font-size:11px;color:var(--text2);}
  .grp-grid{display:grid;grid-template-columns:2.4fr 0.8fr 0.9fr 0.9fr 1fr 140px;gap:0;}
  @media(max-width:760px){.grp-grid{grid-template-columns:1.8fr 0.7fr 0.8fr 0.8fr 0.8fr 110px;}}
  .grp-grid>div{padding:9px 12px;font-size:12px;display:flex;align-items:center;}
  .grp-grid.grp-th>div{font-size:10px;font-weight:600;letter-spacing:.8px;text-transform:uppercase;color:var(--text2);background:var(--surface2);border-bottom:1px solid var(--border);}
  .grp-row-data{border-bottom:1px solid rgba(42,48,80,.4);}
  .grp-row-data:last-child{border-bottom:none;}
  .grp-id-badge{font-family:var(--mono);font-size:11px;font-weight:700;padding:2px 7px;border-radius:5px;margin-right:8px;}
  .num-ok{color:var(--green);font-weight:700;font-family:var(--mono);}
  .num-bad{color:var(--red);font-weight:700;font-family:var(--mono);}
  .num-neutral{color:var(--text);font-family:var(--mono);}

  .section-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:16px;}
  .section-head{padding:12px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:8px;background:var(--surface2);flex-wrap:wrap;}
  .section-head-left{display:flex;align-items:center;gap:8px;}
  .section-head-title{font-size:13px;font-weight:600;}
  .section-head-count{font-size:11px;color:var(--text2);}
  .section-body{padding:14px 16px;}

  .alert-banner{border-radius:10px;padding:12px 15px;margin-bottom:14px;font-size:12.5px;line-height:1.6;display:flex;gap:10px;align-items:flex-start;}
  .alert-warn{background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.28);color:#fcd34d;}
  .alert-ok{background:rgba(34,197,94,.08);border:1px solid rgba(34,197,94,.25);color:#86efac;}
  .alert-pink{background:rgba(236,72,153,.09);border:1px solid rgba(236,72,153,.3);color:#fbcfe8;}

  .toolbar{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:10px;}
  .export-btn{padding:5px 11px;background:var(--surface2);border:1px solid var(--border);border-radius:7px;color:var(--text2);font-size:11px;font-family:var(--font);cursor:pointer;transition:all .15s;white-space:nowrap;}
  .export-btn:hover{border-color:var(--accent);color:var(--accent);}
  .red-action-btn{padding:11px 20px;background:linear-gradient(135deg,#ef4444,#dc2626);border:none;border-radius:9px;color:#fff;font-size:13px;font-weight:700;font-family:var(--font);cursor:pointer;transition:all .18s;display:inline-flex;align-items:center;gap:7px;}
  .red-action-btn:hover{opacity:.92;transform:translateY(-1px);}
  .red-action-btn:disabled{opacity:.4;cursor:not-allowed;transform:none;}

  .note-input{width:100%;min-width:165px;background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:5px 8px;color:var(--text);font-size:11.5px;font-family:var(--font);}
  .note-input:focus{outline:none;border-color:var(--red);}
  .note-input.has-val{border-color:rgba(239,68,68,.55);color:#fca5a5;background:rgba(239,68,68,.06);}
  .b4-config-grid{display:grid;grid-template-columns:1.35fr 1fr;gap:14px;margin-bottom:16px;}
  @media(max-width:820px){.b4-config-grid{grid-template-columns:1fr;}}

  /* KHUNG NOTE FILE MYBK ĐỂ BÁO RIÊNG VỚI CÔ */
  .professor-panel{background:linear-gradient(180deg,rgba(249,115,22,.07),rgba(24,28,39,.98));border:1.5px solid rgba(249,115,22,.38);border-radius:14px;overflow:hidden;margin-top:26px;}
  .prof-head{padding:14px 18px;background:rgba(249,115,22,.12);border-bottom:1px solid rgba(249,115,22,.28);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;}
  .prof-body{padding:16px 18px;display:grid;grid-template-columns:1.1fr 1fr;gap:18px;}
  @media(max-width:860px){.prof-body{grid-template-columns:1fr;}}
  .prof-item{display:flex;gap:8px;align-items:flex-start;padding:8px 10px;border-radius:8px;background:rgba(15,17,23,.55);border:1px solid var(--border);margin-bottom:7px;font-size:12px;line-height:1.5;}
  .prof-textarea{width:100%;min-height:220px;background:#0b0d13;border:1px solid rgba(249,115,22,.4);border-radius:9px;padding:12px;color:#fde68a;font-size:12.5px;line-height:1.65;font-family:var(--font);resize:vertical;}
  .prof-textarea:focus{outline:none;border-color:var(--orange);}

  .option-btn{padding:6px 13px;border-radius:7px;border:1.5px solid var(--border);background:var(--surface);color:var(--text2);font-size:12px;font-family:var(--font);cursor:pointer;transition:all .15s;}
  .option-btn:hover{border-color:var(--accent);color:var(--accent);}
  .option-btn.active{border-color:var(--accent);background:rgba(79,142,247,.12);color:var(--accent);font-weight:600;}

  .step-nav-footer{display:flex;justify-content:space-between;align-items:center;margin-top:20px;padding-top:14px;border-top:1px solid var(--border);}
  .error-box{background:rgba(239,68,68,.07);border:1px solid rgba(239,68,68,.22);border-radius:10px;padding:12px 14px;color:#f87171;font-size:12.5px;margin-bottom:18px;line-height:1.6;white-space:pre-line;}
  .empty-state{text-align:center;padding:38px 20px;color:var(--text2);}
  .empty-state .icon{font-size:36px;margin-bottom:10px;}
  .empty-state p{font-size:13px;line-height:1.7;}
  .diff-hint{font-size:10.5px;color:var(--orange);margin-top:3px;}
  .diff-hint-blue{font-size:10.5px;color:var(--accent);margin-top:3px;}
  .diff-hint-pink{font-size:10.5px;color:var(--pink);margin-top:3px;}
`;

const DEFAULT_STEP_ORDER = ['b1', 'b2', 'b3', 'b4'];

const STEP_META = {
  b1: {
    id: 'b1',
    code: 'B1',
    title: 'Phân nhóm & Bắt buộc / Tự chọn',
    desc: 'Kiểm tra môn BB / TC thuộc đúng nhóm & học kỳ trên từng file',
  },
  b2: {
    id: 'b2',
    code: 'B2',
    title: 'Sót môn, Tín chỉ, Nội dung & Chính tả',
    desc: 'Tách rõ sót/sai ở CTĐT đẹp, KHGD đẹp hay lỗi gốc từ MyBK',
  },
  b3: {
    id: 'b3',
    code: 'B3',
    title: 'Tiên quyết (Quan trọng), Bắt buộc & ĐK Tốt nghiệp',
    desc: 'Rà tiên quyết, môn BB, ghi chú & Đ/kiện ngoại ngữ tốt nghiệp SĐH',
    priority: true,
  },
  b4: {
    id: 'b4',
    code: 'B4',
    title: 'Xuất Bản Sao Tô Đỏ (CTĐT / KHGD)',
    desc: 'Tạo bản sao File đẹp với chữ đỏ / highlight ô đỏ để tự check & sửa',
  },
};

const GROUP_COLORS_CTDT = {
  A: '#4f8ef7',
  'A.1': '#4f8ef7',
  'A.2': '#7c5cfc',
  'A.3': '#14b8a6',
  B: '#22c55e',
  C: '#f59e0b',
  'C.1': '#f59e0b',
  'C.2': '#f97316',
  D: '#ec4899',
  E: '#6b7280',
};

const GROUP_COLORS_KHGD = {
  HK1: '#22c55e',
  HK2: '#4f8ef7',
  HK3: '#f59e0b',
  HK4: '#f97316',
  HK5: '#ec4899',
};

function getGroupColor(gid, mode) {
  return (mode === 'khgd' ? GROUP_COLORS_KHGD : GROUP_COLORS_CTDT)[gid] || '#94a3b8';
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
    } else if (slot === 'ctdt') {
      setFileCtdtPretty(f);
      setOutNameByFile((prev) => ({
        ...prev,
        ctdt: `${base}_check_note_do.xlsx`,
      }));
    } else if (slot === 'khgd') {
      setFileKhgdPretty(f);
      setOutNameByFile((prev) => ({
        ...prev,
        khgd: `${base}_check_note_do.xlsx`,
      }));
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
    } catch (err) {
      setError(err.message);
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
      setTimeout(() => setCopiedProf(false), 2500);
    } catch {
      // Fallback
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
    } catch (err) {
      setError(`Lỗi khi xuất bản sao ${targetRes.fileLabel}: ${err.message}`);
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
  };

  const filteredRows = currentRes?.rowResults.filter((r) => {
    if (r.type === 'group' || r.type === 'separator') return rowFilter === 'all';
    if (rowFilter === 'all') return true;
    if (rowFilter === 'wrong-group') return r.groupMismatch;
    if (rowFilter === 'wrong-prereq') return r.prereqMismatch;
    return r.status === rowFilter;
  });

  const filteredSpelling = currentRes?.spellingIssues.filter((sp) => {
    if (spellFilter === 'all') return true;
    if (spellFilter === 'pretty_only') return !sp.alsoInMybk;
    if (spellFilter === 'from_mybk') return sp.alsoInMybk;
    return sp.severity === spellFilter;
  });

  const filteredPrereqCourses = currentRes?.allPrereqCourses.filter((r) => {
    if (prereqFilter === 'mismatch') return r.prereqMismatch;
    if (prereqFilter === 'tq-only') {
      const hasOrigTQ = r.prereqItems?.some((i) => i.type === 'TQ');
      const hasPrettyTQ = r.pEntry?.prereqItems?.some((i) => i.type === 'TQ');
      return hasOrigTQ || hasPrettyTQ;
    }
    return true;
  });

  const filteredB4Rows = currentRes?.prettyCoursesList.filter((item) => {
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
        <span className="tb tb-warn">{stats.b1Issues} lệch</span>
      ) : (
        <span className="tb tb-ok">✓ Đúng nhóm</span>
      );
    }
    if (stepId === 'b2') {
      return stats.b2Issues > 0 ? (
        <span className="tb tb-red">{stats.b2Issues} lỗi</span>
      ) : (
        <span className="tb tb-ok">✓ Khớp</span>
      );
    }
    if (stepId === 'b3') {
      return stats.b3Issues > 0 ? (
        <span className="tb tb-pink">⚡ {stats.b3Issues} cần xử lý</span>
      ) : (
        <span className="tb tb-ok">✓ Chuẩn ĐK</span>
      );
    }
    if (stepId === 'b4') {
      return (
        <span className="tb tb-red">
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
      <style>{STYLES}</style>
      <div className="app">
        <div className="header">
          <div className="header-tag">WEB CHECK SIU CẤP VIP PRO • V3.5</div>
          <h1>web check siu cấp vip pro</h1>
          <p>
            Hỗ trợ tải cùng lúc <strong>3 file</strong> (1 File gốc MyBK + Bản đẹp CTĐT +
            Bản đẹp KHGD). Tự động <strong>tách riêng lỗi nằm ở file nào</strong>: File bản
            đẹp bị thiếu/sai sẽ xuất bản sao với <strong>chữ đỏ / highlight ô đỏ</strong>{' '}
            ở B4; lỗi gốc từ hệ thống MyBK được tổng hợp vào{' '}
            <strong>Khung Note riêng để báo Cô</strong>.
          </p>
        </div>

        {/* =====================================================================
            KHỐI TẢI LÊN 3 FILE (1 MYBK + 2 BẢN ĐẸP: CTĐT & KHGD)
        ===================================================================== */}
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
                }}
              >
                ✕ Gỡ
              </button>
            )}
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={onPickFile('mybk')}
            />
            <div className="upload-icon">🏛️</div>
            <div className="upload-label" style={{ color: 'var(--orange)' }}>
              File 1 — Gốc MyBK (Bắt buộc)
            </div>
            <div className="upload-title">
              {fileMybk ? fileMybk.name : 'Chọn File tải từ MyBK'}
            </div>
            <div className="upload-desc">
              Gồm 2 sheet gốc: Khối kiến thức & Kế hoạch giảng dạy. Các bất cập từ file
              này sẽ gom vào Khung Note báo Cô.
            </div>
            {fileMybk && (
              <div
                className="upload-filename"
                style={{
                  color: '#fdba74',
                  background: 'rgba(249,115,22,.15)',
                }}
              >
                ✓ {fileMybk.name}
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
                }}
              >
                ✕ Gỡ
              </button>
            )}
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={onPickFile('ctdt')}
            />
            <div className="upload-icon">📘</div>
            <div className="upload-label" style={{ color: 'var(--accent)' }}>
              File 2 — Bản đẹp CTĐT (Khối kiến thức)
            </div>
            <div className="upload-title">
              {fileCtdtPretty ? fileCtdtPretty.name : 'Chọn File CTDT_... bản đẹp'}
            </div>
            <div className="upload-desc">
              Dò với sheet Khối kiến thức MyBK. Nếu thiếu/sai sẽ xuất bản sao tô đỏ ô lỗi
              ở B4.
            </div>
            {fileCtdtPretty && (
              <div
                className="upload-filename"
                style={{
                  color: '#93c5fd',
                  background: 'rgba(79,142,247,.16)',
                }}
              >
                ✓ {fileCtdtPretty.name}
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
                }}
              >
                ✕ Gỡ
              </button>
            )}
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={onPickFile('khgd')}
            />
            <div className="upload-icon">📗</div>
            <div className="upload-label" style={{ color: 'var(--teal)' }}>
              File 3 — Bản đẹp KHGD (Kế hoạch giảng dạy)
            </div>
            <div className="upload-title">
              {fileKhgdPretty ? fileKhgdPretty.name : 'Chọn File KHGD_... bản đẹp'}
            </div>
            <div className="upload-desc">
              Dò với sheet KHGD MyBK & đối chiếu chéo với CTĐT đẹp. Xuất bản sao KHGD tô
              đỏ ô lỗi ở B4.
            </div>
            {fileKhgdPretty && (
              <div
                className="upload-filename"
                style={{
                  color: '#5eead4',
                  background: 'rgba(20,184,166,.16)',
                }}
              >
                ✓ {fileKhgdPretty.name}
              </div>
            )}
          </div>
        </div>

        <button
          className="run-btn"
          onClick={handleRun}
          disabled={!fileMybk || (!fileCtdtPretty && !fileKhgdPretty) || loading}
        >
          {loading
            ? '⏳ Đang phân tích & tách lỗi theo từng file...'
            : fileCtdtPretty && fileKhgdPretty
            ? '🔍 Bắt đầu rà soát cả 3 File (MyBK + Bản đẹp CTĐT + Bản đẹp KHGD)'
            : '🔍 Bắt đầu rà soát các File đã chọn'}
        </button>

        {error && <div className="error-box">⚠️ {error}</div>}

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
                style={{
                  borderColor: reviewData.ctdtResult
                    ? 'rgba(79,142,247,.45)'
                    : 'var(--border)',
                }}
              >
                <div>
                  <div className="fh-top">
                    <span className="fh-tag fo-ctdt">📘 File 2 • Bản đẹp CTĐT</span>
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
                  <div className="fh-title">
                    {fileCtdtPretty ? fileCtdtPretty.name : 'Chưa chọn Bản đẹp CTĐT'}
                  </div>
                  {reviewData.ctdtResult ? (
                    <div className="fh-breakdown">
                      • Sót/thừa môn:{' '}
                      <b>
                        {reviewData.ctdtResult.stats.miss +
                          reviewData.ctdtResult.stats.extra}
                      </b>{' '}
                      &nbsp;|&nbsp; Lệch nhóm/BB-TC:{' '}
                      <b>
                        {reviewData.ctdtResult.stats.groupMismatch +
                          reviewData.ctdtResult.stats.categoryMismatch}
                      </b>
                      <br />• Sai Tên/Tín chỉ:{' '}
                      <b>
                        {reviewData.ctdtResult.stats.wrong +
                          reviewData.ctdtResult.stats.creditMismatch}
                      </b>{' '}
                      &nbsp;|&nbsp; Chính tả:{' '}
                      <b>{reviewData.ctdtResult.stats.typoCount}</b>
                      <br />• Thiếu/sai Tiên quyết:{' '}
                      <b style={{ color: '#f472b6' }}>
                        {reviewData.ctdtResult.stats.wrongPrereq}
                      </b>
                    </div>
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
                      👁️ Xem chi tiết CTĐT đẹp
                    </button>
                    <button
                      className="fh-btn btn-red"
                      onClick={() => {
                        setActiveFileScope('ctdt');
                        setActiveStep('b4');
                      }}
                    >
                      🖍️ Tạo bản sao CTĐT tô đỏ
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
                style={{
                  borderColor: reviewData.khgdResult
                    ? 'rgba(20,184,166,.45)'
                    : 'var(--border)',
                }}
              >
                <div>
                  <div className="fh-top">
                    <span className="fh-tag fo-khgd">📗 File 3 • Bản đẹp KHGD</span>
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
                  <div className="fh-title">
                    {fileKhgdPretty ? fileKhgdPretty.name : 'Chưa chọn Bản đẹp KHGD'}
                  </div>
                  {reviewData.khgdResult ? (
                    <div className="fh-breakdown">
                      • Sót/thừa môn:{' '}
                      <b>
                        {reviewData.khgdResult.stats.miss +
                          reviewData.khgdResult.stats.extra}
                      </b>{' '}
                      &nbsp;|&nbsp; Sai HK/BB-TC:{' '}
                      <b>
                        {reviewData.khgdResult.stats.groupMismatch +
                          reviewData.khgdResult.stats.categoryMismatch}
                      </b>
                      <br />• Sai Tên/Tín chỉ:{' '}
                      <b>
                        {reviewData.khgdResult.stats.wrong +
                          reviewData.khgdResult.stats.creditMismatch}
                      </b>{' '}
                      &nbsp;|&nbsp; Chính tả:{' '}
                      <b>{reviewData.khgdResult.stats.typoCount}</b>
                      <br />• Thiếu/sai Tiên quyết:{' '}
                      <b style={{ color: '#f472b6' }}>
                        {reviewData.khgdResult.stats.wrongPrereq}
                      </b>
                    </div>
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
                      👁️ Xem chi tiết KHGD đẹp
                    </button>
                    <button
                      className="fh-btn btn-red"
                      onClick={() => {
                        setActiveFileScope('khgd');
                        setActiveStep('b4');
                      }}
                    >
                      🖍️ Tạo bản sao KHGD tô đỏ
                    </button>
                  </div>
                )}
              </div>

              {/* THẺ 3: LỖI GỐC TỪ FILE MYBK -> BÁO RIÊNG VỚI CÔ */}
              <div
                className="file-health-card"
                style={{ borderColor: 'rgba(249,115,22,.45)' }}
              >
                <div>
                  <div className="fh-top">
                    <span className="fh-tag fo-mybk">🏛️ File 1 • Gốc MyBK (Báo Cô)</span>
                    <span className="tb tb-orange">
                      {reviewData.mybkAudit.totalIssues + mybkCustomItems.length} mục báo
                      Cô
                    </span>
                  </div>
                  <div className="fh-title">{fileMybk?.name}</div>
                  <div className="fh-breakdown">
                    • Lệch mục BB/TC giữa 2 sheet MyBK:{' '}
                    <b>{reviewData.mybkAudit.khgdMisplacedSubgroups.length} nhóm</b> (
                    {reviewData.mybkAudit.bbInTcIssues.length} môn)
                    <br />• Chính tả gốc từ hệ thống MyBK:{' '}
                    <b>{reviewData.mybkAudit.mybkTypos.length} lỗi</b>
                    <br />• Khác biệt danh mục/ĐK 2 sheet MyBK:{' '}
                    <b>
                      {reviewData.mybkAudit.missingInKhgd.length +
                        reviewData.mybkAudit.prereqCrossIssues.length}
                    </b>
                  </div>
                </div>
                <div className="fh-actions">
                  <button className="fh-btn btn-orange" onClick={scrollToProfPanel}>
                    📨 Mở Khung Note Báo Riêng Với Cô ↓
                  </button>
                </div>
              </div>
            </div>

            {/* CẢNH BÁO NẾU 2 FILE BẢN ĐẸP (CTĐT ĐẸP & KHGD ĐẸP) LỆCH NHAU */}
            {reviewData.prettyCrossDiffs.length > 0 && (
              <div className="alert-banner alert-warn">
                <span>⚡</span>
                <div style={{ width: '100%' }}>
                  <strong>
                    Phát hiện {reviewData.prettyCrossDiffs.length} điểm chưa đồng bộ giữa
                    2 File Bản Đẹp (Bản đẹp CTĐT ↔ Bản đẹp KHGD):
                  </strong>
                  <div style={{ marginTop: 6, fontSize: 12 }}>
                    {reviewData.prettyCrossDiffs.slice(0, 6).map((d, i) => (
                      <div key={i} style={{ marginBottom: 3 }}>
                        • <code>{d.code}</code> ({d.name}) —{' '}
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
                <span style={{ fontSize: 12, color: 'var(--text2)', fontWeight: 600 }}>
                  📂 Đang xem kết quả rà soát của:
                </span>
                {reviewData.ctdtResult && (
                  <button
                    className={`scope-btn ${
                      currentScopeKey === 'ctdt' ? 'active-ctdt' : ''
                    }`}
                    onClick={() => setActiveFileScope('ctdt')}
                  >
                    📘 Bản đẹp CTĐT ({fileCtdtPretty?.name})
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
                    📗 Bản đẹp KHGD ({fileKhgdPretty?.name})
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
                📨 Khung Note File MyBK báo Cô (
                {Object.values(mybkSelectedIds).filter(Boolean).length})
              </button>
            </div>
          </>
        )}

        {/* THANH QUY TRÌNH 4 BƯỚC (CÓ THỂ ĐỔI THỨ TỰ) */}
        <div className="stepper-header">
          <div className="stepper-hint">
            🧭 <strong>Quy trình 4 bước rà soát:</strong> Có thể nhấn <code>◀</code>{' '}
            <code>▶</code> để đổi thứ tự các bước theo thói quen làm việc của bạn.
          </div>
          {stepOrder.join(',') !== DEFAULT_STEP_ORDER.join(',') && (
            <button
              className="reset-order-btn"
              onClick={() => setStepOrder(DEFAULT_STEP_ORDER)}
            >
              ↺ Đặt lại mặc định (B1 → B4)
            </button>
          )}
        </div>

        <div className="stepper-grid">
          {stepOrder.map((stepId, idx) => {
            const meta = STEP_META[stepId];
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
                  <span style={{ fontSize: 10.5, color: 'var(--text2)' }}>
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
          <>
            {/* =================================================================
                BƯỚC B1: PHÂN NHÓM & BẮT BUỘC / TỰ CHỌN
            ================================================================= */}
            {activeStep === 'b1' && (
              <div>
                {/* 1. Lỗi phân nhóm / học kỳ / BB-TC ngay trên File Bản đẹp đang xem */}
                <div className="section-card">
                  <div className="section-head">
                    <div className="section-head-left">
                      <span>🗂️</span>
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
                    <div className="empty-state" style={{ padding: '26px 16px' }}>
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
                          {currentRes.groupMismatches.map((m, idx) => (
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
                          {currentRes.categoryMismatches.map((m, idx) => (
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
                          {currentRes.khgdPrettyBbInTc.map((it, idx) => (
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
                              <td style={{ fontSize: 11.5, color: 'var(--yellow)' }}>
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
                        <span>🏛️</span>
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
                        📨 Đã tự động đưa vào Khung Note Báo Cô ↓
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
                          {reviewData.mybkAudit.bbInTcIssues.map((it, idx) => (
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
                      <span>📑</span>
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
                            <td style={{ fontFamily: 'var(--mono)' }}>{g.origCount}</td>
                            <td style={{ fontFamily: 'var(--mono)' }}>{g.prettyCount}</td>
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
                    📋 Đối chiếu từng dòng ({currentRes.fileLabel})
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
                    🔎 Sót môn / Thừa môn ({currentRes.fileLabel})
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
                    🔢 Tín chỉ ({currentRes.fileLabel})
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
                    ✍️ Chính tả (Phân biệt lỗi Bản đẹp vs lỗi từ MyBK)
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
                          { key: 'match', label: 'Khớp', dot: '#22c55e' },
                          { key: 'miss', label: 'Thiếu môn', dot: '#ef4444' },
                          { key: 'wrong', label: 'Sai thông tin', dot: '#f59e0b' },
                          {
                            key: 'wrong-prereq',
                            label: 'Sai ĐK tiên quyết',
                            dot: '#ec4899',
                          },
                          { key: 'wrong-en', label: 'Sai tiếng Anh', dot: '#4f8ef7' },
                          { key: 'wrong-group', label: 'Sai nhóm/HK', dot: '#f97316' },
                          { key: 'na', label: '#N/A', dot: '#f87171' },
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
                        ⬇ Xuất CSV ({currentRes.fileLabel})
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
                                    padding: 26,
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
                                        fontWeight: 600,
                                      }}
                                    >
                                      {row.name}
                                    </td>
                                  </tr>
                                );
                              }

                              if (row.type === 'separator') {
                                return (
                                  <tr key={idx} style={{ opacity: 0.65 }}>
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
                                      fontSize: 10,
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
                                            <div style={{ fontWeight: 600, marginTop: 2 }}>
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
                                          gap: 4,
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
                        <span>🔎</span>
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
                      <div className="empty-state" style={{ padding: '28px 16px' }}>
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
                            {currentRes.missingInPretty.map((m, i) => (
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
                                <td style={{ color: '#fca5a5', fontSize: 11.5 }}>
                                  Sẽ chèn dòng cảnh báo ĐỎ cuối bản sao{' '}
                                  {currentRes.fileLabel}
                                </td>
                              </tr>
                            ))}
                            {currentRes.extraInPretty.map((m, i) => (
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
                                <td style={{ color: '#fcd34d', fontSize: 11.5 }}>
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
                        <span>🎓</span>
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
                          🔢 Đối chiếu tín chỉ từng{' '}
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
                            style={isBad ? { background: 'rgba(239,68,68,.05)' } : {}}
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
                                <div style={{ fontSize: 11.5 }}>{g.label}</div>
                                {g.note && (
                                  <div
                                    style={{
                                      fontSize: 10,
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
                        <span>✍️</span>
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
                            dot: '#ef4444',
                          },
                          {
                            key: 'from_mybk',
                            label: `Lỗi có từ Gốc MyBK (${
                              currentRes.spellingIssues.filter((x) => x.alsoInMybk)
                                .length
                            })`,
                            dot: '#f97316',
                          },
                          {
                            key: 'typo',
                            label: `Chỉ lỗi từ vựng (${stats.typoCount})`,
                            dot: '#ec4899',
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
                      <div className="empty-state">
                        <div className="icon">✨</div>
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
                                  <div style={{ color: '#fca5a5', fontSize: 12 }}>
                                    {sp.detail}
                                  </div>
                                  {sp.suggestion && (
                                    <div
                                      style={{
                                        fontSize: 11,
                                        color: 'var(--green)',
                                        marginTop: 2,
                                      }}
                                    >
                                      ↳ Gợi ý: <code>{sp.suggestion}</code>
                                    </div>
                                  )}
                                </td>
                                <td>
                                  <div style={{ display: 'flex', gap: 5 }}>
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
                    <span>⚡</span>
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
                    <span>📝</span>
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
                      <span>🔗</span>
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
                          dot: '#ef4444',
                        },
                        {
                          key: 'tq-only',
                          label: 'Chỉ môn có Tiên quyết (TQ)',
                          dot: '#ec4899',
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
                    <div className="empty-state">
                      <div className="icon">✅</div>
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
                                          color: '#f87171',
                                          fontSize: 11,
                                          fontWeight: 700,
                                          marginTop: 3,
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
                              <td style={{ fontSize: 11 }}>
                                {(item.prereqAnalysis || []).map((pa, i) => (
                                  <div key={i} style={{ marginBottom: 2 }}>
                                    <code>{pa.code}</code>: {pa.scopeLabel}
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
                      <span>🎓</span>
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
                        {currentRes.specialRequirements.map((sr, idx) => (
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
                                <div style={{ color: '#f87171', fontWeight: 700 }}>
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
                        <span>🖍️</span>
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
                          fontSize: 12,
                          color: 'var(--text2)',
                          lineHeight: 1.6,
                          marginBottom: 12,
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
                        <strong style={{ color: '#f87171' }}>
                          đánh dấu màu đỏ trực tiếp ngay tại ô đó
                        </strong>{' '}
                        kèm cột Note màu đỏ bên phải bảng.
                      </p>

                      {/* Chọn kiểu tô đỏ */}
                      <div style={{ marginBottom: 12 }}>
                        <div
                          style={{
                            fontSize: 11.5,
                            color: 'var(--text2)',
                            marginBottom: 6,
                            fontWeight: 600,
                          }}
                        >
                          Kiểu đánh dấu ô bị sai / thiếu trên bản sao Excel:
                        </div>
                        <div style={{ display: 'flex', gap: 7, flexWrap: 'wrap' }}>
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
                          marginBottom: 12,
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

                      <div style={{ marginBottom: 12 }}>
                        <label
                          style={{
                            display: 'block',
                            fontSize: 11.5,
                            color: 'var(--text2)',
                            marginBottom: 4,
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
                            {exportingKey === 'ctdt'
                              ? '⏳ Đang xuất CTĐT...'
                              : `📥 Tải Bản Sao CTĐT Đẹp Tô Đỏ (${reviewData.ctdtResult.stats.totalFileErrors} lỗi)`}
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
                            {exportingKey === 'khgd'
                              ? '⏳ Đang xuất KHGD...'
                              : `📥 Tải Bản Sao KHGD Đẹp Tô Đỏ (${reviewData.khgdResult.stats.totalFileErrors} lỗi)`}
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  <div className="section-card" style={{ marginBottom: 0 }}>
                    <div className="section-head">
                      <div className="section-head-left">
                        <span>📝</span>
                        <span className="section-head-title">
                          Ghi chú chung đính kèm đầu bản sao ({currentRes.fileLabel})
                        </span>
                      </div>
                    </div>
                    <div className="section-body">
                      <textarea
                        className="note-input"
                        rows={5}
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
                          marginTop: 8,
                          fontSize: 11.5,
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
                      <span>📋</span>
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
                          dot: '#ef4444',
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
                          dot: '#22c55e',
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
                                      fontSize: 11,
                                      color:
                                        item.severity === 'error'
                                          ? '#fca5a5'
                                          : '#fcd34d',
                                      lineHeight: 1.45,
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
                ← Bước trước
              </button>
              <span style={{ fontSize: 12, color: 'var(--text2)' }}>
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
                Bước tiếp theo →
              </button>
            </div>
          </>
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
                    fontSize: 10.5,
                    fontFamily: 'var(--mono)',
                    color: '#fdba74',
                    textTransform: 'uppercase',
                    letterSpacing: 1,
                    fontWeight: 700,
                  }}
                >
                  🏛️ Tách riêng bất cập từ File 1 — Hệ thống MyBK ({fileMybk?.name})
                </div>
                <div style={{ fontSize: 15, fontWeight: 700, marginTop: 2 }}>
                  📨 Khung Tổng Hợp Note File MyBK Để Báo Riêng Với Cô
                </div>
              </div>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <button className="fh-btn btn-orange" onClick={handleCopyProfNote}>
                  {copiedProf ? '✓ Đã Copy vào Clipboard!' : '📋 Copy nội dung báo Cô'}
                </button>
                <button className="fh-btn" onClick={handleDownloadProfTxt}>
                  ⬇ Tải file .txt
                </button>
              </div>
            </div>

            <div className="prof-body">
              {/* Cột trái: Danh sách các điểm phát hiện trên File MyBK + ô thêm nhanh */}
              <div>
                <div
                  style={{
                    fontSize: 12,
                    fontWeight: 600,
                    marginBottom: 8,
                    color: '#fed7aa',
                  }}
                >
                  Tick chọn các điểm trên File gốc MyBK muốn đưa vào tin nhắn báo Cô:
                </div>

                <div style={{ maxHeight: 260, overflowY: 'auto', paddingRight: 4 }}>
                  {[...reviewData.mybkAudit.reportItems, ...mybkCustomItems].length ===
                    0 && (
                    <div
                      style={{
                        fontSize: 12,
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
                          style={{ marginTop: 3 }}
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
                <div style={{ display: 'flex', gap: 6, marginTop: 10 }}>
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
                    + Thêm ý
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
                    marginBottom: 8,
                  }}
                >
                  <span
                    style={{ fontSize: 12, fontWeight: 600, color: '#fed7aa' }}
                  >
                    Nội dung tin nhắn / ghi chú gửi Cô (có thể sửa trực tiếp):
                  </span>
                  {profDraftEdited !== null && (
                    <button
                      className="reset-order-btn"
                      onClick={() => setProfDraftEdited(null)}
                    >
                      ↺ Tạo lại tự động
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
            <div className="icon">📊</div>
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
