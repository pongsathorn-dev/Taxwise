import { useEffect, useMemo, useRef, useState } from 'react';
import { ArrowDownUp, ArrowRight, Check, ChevronDown, Download, FileSpreadsheet, Filter, History, LayoutDashboard, LoaderCircle, Plus, Search, Settings2, ShieldCheck, Trash2, Upload, Wallet, X } from 'lucide-react';
import type * as ExcelJSTypes from 'exceljs';
import { adminEmail, createNameRule, deleteNameRule, firebaseEnabled, signInOwner, signOutOwner, watchAuth, watchNameRules, type NameRule } from './firebase';

type Item = { payer: string; income_type: string; a_count: number; b_count: number; amount_diff: number; tax_diff: number; a_amount: number; b_amount: number; a_tax: number; b_tax: number; status: string };
type Result = { label_a: string; label_b: string; file_a: string; file_b: string; items: Item[]; groups_a: Group[]; groups_b: Group[]; rules: Rule[] };
type Rule = NameRule;
type Group = { payer: string; income: string; amount: number; tax: number; count: number; tins: string[] };
const money = (n: number) => new Intl.NumberFormat('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
const fmt = (n: number) => `${n < 0 ? '−' : ''}฿ ${money(Math.abs(n))}`;
const aliases = {
  payer: ['ชื่อผู้มีหน้าที่หัก', 'ผู้มีหน้าที่หัก', 'ชื่อผู้จ่าย', 'ผู้จ่าย', 'payer', 'payer name', 'withholding agent', 'ชื่อผู้หักภาษี'],
  income: ['ประเภทเงินได้', 'ประเภทเงินได้พึงประเมิน', 'income type', 'type'],
  amount: ['จำนวนเงิน', 'ยอดเงิน', 'มูลค่า', 'amount', 'income amount', 'ยอดเงินได้', 'มูลค่ารวม'],
  tax: ['ภาษีหัก ณ ที่จ่าย', 'จำนวนภาษี', 'ภาษี', 'tax', 'withholding tax', 'tax amount', 'ภาษีหัก ณ ที่จ่ายรวม'],
  count: ['จำนวนรายการ', 'จำนวนครั้ง', 'จำนวนธุรกรรม', 'count', 'รายการ'],
  tin: ['เลขประจำตัวผู้เสียภาษี', 'เลขประจำตัวผู้เสียภาษีอากร', 'เลขผู้เสียภาษี', 'เลขประจำตัว', 'tin', 'tax id', 'taxpayer id', 'tax identification number'],
};
const headerKey = (value: unknown) => String(value ?? '').replace(/[\s_-]+/g, '').trim().toLocaleLowerCase();
function cleanCell(value: unknown): string {
  if (value && typeof value === 'object') {
    if ('result' in value) return cleanCell((value as { result?: unknown }).result);
    if ('richText' in value) return (value as { richText: Array<{ text: string }> }).richText.map((part) => part.text).join('').trim();
    if ('text' in value) return String((value as { text?: unknown }).text ?? '').trim();
  }
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}
function amountValue(value: unknown, row: number, column: string): number {
  if (value === null || value === undefined || cleanCell(value) === '') return 0;
  if (typeof value === 'number') return value;
  let text = cleanCell(value).replace(/,/g, '').replace(/฿|บาท/g, '');
  if (text.startsWith('(') && text.endsWith(')')) text = '-' + text.slice(1, -1);
  const result = Number(text);
  if (!Number.isFinite(result)) throw new Error(`แถว ${row}: ค่าในคอลัมน์${column}ไม่ใช่ตัวเลข (${text})`);
  return result;
}
async function readWorkbook(file: File): Promise<Group[]> {
  const ExcelJS = await import('exceljs');
  let workbook: ExcelJSTypes.Workbook;
  try {
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
  } catch {
    throw new Error(`เปิดไฟล์ ${file.name} ไม่ได้ กรุณาใช้ไฟล์ .xlsx ที่ถูกต้อง`);
  }
  const sheet = workbook.worksheets[0];
  if (!sheet || sheet.rowCount < 1) throw new Error(`ไฟล์ ${file.name} ไม่มีข้อมูล`);
  const header = Array.from({ length: sheet.getRow(1).cellCount }, (_, i) => cleanCell(sheet.getRow(1).getCell(i + 1).value));
  const positions: Record<string, number> = {};
  for (const field of Object.keys(aliases) as Array<keyof typeof aliases>) {
    const accepted = aliases[field].map(headerKey);
    const index = header.findIndex((cell) => accepted.includes(headerKey(cell)));
    if (index >= 0) positions[field] = index + 1;
  }
  if (!positions.payer || !positions.income) throw new Error(`ไฟล์ ${file.name}: ต้องมีคอลัมน์ชื่อผู้มีหน้าที่หักและประเภทเงินได้`);
  if (!positions.amount) throw new Error(`ไฟล์ ${file.name}: ต้องมีคอลัมน์จำนวนเงิน`);
  const groups = new Map<string, Group>();
  for (let rowNo = 2; rowNo <= sheet.rowCount; rowNo++) {
    const row = sheet.getRow(rowNo);
    const payer = cleanCell(row.getCell(positions.payer).value);
    const income = cleanCell(row.getCell(positions.income).value);
    if (!payer && !income) continue;
    if (!payer || !income) throw new Error(`ไฟล์ ${file.name} แถว ${rowNo}: ต้องระบุชื่อผู้หักและประเภทเงินได้`);
    const key = `${payer.toLocaleLowerCase()}\u0000${income.toLocaleLowerCase()}`;
    const group = groups.get(key) ?? { payer, income, amount: 0, tax: 0, count: 0, tins: [] };
    group.amount += amountValue(row.getCell(positions.amount).value, rowNo, 'จำนวนเงิน');
    group.tax += positions.tax ? amountValue(row.getCell(positions.tax).value, rowNo, 'ภาษี') : 0;
    group.count += positions.count ? amountValue(row.getCell(positions.count).value, rowNo, 'จำนวนรายการ') : 1;
    if (positions.tin) {
      const tinCell = row.getCell(positions.tin);
      const tin = (tinCell.text || cleanCell(tinCell.value)).trim();
      if (tin && !group.tins.includes(tin)) group.tins.push(tin);
    }
    groups.set(key, group);
  }
  if (!groups.size) throw new Error(`ไฟล์ ${file.name}: ไม่พบแถวข้อมูล`);
  return [...groups.values()];
}
function compareWorkbooks(left: Group[], right: Group[], rules: Rule[]): Item[] {
  const replacements = new Map(rules.map((rule) => [rule.source.trim().toLocaleLowerCase(), rule.target.trim()]));
  const merged = new Map<string, Item & { a_rows: number; b_rows: number }>();
  for (const [side, groups] of [['a', left], ['b', right]] as const) {
    for (const group of groups) {
      const payer = replacements.get(group.payer.trim().toLocaleLowerCase()) ?? group.payer;
      const key = `${payer.toLocaleLowerCase()}\u0000${group.income.toLocaleLowerCase()}`;
      const item = merged.get(key) ?? { payer, income_type: group.income, a_count: 0, b_count: 0, a_amount: 0, b_amount: 0, amount_diff: 0, a_tax: 0, b_tax: 0, tax_diff: 0, status: '', a_rows: 0, b_rows: 0 };
      item[`${side}_amount`] += group.amount;
      item[`${side}_tax`] += group.tax;
      item[`${side}_count`] += group.count;
      item[`${side}_rows`] += group.count;
      merged.set(key, item);
    }
  }
  return [...merged.values()].map((item) => {
    item.amount_diff = item.b_amount - item.a_amount;
    item.tax_diff = item.b_tax - item.a_tax;
    item.status = Math.abs(item.amount_diff) < 0.005 && Math.abs(item.tax_diff) < 0.005 ? 'ตรงกัน' : (!item.a_rows ? 'พบเฉพาะชุด B' : (!item.b_rows ? 'พบเฉพาะชุด A' : 'ยอดไม่ตรงกัน'));
    return item;
  }).sort((a, b) => a.payer.localeCompare(b.payer, 'th') || a.income_type.localeCompare(b.income_type, 'th'));
}
function formLabel(file: File, fallback: string) {
  const name = file.name.replace(/[\s._-]/g, '').toLocaleLowerCase();
  if (name.includes('ภงด2ก') || /pnd2k/.test(name)) return 'ภงด2ก';
  if (name.includes('ภงด2') || /pnd2(?!k)/.test(name)) return 'ภงด2';
  return fallback;
}
function normalizedPayer(name: string, rules: Rule[]) {
  return new Map(rules.map((rule) => [rule.source.trim().toLocaleLowerCase(), rule.target.trim()])).get(name.trim().toLocaleLowerCase()) ?? name;
}
function summarizeGroups(groups: Group[], rules: Rule[]) {
  const summary = new Map<string, { payer: string; income: string; count: number; amount: number; tax: number; names: Set<string>; tins: Set<string> }>();
  for (const group of groups) {
    const payer = normalizedPayer(group.payer, rules);
    const key = `${payer.toLocaleLowerCase()}\u0000${group.income.toLocaleLowerCase()}`;
    const item = summary.get(key) ?? { payer, income: group.income, count: 0, amount: 0, tax: 0, names: new Set<string>(), tins: new Set<string>() };
    item.count += group.count;
    item.amount += group.amount;
    item.tax += group.tax;
    item.names.add(group.payer);
    group.tins.forEach((tin) => item.tins.add(tin));
    summary.set(key, item);
  }
  return [...summary.values()].sort((a, b) => a.payer.localeCompare(b.payer, 'th') || a.income.localeCompare(b.income, 'th'));
}
function App() {
  const [a, setA] = useState<File | null>(null), [b, setB] = useState<File | null>(null), [result, setResult] = useState<Result | null>(null), [rules, setRules] = useState<Rule[]>([]), [source, setSource] = useState(''), [target, setTarget] = useState(''), [busy, setBusy] = useState(false), [error, setError] = useState(''), [query, setQuery] = useState(''), [status, setStatus] = useState('ทั้งหมด'), [tab, setTab] = useState<'compare' | 'rules'>('compare'), [drag, setDrag] = useState<'a' | 'b' | null>(null), [ownerUser, setOwnerUser] = useState<import('firebase/auth').User | null>(null), [loginEmail, setLoginEmail] = useState(''), [loginPassword, setLoginPassword] = useState('');
  const inputA = useRef<HTMLInputElement>(null), inputB = useRef<HTMLInputElement>(null);
  useEffect(() => watchNameRules(setRules, (err) => setError(err.message)), []);
  useEffect(() => watchAuth(setOwnerUser), []);
  const isAdmin = !firebaseEnabled || Boolean(ownerUser && ownerUser.email?.toLowerCase() === adminEmail.toLowerCase());
  const shown = useMemo(() => result?.items.filter((item) => (`${item.payer} ${item.income_type}`).toLowerCase().includes(query.toLowerCase()) && (status === 'ทั้งหมด' || item.status === status)) || [], [result, query, status]);
  const summary = useMemo(() => { const items = result?.items || []; return { total: items.length, matched: items.filter((item) => item.status === 'ตรงกัน').length, issues: items.filter((item) => item.status !== 'ตรงกัน').length, delta: items.reduce((sum, item) => sum + item.amount_diff, 0) }; }, [result]);
  async function run() {
    if (!a || !b) return;
    setBusy(true); setError('');
    try {
      const [left, right] = await Promise.all([readWorkbook(a), readWorkbook(b)]);
      setResult({ label_a: formLabel(a, 'ชุด A'), label_b: formLabel(b, 'ชุด B'), file_a: a.name, file_b: b.name, items: compareWorkbooks(left, right, rules), groups_a: left, groups_b: right, rules: [...rules] });
    } catch (err) { setError(err instanceof Error ? err.message : 'อ่านไฟล์ไม่สำเร็จ'); }
    finally { setBusy(false); }
  }
  async function addRule(event: React.FormEvent) {
    event.preventDefault(); setError('');
    const oldName = source.trim(), standardName = target.trim();
    if (!oldName || !standardName || oldName.toLocaleLowerCase() === standardName.toLocaleLowerCase()) { setError('กรุณาระบุชื่อเดิมและชื่อมาตรฐานที่ต่างกัน'); return; }
    try { await createNameRule(oldName, standardName); setSource(''); setTarget(''); }
    catch (err) { setError(err instanceof Error ? err.message : 'บันทึกกฎไม่สำเร็จ'); }
  }
  async function removeRule(id: string) {
    try { await deleteNameRule(id); }
    catch (err) { setError(err instanceof Error ? err.message : 'ลบกฎไม่สำเร็จ'); }
  }
  async function ownerLogin(event: React.FormEvent) {
    event.preventDefault(); setError('');
    try { await signInOwner(loginEmail.trim(), loginPassword); setLoginPassword(''); }
    catch { setError('เข้าสู่ระบบผู้ดูแลไม่สำเร็จ กรุณาตรวจสอบอีเมลและรหัสผ่าน'); }
  }
  async function ownerLogout() { await signOutOwner(); }
  async function exportXlsx() {
    if (!result) return;
    setError('');
    try {
      if (!result.items.length) throw new Error('ไม่มีแถวผลเปรียบเทียบสำหรับส่งออก กรุณาเปรียบเทียบไฟล์ใหม่');
      const ExcelJS = await import('exceljs');
      const workbook = new ExcelJS.Workbook();
      const headerFont = { name: 'Tahoma', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      const dataFont = { name: 'Tahoma', size: 11, color: { argb: 'FF000000' } };
      const headerFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF1F4E78' } };
      const moneyFormat = '#,##0.00;[Red]-#,##0.00;0.00';
      function formatSheet(sheet: ExcelJSTypes.Worksheet, headers: string[], widths: number[], moneyColumns: number[]) {
        sheet.addRow(headers);
        sheet.columns.forEach((column, index) => { column.width = widths[index]; });
        sheet.views = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2' }];
        sheet.getRow(1).height = 32.1;
        sheet.getRow(1).eachCell((cell) => {
          cell.font = headerFont;
          cell.fill = headerFill;
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        });
        for (let row = 2; row <= sheet.rowCount; row++) {
          sheet.getRow(row).eachCell((cell) => {
            cell.font = dataFont;
            cell.alignment = { vertical: 'top', wrapText: true };
          });
          moneyColumns.forEach((column) => { sheet.getCell(row, column).numFmt = moneyFormat; });
        }
      }
      const comparison = workbook.addWorksheet('เปรียบเทียบ');
      formatSheet(comparison,
        ['ชื่อผู้มีหน้าที่หัก (ชื่อมาตรฐาน)', 'ประเภทเงินได้', `สถานะใน ${result.label_a}`, `จำนวนรายการ ${result.label_a}`, `มูลค่า ${result.label_a}`, `ภาษี ${result.label_a}`, `สถานะใน ${result.label_b}`, `จำนวนรายการ ${result.label_b}`, `มูลค่า ${result.label_b}`, `ภาษี ${result.label_b}`, 'ผลการเปรียบเทียบ'],
        [55, 45, 14, 18, 13, 12, 15, 19, 14, 12, 18], [5, 6, 9, 10]);
      for (const item of result.items) {
        const comparisonStatus = item.a_count === 0 ? `มีเฉพาะ ${result.label_b}` : item.b_count === 0 ? `มีเฉพาะ ${result.label_a}` : item.status === 'ตรงกัน' ? 'ไม่ต่างกัน' : 'ยอดไม่ตรงกัน';
        comparison.addRow([item.payer, item.income_type, item.a_count ? 'มี' : 'ไม่มี', item.a_count, item.a_amount, item.a_tax, item.b_count ? 'มี' : 'ไม่มี', item.b_count, item.b_amount, item.b_tax, comparisonStatus]);
        const statusCell = comparison.getCell(comparison.rowCount, 11);
        statusCell.font = { name: 'Tahoma', size: 11, bold: true, color: { argb: comparisonStatus === 'ไม่ต่างกัน' ? 'FF006100' : 'FF9C0006' } };
      }
      const summaryHeaders = ['ชื่อผู้มีหน้าที่หัก (ชื่อมาตรฐาน)', 'ประเภทเงินได้', 'จำนวนรายการ', 'มูลค่ารวม', 'ภาษีหัก ณ ที่จ่ายรวม', 'ชื่อที่พบในต้นฉบับ', 'TIN ที่พบ'];
      for (const [label, groups] of [[result.label_a, result.groups_a], [result.label_b, result.groups_b]] as const) {
        const summary = workbook.addWorksheet(`สรุป_${label}`);
        formatSheet(summary, summaryHeaders, [55, label === 'ภงด2ก' ? 22 : 45, 13, 12, 22, 55, 12], [4, 5]);
        for (const item of summarizeGroups(groups, result.rules)) {
          summary.addRow([item.payer, item.income, item.count, item.amount, item.tax, [...item.names].sort((a, b) => a.localeCompare(b, 'th')).join(' / '), [...item.tins].sort().join(', ')]);
        }
      }
      const bytes = await workbook.xlsx.writeBuffer();
      if (!bytes.byteLength) throw new Error('สร้างไฟล์ Excel ไม่สำเร็จ กรุณาลองอีกครั้ง');
      const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const link = document.createElement('a');
      link.href = url;
      link.download = 'tax-comparison.xlsx';
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'ส่งออกไฟล์ Excel ไม่สำเร็จ');
    }
  }
  function drop(event: React.DragEvent, slot: 'a' | 'b') { event.preventDefault(); setDrag(null); const file = event.dataTransfer.files[0]; if (file) setFile(file, slot); }
  function setFile(file: File, slot: 'a' | 'b') { if (!file.name.toLowerCase().endsWith('.xlsx')) { setError('กรุณาเลือกไฟล์ Excel นามสกุล .xlsx'); return; } if (file.size > 20 * 1024 * 1024) { setError('ไฟล์ต้องมีขนาดไม่เกิน 20 MB'); return; } slot === 'a' ? setA(file) : setB(file); setResult(null); setError(''); }
  
return <div className="app"><aside className="sidebar"><div className="brand"><div className="brand-mark"><span></span><span></span><span></span><span></span></div><div><b>taxwise</b><small>DATA STUDIO</small></div></div><div className="workspace-label">WORKSPACE</div><div className="workspace"><div className="workspace-icon">N</div><div><b>Northstar Group</b><small>Finance team</small></div><ChevronDown size={15}/></div><div className="nav-label">จัดการข้อมูล</div><button className={`nav-item ${tab==='compare'?'active':''}`} onClick={()=>setTab('compare')}><LayoutDashboard size={17}/>ภาพรวมและเปรียบเทียบ</button><button className={`nav-item ${tab==='rules'?'active':''}`} onClick={()=>setTab('rules')}><Settings2 size={17}/>กฎปรับชื่อ<span className="nav-count">{rules.length}</span></button><div className="side-bottom"><div className="secure"><ShieldCheck size={16}/><div><b>ข้อมูลปลอดภัย</b><small>ประมวลผลในเครื่องคุณ</small></div></div><div className="profile"><div className="avatar">PK</div><div><b>Ploy K.</b><small>Administrator</small></div><span>•••</span></div></div></aside><main className="main"><header><div className="crumb">เครื่องมือ <span>/</span> <b>{tab==='compare'?'เปรียบเทียบข้อมูล':'กฎปรับชื่อ'}</b></div><div className="header-right"><span className="live"><i/>ระบบพร้อมใช้งาน</span><span className="date">26 พ.ย. 2569</span><div className="avatar sm">PK</div></div></header><div className="content">{tab==='compare'?<><div className="title-row"><div><div className="eyebrow">TAX RECONCILIATION <span>•</span> FY 2025</div><h1>เปรียบเทียบข้อมูลภาษี</h1><p>ตรวจสอบยอดเงินได้และภาษีหัก ณ ที่จ่าย ระหว่างสองชุดข้อมูล</p></div><button className="outline" onClick={()=>setTab('rules')}><Settings2 size={16}/> จัดการกฎปรับชื่อ <ArrowRight size={15}/></button></div><section className="upload-card"><div className="section-heading"><div className="step">01</div><div><h2>อัปโหลดไฟล์เพื่อเปรียบเทียบ</h2><p>เลือกไฟล์ Excel สองชุดที่ต้องการตรวจสอบ</p></div><div className="format-pill"><FileSpreadsheet size={14}/> XLSX</div></div><div className="upload-grid">{(['a','b'] as const).map((slot,i)=>{const file=slot==='a'?a:b;return <div className={`dropzone ${drag===slot?'dragging':''} ${file?'has-file':''}`} key={slot} onDragOver={e=>{e.preventDefault();setDrag(slot)}} onDragLeave={()=>setDrag(null)} onDrop={e=>drop(e,slot)} onClick={()=>slot==='a'?inputA.current?.click():inputB.current?.click()}><input ref={slot==='a'?inputA:inputB} type="file" accept=".xlsx" hidden onChange={e=>e.target.files?.[0]&&setFile(e.target.files[0],slot)}/>{file?<><div className="file-icon"><FileSpreadsheet size={21}/></div><div className="file-info"><b>{file.name}</b><small>{(file.size/1024).toFixed(0)} KB · พร้อมเปรียบเทียบ</small></div><button className="remove" onClick={e=>{e.stopPropagation();slot==='a'?setA(null):setB(null)}}><X size={15}/></button></>:<><div className="upload-icon"><Upload size={19}/></div><div className="drop-copy"><b>{i===0?'ชุดข้อมูล A':'ชุดข้อมูล B'} <span className="required">*</span></b><small>ลากไฟล์มาวาง หรือ <em>เลือกไฟล์</em></small></div><div className="slot-badge">0{i+1}</div></>}</div>})}</div><div className="upload-footer"><div className="hint"><span className="hint-dot">i</span> รองรับไฟล์ .xlsx ขนาดไม่เกิน 20 MB · ระบบตรวจจับคอลัมน์อัตโนมัติ</div><button className="primary" disabled={!a||!b||busy} onClick={run}>{busy?<LoaderCircle className="spin" size={16}/>:<ArrowDownUp size={16}/>}เปรียบเทียบข้อมูล</button></div>{error&&<div className="error">{error}</div>}</section><div className="results-head"><div><div className="eyebrow">ANALYSIS <span>•</span> ผลลัพธ์ล่าสุด</div><h2>สรุปผลการตรวจสอบ</h2></div>{result&&<button className="outline export" onClick={exportXlsx}><Download size={16}/> ส่งออก Excel</button>}</div>{result?<><div className="summary-grid"><div className="metric"><div className="metric-top"><span>รายการทั้งหมด</span><div className="metric-icon neutral"><FileSpreadsheet size={16}/></div></div><strong>{summary.total}</strong><small>กลุ่มผู้หักและประเภทเงินได้</small></div><div className="metric"><div className="metric-top"><span>ยอดตรงกัน</span><div className="metric-icon green"><Check size={16}/></div></div><strong>{summary.matched}<small className="unit">รายการ</small></strong><small className="green-text">{summary.total?Math.round(summary.matched/summary.total*100):0}% ของรายการทั้งหมด</small></div><div className="metric"><div className="metric-top"><span>ต้องตรวจสอบ</span><div className="metric-icon amber"><Filter size={15}/></div></div><strong>{summary.issues}<small className="unit">รายการ</small></strong><small className="amber-text">ยอดต่างหรือพบในชุดเดียว</small></div><div className="metric"><div className="metric-top"><span>ผลต่างสุทธิ · เงินได้</span><div className="metric-icon neutral"><Wallet size={16}/></div></div><strong className="amount">{fmt(summary.delta)}</strong><small>ชุด B − ชุด A</small></div></div><section className="table-card"><div className="table-toolbar"><div className="table-title"><h3>รายละเอียดรายกลุ่ม</h3><span>{shown.length} รายการ</span></div><div className="table-tools"><div className="search"><Search size={15}/><input placeholder="ค้นหาชื่อหรือประเภทเงินได้" value={query} onChange={e=>setQuery(e.target.value)}/></div><select value={status} onChange={e=>setStatus(e.target.value)}><option>ทั้งหมด</option><option>ตรงกัน</option><option>ยอดไม่ตรงกัน</option><option>พบเฉพาะชุด A</option><option>พบเฉพาะชุด B</option></select></div></div><div className="table-wrap"><table><thead><tr><th>ชื่อผู้มีหน้าที่หัก / ประเภทเงินได้</th><th>{result.label_a}<small>เงินได้</small></th><th>{result.label_b}<small>เงินได้</small></th><th>ผลต่าง</th><th>ภาษีหัก ณ ที่จ่าย <small>A → B</small></th><th>สถานะ</th></tr></thead><tbody>{shown.map((x,i)=><tr key={i}><td><b>{x.payer}</b><small className="sub">{x.income_type}</small></td><td className="num">{fmt(x.a_amount)}</td><td className="num">{fmt(x.b_amount)}</td><td className={`num ${x.amount_diff===0?'muted':x.amount_diff>0?'positive':'negative'}`}>{x.amount_diff===0?'—':fmt(x.amount_diff)}</td><td className="tax">{fmt(x.a_tax)} <span>→</span> {fmt(x.b_tax)}</td><td><span className={`status ${x.status==='ตรงกัน'?'ok':'warn'}`}><i/>{x.status}</span></td></tr>)}</tbody></table>{!shown.length&&<div className="empty">ไม่พบรายการที่ตรงกับตัวกรอง</div>}</div><div className="table-foot"><span>แสดง {shown.length} จาก {result.items.length} รายการ</span><span><History size={13}/> คำนวณจากไฟล์ {result.file_a} และ {result.file_b}</span></div></section></>:<div className="empty-state"><div className="empty-art"><div className="art-sheet"><div/><div/><div/></div><div className="art-arrow"><ArrowRight size={20}/></div><div className="art-sheet second"><div/><div/><div/></div></div><h3>พร้อมตรวจสอบข้อมูลของคุณ</h3><p>อัปโหลดไฟล์ Excel สองชุดด้านบน แล้วระบบจะจัดกลุ่มตามผู้มีหน้าที่หักและประเภทเงินได้ให้อัตโนมัติ</p><div className="empty-steps"><span><i>1</i> อัปโหลดไฟล์</span><ArrowRight size={14}/><span><i>2</i> ตรวจสอบผลต่าง</span><ArrowRight size={14}/><span><i>3</i> ส่งออก Excel</span></div></div>}</>:<><div className="title-row"><div><div className="eyebrow">CONFIGURATION <span>•</span> MASTER DATA</div><h1>กฎปรับชื่อผู้หัก</h1><p>รวมชื่อที่สะกดต่างกันให้เป็นชื่อมาตรฐานเดียวกันก่อนเปรียบเทียบ</p></div><button className="outline" onClick={()=>setTab('compare')}><ArrowRight size={15}/> กลับไปเปรียบเทียบ</button></div><div className="rules-layout"><section className="upload-card rule-form"><div className="section-heading"><div className="step"><Plus size={17}/></div><div><h2>เพิ่มกฎใหม่</h2><p>กำหนดชื่อเดิมและชื่อมาตรฐานที่ต้องการใช้</p></div></div>{isAdmin ? <form onSubmit={addRule}><label>ชื่อที่พบในไฟล์ <span>*</span><input required value={source} onChange={e=>setSource(e.target.value)} placeholder="เช่น บริษัท เอบีซี จำกัด"/></label><div className="rule-arrow"><ArrowDownUp size={16}/></div><label>ชื่อมาตรฐาน <span>*</span><input required value={target} onChange={e=>setTarget(e.target.value)} placeholder="เช่น บริษัท ABC จำกัด"/></label>{error&&<div className="error">{error}</div>}<button className="primary full" disabled={!source||!target}><Plus size={16}/> บันทึกกฎอย่างถาวร</button><button type="button" className="admin-signout" onClick={ownerLogout}>ออกจากระบบผู้ดูแล</button><div className="persist-note"><ShieldCheck size={14}/>{firebaseEnabled ? 'บันทึกเฉพาะกฎชื่อร่วมกันใน Firebase' : 'โหมดทดลอง: กฎชื่อเก็บในเบราว์เซอร์นี้'}</div></form> : <form className="admin-login" onSubmit={ownerLogin}><div className="admin-login-title">จัดการกฎชื่อ</div><p>เข้าสู่ระบบด้วยบัญชีผู้ดูแล Firebase เพื่อเพิ่มหรือลบกฎ ผู้ใช้งานทั่วไปไม่ต้องเข้าสู่ระบบ</p><label>อีเมลผู้ดูแล<input type="email" autoComplete="username" required value={loginEmail} onChange={(event) => setLoginEmail(event.target.value)} /></label><label>รหัสผ่าน<input type="password" autoComplete="current-password" required value={loginPassword} onChange={(event) => setLoginPassword(event.target.value)} /></label>{error && <div className="error">{error}</div>}<button className="primary full"><ShieldCheck size={16}/> เข้าสู่ระบบผู้ดูแล</button></form>}</section><section className="table-card rules-list"><div className="table-toolbar"><div className="table-title"><h3>กฎที่บันทึกไว้</h3><span>{rules.length} กฎ</span></div></div>{rules.length?<div className="rules-rows">{rules.map(r=><div className="rule-row" key={r.id}><div className="rule-name"><small>ชื่อที่พบ</small><b>{r.source}</b></div><ArrowRight size={17} className="rule-connect"/><div className="rule-name target"><small>ชื่อมาตรฐาน</small><b>{r.target}</b></div><button className="remove-rule" onClick={()=>removeRule(r.id)} title="ลบกฎ"><Trash2 size={16}/></button></div>)}</div>:<div className="rules-empty"><div className="empty-rule-icon"><Settings2 size={22}/></div><b>ยังไม่มีกฎปรับชื่อ</b><span>เพิ่มกฎด้านซ้ายเพื่อรวมชื่อที่สะกดต่างกัน</span></div>}</section></div></>}</div><footer>Taxwise <span>·</span> Internal Finance Tools <span className="foot-right">ไฟล์ต้นฉบับและผลเปรียบเทียบไม่ถูกส่งไปเก็บบนเซิร์ฟเวอร์</span></footer></main></div>}
export default App;
