"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { buildCustomerInteractiveReportHtml } from "./interactive-report";
import {
  Bar, BarChart, CartesianGrid, Cell, ComposedChart, Legend, Line, LineChart,
  Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis,
} from "recharts";
import {
  Activity, ArrowDownRight, ArrowUpRight, Boxes, CalendarClock, Check, CheckCircle2,
  ChevronDown, ChevronLeft, ChevronRight, ClipboardList, Factory,
  Download, FileSpreadsheet, Filter, Leaf, LoaderCircle, PackageOpen, Printer,
  Search, ShieldCheck, Sparkles, TimerReset, Trash2, TrendingUp, RotateCcw, ArrowUp,
  Truck, UploadCloud, Warehouse, X,
} from "lucide-react";

type CellValue = string | number | boolean | Date | null | undefined;
type Matrix = CellValue[][];
type LoadStage = "idle" | "reading" | "analyzing" | "success" | "error";
type MovementRow = { part: string; description: string; start: number; inbound: number; outbound: number; end: number };
type OrderRow = { number: string; po: string; pro: string; consignee: string; creation: Date; confirmation: Date | null; creationToConfirmation: number | null; pick: Date | null; confirmationToPick: number | null; pickup: Date | null; pickToPickup: number | null; cycleDays: number | null; booking: Date | null; sailing: Date | null; arrival: Date | null; notes: string };
type AgingRow = { status: string; receipt: string; po: string; receiptDate: Date; part: string; description: string; consignee: string; pieces: number; ageDays: number; ageBucket: string };
type DetailRow = { status: string; consignee: string; receipt: string; location: string; description: string; pieces: number; packageType: string; entryNumber: string; entryDate: Date | null; customsDate: Date | null; part: string };
type VolumeSeries = { year: string; values: number[] };
type Co2Row = { number: string; po: string; consignee: string; weight: number; distance: number; factor: number; emission: number };
type ReportData = { overview: { months: string[]; rows: { label: string; values: (string | number | null)[] }[] }; fileName: string; customer: string; period: string; reportDate: Date | null; generatedDate: string; movements: MovementRow[]; orders: OrderRow[]; aging: AgingRow[]; details: DetailRow[]; inboundVolumes: VolumeSeries[]; outboundVolumes: VolumeSeries[]; accidentDays: number | null; accidentAsOf: string; co2: Co2Row[]; recognizedSheets: string[] };

const HIDDEN_OVERVIEW_ROWS = new Set(["Inbound & Outbound Volume", "Shipping Accuracy", "PO's Shipped", "Cases Adjusted", "WTDC Damaged Cases", "Campari Damaged Cases (Unloaded damage)", "Audit Accuracy", "Full Pallet Audits (Total pallets)", "Full Pallets Audits Failed (Total Pallets)", "Full Pallet Accuracy %", "Loose Case Pallet Audits (QTY of Loose cases)", "Loose Case Pallet Audits Failed (QTY of Loose cases)", "Loose Case Accuracy %"]);

const BRAND = { navy: "#002D62", gold: "#AB8820", navyMid: "#315B84", navyLight: "#6F8AA8", goldLight: "#D2B05A" };
const SERIES = [BRAND.navy, BRAND.gold, BRAND.navyMid, BRAND.goldLight, BRAND.navyLight];
const RECEIPT_COLORS = [
  "#002D62", "#AB8820", "#2F6B8A", "#6B5B95", "#2E7D5B", "#C46B2D",
  "#4F6D3A", "#8C4A5B", "#4B6CB7", "#9A7B4F", "#7A5EA8", "#367A78",
];
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function numberValue(value: CellValue) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return 0;
  const parsed = Number(value.replace(/,/g, "").replace(/[()%$]/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function idValue(value: CellValue) {
  if (typeof value === "number") return Number.isSafeInteger(value) ? String(value) : value.toLocaleString("en-US", { useGrouping: false, maximumSignificantDigits: 21 });
  return String(value ?? "").trim();
}

function dateValue(value: CellValue): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? new Date(parsed.y, parsed.m - 1, parsed.d || 1) : null;
  }
  if (typeof value !== "string" || !value.trim()) return null;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function dateFromPeriod(text: string) {
  const match = text.match(/^([A-Za-z]+)\s+\d{1,2}\s*[-–]\s*(\d{1,2}),\s*(\d{4})/);
  if (!match) return null;
  const parsed = new Date(`${match[1]} ${match[2]}, ${match[3]}`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatDate(value: Date | null) {
  return value ? new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric", year: "numeric" }).format(value) : "—";
}

function formatNumber(value: number, compact = false) {
  return new Intl.NumberFormat("en-US", { notation: compact ? "compact" : "standard", maximumFractionDigits: compact ? 1 : 0 }).format(value || 0);
}

function businessDaysBetween(start: Date, end: Date | null) {
  if (!end || end < start) return null;
  const cursor = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const finish = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  let days = 0;
  while (cursor < finish) { cursor.setDate(cursor.getDate() + 1); if (cursor.getDay() !== 0 && cursor.getDay() !== 6) days += 1; }
  return days;
}

function normalizedSheetName(value: string) { return value.toLowerCase().replace(/and/g, "").replace(/[^a-z0-9]/g, ""); }
function findSheet(workbook: XLSX.WorkBook, candidates: string[]) {
  const normalized = new Map(workbook.SheetNames.map((name) => [normalizedSheetName(name), name]));
  for (const candidate of candidates) { const found = normalized.get(normalizedSheetName(candidate)); if (found) return found; }
  return "";
}
function sheetMatrix(workbook: XLSX.WorkBook, name: string): Matrix {
  if (!name || !workbook.Sheets[name]) return [];
  return XLSX.utils.sheet_to_json<CellValue[]>(workbook.Sheets[name], { header: 1, defval: null, raw: true }) as Matrix;
}
function parseVolumes(matrix: Matrix): VolumeSeries[] {
  if (matrix.length < 3) return [];
  return matrix.slice(2).filter((row) => numberValue(row[0]) > 2000).map((row) => ({ year: String(Math.trunc(numberValue(row[0]))), values: MONTHS.map((_, index) => numberValue(row[index + 1])) }));
}
function mostCommon(values: string[]) {
  const counts = new Map<string, number>(); values.filter(Boolean).forEach((value) => counts.set(value, (counts.get(value) ?? 0) + 1));
  return Array.from(counts).sort((a, b) => b[1] - a[1])[0]?.[0] ?? "Customer";
}

function parseWorkbook(workbook: XLSX.WorkBook, fileName: string): ReportData {
  const overviewName = findSheet(workbook, ["Overview"]);
  const summaryName = findSheet(workbook, ["In Out Summary", "In and Out Summary"]);
  const detailName = findSheet(workbook, ["In Out Detailed", "In and Out Detailed"]);
  const inboundName = findSheet(workbook, ["Inbound Comparison"]);
  const outboundName = findSheet(workbook, ["Outbound Comparison"]);
  const kpiName = findSheet(workbook, ["KPI"]);
  const agingName = findSheet(workbook, ["Inventory >1 Year", "Inventory Over 1 Year"]);
  const co2Name = findSheet(workbook, ["CO2", "Carbon"]);
  const accidentName = findSheet(workbook, ["Days Accident Free", "Accident Free Days"]);
  const recognizedSheets = [overviewName, summaryName, detailName, inboundName, outboundName, kpiName, agingName, co2Name, accidentName].filter(Boolean);
  if (!recognizedSheets.length) throw new Error("This workbook does not contain any recognized customer-report worksheets.");

  const overviewMatrix = sheetMatrix(workbook, overviewName);
  const overview = { months: (overviewMatrix[0] ?? []).slice(1).map(value => { const date = dateValue(value); return date ? date.toLocaleDateString("en-US", { month: "short", year: "numeric" }) : String(value ?? ""); }), rows: overviewMatrix.slice(1).filter(row => row[0]).map(row => ({ label: String(row[0]), values: row.slice(1).map(value => typeof value === "number" ? value : value == null ? null : String(value)) })) };
  for (const [name, rowIndex, expected] of [
    [summaryName, 5, ["Part Number", "Description", "Start", "In", "Out", "End"]],
    [detailName, 5, ["Status", "Entry Date", "Entry Number", "Warehouse Receipt", "Purchase Order Number", "Part Number", "Description", "Pieces", "Cargo Release", "Cargo Release:Consignee", "Cargo Release:Purchase Order Number", "Out Date"]],
    [kpiName, 0, ["Number", "PO Number", "PRO Number", "Consignee", "Creation Date", "Confirmation Date", "Days Creation-Confirmation", "Pick Date", "Days Confirmation-Pick", "Delivery Date", "Days Pick-Delivery", "Notes"]]
  ] as [string, number, string[]][]) {
    if (!name) throw new Error("Missing a required worksheet: In Out Summary, In Out Detailed and KPI are required.");
    const header = sheetMatrix(workbook, name)[rowIndex] ?? [];
    if (expected.some((label, index) => String(header[index] ?? "").trim().toLowerCase() !== label.toLowerCase())) throw new Error(`${name}: column headings do not match the standard customer KPI workbook. Please use the supplied Excel format.`);
  }
  const summary = sheetMatrix(workbook, summaryName);
  const accident = sheetMatrix(workbook, accidentName);
  const period = String(summary[2]?.[0] ?? "Reporting period not stated");
  const accidentHeading = String(accident[0]?.[0] ?? "");
  const accidentDateMatch = accidentHeading.match(/as of\s+(.+)$/i);
  const accidentDate = accidentDateMatch ? dateValue(accidentDateMatch[1]) : null;
  const reportDate = accidentDate ?? dateFromPeriod(period) ?? dateValue(summary[3]?.[0]);

  const movements: MovementRow[] = summary.slice(6).flatMap((row) => {
    if (row[0] === null || row[0] === undefined || !String(row[0]).trim()) return [];
    const part = idValue(row[0]); const start = numberValue(row[2]); const inbound = numberValue(row[3]); const outbound = numberValue(row[4]); const end = numberValue(row[5]);
    if (!part || !/\d/.test(part)) return [];
    return [{ part, description: String(row[1] ?? ""), start, inbound, outbound, end }];
  });

  const kpi = sheetMatrix(workbook, kpiName);
  const orders: OrderRow[] = kpi.slice(1).flatMap((row) => {
    const creation = dateValue(row[4]); if (!idValue(row[0]) || !creation) return [];
    const confirmation = dateValue(row[5]); const pick = dateValue(row[7]); const pickup = dateValue(row[9]);
    return [{ number: idValue(row[0]), po: idValue(row[1]), pro: idValue(row[2]), consignee: String(row[3] ?? "Unassigned"), creation, confirmation, creationToConfirmation: typeof row[6] === "number" ? row[6] : businessDaysBetween(creation, confirmation), pick, confirmationToPick: typeof row[8] === "number" ? row[8] : confirmation ? businessDaysBetween(confirmation, pick) : null, pickup, pickToPickup: typeof row[10] === "number" ? row[10] : pick ? businessDaysBetween(pick, pickup) : null, cycleDays: businessDaysBetween(creation, pickup), booking: null, sailing: null, arrival: null, notes: String(row[11] ?? "") }];
  });

  const agingMatrix = sheetMatrix(workbook, agingName);
  const aging: AgingRow[] = agingMatrix.slice(1).flatMap((row) => {
    const receiptDate = dateValue(row[3]); if (!idValue(row[1]) || !idValue(row[4]) || !receiptDate) return [];
    const ageDays = reportDate ? Math.max(0, Math.round((reportDate.getTime() - receiptDate.getTime()) / 86400000)) : 365;
    const ageBucket = ageDays < 730 ? "1–2 years" : ageDays < 1095 ? "2–3 years" : "3+ years";
    return [{ status: String(row[0] ?? ""), receipt: idValue(row[1]), po: idValue(row[2]), receiptDate, part: idValue(row[4]), description: String(row[5] ?? ""), consignee: String(row[6] ?? ""), pieces: numberValue(row[7]), ageDays, ageBucket }];
  });

  const detailMatrix = sheetMatrix(workbook, detailName); let currentPart = ""; const details: DetailRow[] = [];
  detailMatrix.slice(6).forEach((row) => {
    const statusText = String(row[0] ?? "").replace(/\u00a0/g, " ").trim(); const groupMatch = statusText.match(/^(\d{6,})\s/);
    if (groupMatch && !row[1] && !row[2]) currentPart = groupMatch[1];
    if (!row[3] || !row[5] || !row[6] || typeof row[7] !== "number") return;
    details.push({ status: statusText, consignee: String(row[9] ?? ""), receipt: idValue(row[3]), location: idValue(row[4]), description: String(row[6]), pieces: numberValue(row[7]), packageType: idValue(row[8]), entryNumber: idValue(row[2]), entryDate: dateValue(row[1]), customsDate: dateValue(row[11]), part: idValue(row[5]) || currentPart });
  });

  const co2Matrix = sheetMatrix(workbook, co2Name);
  const co2: Co2Row[] = co2Matrix.slice(1).flatMap((row) => {
    if (![0, 1, 2, 3, 4, 5].some((index) => row[index] !== null && row[index] !== "")) return [];
    const weight = numberValue(row[3]); const distance = numberValue(row[4]); const factor = numberValue(row[5]); const supplied = typeof row[6] === "number" ? row[6] : 0; const emission = supplied || ((weight / 2000) * distance * factor) / 1000000;
    return [{ number: idValue(row[0]), po: idValue(row[1]), consignee: String(row[2] ?? ""), weight, distance, factor, emission }];
  });
  const customer = mostCommon([...aging.map((row) => row.consignee), ...details.map((row) => row.consignee)]);
  return { overview, fileName, customer, period, reportDate, generatedDate: String(summary[3]?.[0] ?? ""), movements, orders, aging, details, inboundVolumes: parseVolumes(sheetMatrix(workbook, inboundName)), outboundVolumes: parseVolumes(sheetMatrix(workbook, outboundName)), accidentDays: typeof accident[1]?.[0] === "number" ? accident[1][0] : null, accidentAsOf: accidentDateMatch?.[1] ?? "", co2, recognizedSheets };
}

function volumeChartData(series: VolumeSeries[]) { return MONTHS.map((month, index) => Object.fromEntries([["month", month.slice(0, 3)], ...series.map((item) => [item.year, item.values[index] ?? 0])])) as Array<Record<string, string | number>>; }
function KpiCard({ icon, label, value, detail, target, tone = "navy" }: { icon: React.ReactNode; label: string; value: string; detail: string; target: string; tone?: "navy" | "gold" }) { return <a className={`kpi-card ${tone}`} href={`#${target}`} aria-label={`${label}: ${value}. Jump to ${target} section`}><div className="kpi-label"><span>{icon}</span><small>{label}</small></div><strong>{value}</strong><p>{detail}</p></a>; }
function EmptyPanel({ icon, text }: { icon: React.ReactNode; text: string }) { return <div className="empty-panel"><span>{icon}</span><p>{text}</p></div>; }

function groupAgingByPart(rows: AgingRow[]) {
  const groups = new Map<string, { part: string; descriptions: Set<string>; statuses: Set<string>; receipts: Set<string>; pos: Set<string>; earliest: Date; latest: Date; minAge: number; maxAge: number; pieces: number }>();
  for (const row of rows) {
    let group = groups.get(row.part);
    if (!group) { group = { part: row.part, descriptions: new Set(), statuses: new Set(), receipts: new Set(), pos: new Set(), earliest: row.receiptDate, latest: row.receiptDate, minAge: row.ageDays, maxAge: row.ageDays, pieces: 0 }; groups.set(row.part, group); }
    group.descriptions.add(row.description); group.statuses.add(row.status); group.receipts.add(row.receipt); if (row.po) group.pos.add(row.po);
    group.pieces += row.pieces; group.minAge = Math.min(group.minAge, row.ageDays); group.maxAge = Math.max(group.maxAge, row.ageDays);
    if (row.receiptDate < group.earliest) group.earliest = row.receiptDate;
    if (row.receiptDate > group.latest) group.latest = row.receiptDate;
  }
  return [...groups.values()];
}

function VolumeChart({ title, series }: { title: string; series: VolumeSeries[] }) {
  const years = [...series].sort((a,b) => Number(b.year)-Number(a.year));
  const [selected, setSelected] = useState<string[] | null>(null);
  const [showTable, setShowTable] = useState(false);
  const active = selected ?? years.slice(0,2).map(item => item.year);
  const colors = ["#002D62", "#B57B16", "#167D8D", "#8352A3", "#C44E52", "#477A38", "#6C647D"];
  return <article className="panel chart-panel volume-panel"><div className="panel-head"><div><span>{title}</span><strong>Monthly cases</strong></div><button className="button ghost" type="button" onClick={() => setShowTable(!showTable)}>{showTable ? "Hide values" : "Show values"}</button></div><div className="year-controls" aria-label={`${title} years`}>{years.map((item,index) => <label key={item.year} style={{ borderColor: active.includes(item.year) ? colors[index % colors.length] : undefined }}><input type="checkbox" checked={active.includes(item.year)} onChange={() => setSelected(active.includes(item.year) ? active.filter(year => year !== item.year) : [...active,item.year])} /><span style={{ color: colors[index % colors.length] }}>●</span>{item.year}</label>)}</div><div className="volume-plot">{series.length && active.length ? <ResponsiveContainer width="100%" height="100%"><LineChart data={volumeChartData(series)} margin={{top:20,right:24,left:12,bottom:12}}><CartesianGrid stroke="#e2e8f0" strokeDasharray="3 3" vertical={false}/><XAxis dataKey="month" interval="preserveStartEnd" minTickGap={15} tick={{fontSize:13,fill:"#334155"}} tickLine={false} axisLine={{stroke:"#cbd5e1"}} height={35}/><YAxis width={65} tick={{fontSize:12,fill:"#475569"}} tickLine={false} axisLine={false} tickFormatter={value => formatNumber(value,true)} /><Tooltip contentStyle={{borderRadius:8,border:"1px solid #cbd5e1",fontSize:14}} labelFormatter={label => `${label} · cases`} formatter={(value,name) => [formatNumber(Number(value)),String(name)]}/>{years.filter(item => active.includes(item.year)).map(item => <Line key={item.year} dataKey={item.year} type="linear" stroke={colors[years.findIndex(y=>y.year===item.year)%colors.length]} strokeWidth={3} dot={{r:3,strokeWidth:2,fill:"white"}} activeDot={{r:6}} connectNulls={false}/>)}</LineChart></ResponsiveContainer> : <EmptyPanel icon={<TrendingUp />} text={series.length ? "Select a year above to view its monthly volumes." : "Load a workbook to view monthly volumes."}/>}</div>{showTable && <div className="table-scroll"><table><thead><tr><th>Month</th>{years.filter(item=>active.includes(item.year)).map(item=><th key={item.year}>{item.year}</th>)}</tr></thead><tbody>{MONTHS.map((month,index)=><tr key={month}><td>{month}</td>{years.filter(item=>active.includes(item.year)).map(item=><td key={item.year}>{item.values[index] == null ? "—" : formatNumber(item.values[index])}</td>)}</tr>)}</tbody></table></div>}</article>;
}

export default function Home() {
  const [report, setReport] = useState<ReportData | null>(null);
  const [loadStage, setLoadStage] = useState<LoadStage>("idle");
  const [notice, setNotice] = useState("No customer report is loaded. Choose or drop an Excel workbook to begin.");
  const [dragging, setDragging] = useState(false);
  const [selectedPart, setSelectedPart] = useState("");
  const [selectedReceipt, setSelectedReceipt] = useState("");
  const [selectedConsignee, setSelectedConsignee] = useState("");
  const [movementSearch, setMovementSearch] = useState("");
  const [detailSearch, setDetailSearch] = useState("");
  const [agingSearch, setAgingSearch] = useState("");
  const [detailPage, setDetailPage] = useState(1);
  const [showExport, setShowExport] = useState(false);
  const [exportPackage, setExportPackage] = useState<{ url: string; name: string; size: number } | null>(null);
  const [exportStatus, setExportStatus] = useState<"idle" | "success" | "error">("idle");
  const [exportMessage, setExportMessage] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);
  const isLoading = loadStage === "reading" || loadStage === "analyzing";
  const clearSelections = () => { setSelectedPart(""); setSelectedReceipt(""); setSelectedConsignee(""); setMovementSearch(""); setDetailSearch(""); setAgingSearch(""); setDetailPage(1); };

  useEffect(() => () => { if (exportPackage) URL.revokeObjectURL(exportPackage.url); }, [exportPackage]);

  const handleFile = async (file?: File) => {
    if (!file || isLoading) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) { setLoadStage("error"); setNotice("Please choose an Excel workbook (.xlsx or .xls)."); return; }
    try {
      setLoadStage("reading"); setNotice(`Reading “${file.name}” locally…`); const buffer = await file.arrayBuffer();
      setLoadStage("analyzing"); setNotice("Connecting inventory, order, receipt, volume, safety, and CO₂ worksheets…"); await new Promise<void>((resolve) => window.setTimeout(resolve, 25));
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true }); const parsed = parseWorkbook(workbook, file.name);
      if (exportPackage) URL.revokeObjectURL(exportPackage.url);
      setShowExport(false); setExportPackage(null); setExportStatus("idle"); setExportMessage("");
      setReport(parsed); clearSelections(); setLoadStage("success");
      setNotice(`${parsed.recognizedSheets.length} report worksheets loaded from “${file.name}”. All workbook data remains in this browser.`);
    } catch (error) { setLoadStage("error"); setReport(null); setNotice(error instanceof Error ? `Could not load the workbook: ${error.message}` : "Could not load this workbook."); }
  };
  const clearReport = () => {
    if (report && !window.confirm("Clear this customer report and return the dashboard to blank?")) return;
    if (exportPackage) URL.revokeObjectURL(exportPackage.url);
    setReport(null); clearSelections(); setLoadStage("idle"); setNotice("Report cleared. The dashboard is blank and ready for another customer workbook."); setShowExport(false); setExportPackage(null); setExportStatus("idle"); setExportMessage(""); if (inputRef.current) inputRef.current.value = "";
  };

  const receiptParts = useMemo(() => { const values = new Set<string>(); if (!report || !selectedReceipt) return values; report.aging.filter((row) => row.receipt === selectedReceipt).forEach((row) => values.add(row.part)); report.details.filter((row) => row.receipt === selectedReceipt).forEach((row) => values.add(row.part)); return values; }, [report, selectedReceipt]);
  const movementRows = useMemo(() => { if (!report) return []; const search = movementSearch.trim().toLowerCase(); return report.movements.filter((row) => { if (selectedPart && row.part !== selectedPart) return false; if (selectedReceipt && !receiptParts.has(row.part)) return false; return !search || `${row.part} ${row.description}`.toLowerCase().includes(search); }); }, [report, movementSearch, selectedPart, selectedReceipt, receiptParts]);
  const agingRows = useMemo(() => { if (!report) return []; const search = agingSearch.trim().toLowerCase(); return report.aging.filter((row) => { if (selectedPart && row.part !== selectedPart) return false; if (selectedReceipt && row.receipt !== selectedReceipt) return false; return !search || `${row.part} ${row.receipt} ${row.description} ${row.po}`.toLowerCase().includes(search); }); }, [report, agingSearch, selectedPart, selectedReceipt]);
  const groupedAgingRows = useMemo(() => groupAgingByPart(agingRows), [agingRows]);
  const agingRowsForReceiptChart = useMemo(() => { if (!report) return []; const search = agingSearch.trim().toLowerCase(); return report.aging.filter((row) => { if (selectedPart && row.part !== selectedPart) return false; return !search || `${row.part} ${row.receipt} ${row.description} ${row.po}`.toLowerCase().includes(search); }); }, [report, agingSearch, selectedPart]);
  const detailRows = useMemo(() => { if (!report) return []; const search = detailSearch.trim().toLowerCase(); return report.details.filter((row) => { if (selectedPart && row.part !== selectedPart) return false; if (selectedReceipt && row.receipt !== selectedReceipt) return false; return !search || `${row.part} ${row.receipt} ${row.description} ${row.entryNumber} ${row.status}`.toLowerCase().includes(search); }); }, [report, detailSearch, selectedPart, selectedReceipt]);
  const orderRows = useMemo(() => report?.orders.filter((row) => !selectedConsignee || row.consignee === selectedConsignee) ?? [], [report, selectedConsignee]);
  const movementChart = useMemo(() => [...movementRows].sort((a, b) => b.end - a.end).slice(0, 12), [movementRows]);
  const agingPieces = agingRows.reduce((sum, row) => sum + row.pieces, 0);
  const totals = useMemo(() => (report?.movements ?? []).reduce((sum, row) => ({ start: sum.start + row.start, inbound: sum.inbound + row.inbound, outbound: sum.outbound + row.outbound, end: sum.end + row.end }), { start: 0, inbound: 0, outbound: 0, end: 0 }), [report]);
  const averageCycle = useMemo(() => { const values = orderRows.map((row) => row.cycleDays).filter((value): value is number => value !== null); return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0; }, [orderRows]);
  const cycleData = useMemo(() => { const buckets = new Map<string, number>([["0–1 days", 0], ["2 days", 0], ["3 days", 0], ["4+ days", 0]]); orderRows.forEach((row) => { const value = row.cycleDays; const key = value === null ? "" : value <= 1 ? "0–1 days" : value === 2 ? "2 days" : value === 3 ? "3 days" : "4+ days"; if (key) buckets.set(key, (buckets.get(key) ?? 0) + 1); }); return Array.from(buckets, ([name, value]) => ({ name, value })); }, [orderRows]);
  const consigneeData = useMemo(() => { const grouped = new Map<string, number>(); report?.orders.forEach((row) => grouped.set(row.consignee, (grouped.get(row.consignee) ?? 0) + 1)); return Array.from(grouped, ([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value).slice(0, 8); }, [report]);
  const agingBuckets = useMemo(() => { const grouped = new Map<string, number>([["1–2 years", 0], ["2–3 years", 0], ["3+ years", 0]]); agingRows.forEach((row) => grouped.set(row.ageBucket, (grouped.get(row.ageBucket) ?? 0) + row.pieces)); return Array.from(grouped, ([name, value]) => ({ name, value })).filter((item) => item.value > 0); }, [agingRows]);
  const receiptColorMap = useMemo(() => { const receipts = Array.from(new Set(report?.aging.map((row) => row.receipt) ?? [])).sort(); return new Map(receipts.map((receipt, index) => [receipt, RECEIPT_COLORS[index] ?? `hsl(${(210 + index * 47) % 360} 48% 40%)`])); }, [report]);
  const receiptBreakdown = useMemo(() => { const grouped = new Map<string, number>(); agingRowsForReceiptChart.forEach((row) => grouped.set(row.receipt, (grouped.get(row.receipt) ?? 0) + row.pieces)); return Array.from(grouped, ([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value); }, [agingRowsForReceiptChart]);
  const receiptData = receiptBreakdown.slice(0, 10);
  const currentOutbound = report?.outboundVolumes.find((item) => item.year === String(report.reportDate?.getFullYear())) ?? report?.outboundVolumes[0];
  const priorOutbound = report?.outboundVolumes.find((item) => item.year === String((report.reportDate?.getFullYear() ?? 0) - 1));
  const currentMonthIndex = report?.reportDate?.getMonth() ?? 11;
  const latestOutbound = currentOutbound?.values[currentMonthIndex] ?? 0; const priorOutboundValue = priorOutbound?.values[currentMonthIndex] ?? 0;
  const outboundChange = priorOutboundValue ? (latestOutbound - priorOutboundValue) / Math.abs(priorOutboundValue) : 0;
  const totalCo2 = report?.co2.reduce((sum, row) => sum + row.emission, 0) ?? 0;
  const activeLinks = [selectedPart ? ["Part number", selectedPart] : null, selectedReceipt ? ["Warehouse receipt", selectedReceipt] : null, selectedConsignee ? ["Consignee", selectedConsignee] : null].filter(Boolean) as string[][];
  const inboundChart = useMemo(() => volumeChartData(report?.inboundVolumes ?? []), [report]); const outboundChart = useMemo(() => volumeChartData(report?.outboundVolumes ?? []), [report]);
  const pageSize = 20; const pageCount = Math.max(1, Math.ceil(detailRows.length / pageSize)); const safePage = Math.min(detailPage, pageCount); const pagedDetails = detailRows.slice((safePage - 1) * pageSize, safePage * pageSize);
  const observations = useMemo(() => {
    if (!report) return []; const largest = [...movementRows].sort((a, b) => b.end - a.end)[0]; const oldest = [...agingRows].sort((a, b) => b.ageDays - a.ageDays)[0];
    return [largest ? `${largest.part} has the largest ending inventory in view at ${formatNumber(largest.end)} cases.` : "No part-number inventory is in the current view.", `The report contains ${orderRows.length.toLocaleString()} orders with an average creation-to-delivery cycle of ${averageCycle.toFixed(1)} business days.`, oldest ? `${formatNumber(agingPieces)} aged cases are in view; the oldest receipt is ${oldest.receipt} at ${oldest.ageDays.toLocaleString()} days.` : "No inventory older than one year is in the current view.", priorOutboundValue ? `${MONTHS[currentMonthIndex]} outbound volume is ${Math.abs(outboundChange * 100).toFixed(1)}% ${outboundChange >= 0 ? "higher" : "lower"} than the same month last year.` : "A prior-year outbound comparison is not available for the report month."];
  }, [report, movementRows, agingRows, orderRows.length, averageCycle, agingPieces, priorOutboundValue, currentMonthIndex, outboundChange]);
  const prepareInteractiveReport = () => {
    setShowExport(true); setExportStatus("idle"); setExportMessage("");
    if (exportPackage) URL.revokeObjectURL(exportPackage.url);
    setExportPackage(null);
    if (!report) return;
    try {
      const html = buildCustomerInteractiveReportHtml(report);
      const blob = new Blob([html], { type: "text/html;charset=utf-8" });
      const safePeriod = report.period.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "") || "Customer-Report";
      const safeCustomer = report.customer.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "") || "Customer";
      setExportPackage({ url: URL.createObjectURL(blob), name: `WTDC-Interactive-Customer-Report-${safeCustomer}-${safePeriod}.html`, size: blob.size });
    } catch {
      setExportStatus("error");
      setExportMessage("The interactive file could not be prepared. Close this window and try again.");
    }
  };
  const closeInteractiveReport = () => {
    if (exportPackage) URL.revokeObjectURL(exportPackage.url);
    setExportPackage(null); setShowExport(false); setExportStatus("idle"); setExportMessage("");
  };
  const confirmInteractiveDownload = () => {
    setExportStatus("success");
    setExportMessage(`Download started. Look for “${exportPackage?.name ?? "WTDC Interactive Customer Report.html"}” in your browser’s Downloads folder.`);
    setNotice(`Interactive customer report download started for ${report?.period ?? "the loaded period"}. The downloaded file contains the loaded customer data.`);
  };

  return <main>
    <header id="top" className="topbar"><div className="brand"><img src="/wtdc-logo.png" alt="WTDC Worldwide Logistics" /><div><strong>Customer Reporting</strong><span>Inventory · Orders · Service · Sustainability</span></div></div><div className="top-actions"><span className="privacy-chip"><ShieldCheck size={14} /> Local processing</span><button className="button ghost" type="button" disabled={!report} onClick={clearSelections}><RotateCcw size={15} /> Unfilter report</button><button className="button ghost" type="button" disabled={!report} onClick={() => window.print()}><Printer size={15} /> Print</button><button className="button ghost export-button" type="button" disabled={!report} onClick={prepareInteractiveReport}><Download size={15} /> Export interactive report</button></div></header>
    <div className="floating-navigation" aria-label="Report shortcuts"><button type="button" disabled={!report} onClick={clearSelections}><RotateCcw size={16} /> Unfilter report</button><a href="#top"><ArrowUp size={16} /> Back to top</a></div>
    <section className="workspace-head"><div><p className="eyebrow">WTDC Intelligence</p><h1>{report ? report.customer : "Customer performance dashboard"}</h1><p>{report ? `${report.period} · ${report.fileName}` : "Load a customer workbook to connect inventory, orders, warehouse receipts, volume, safety, and sustainability."}</p></div><button className="button clear" type="button" disabled={!report && loadStage === "idle"} onClick={clearReport}><Trash2 size={15} /> Clear report</button></section>
    <div className="report-actions-grid">
    <section className={`drop-zone ${dragging ? "dragging" : ""} ${isLoading ? "loading" : ""}`} onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(event) => { event.preventDefault(); setDragging(false); void handleFile(event.dataTransfer.files[0]); }}><input ref={inputRef} type="file" accept=".xlsx,.xls" onChange={(event) => void handleFile(event.target.files?.[0])} /><div className="upload-icon">{isLoading ? <LoaderCircle className="spinner" size={24} /> : loadStage === "success" ? <Check size={24} /> : <UploadCloud size={24} />}</div><div><strong>{isLoading ? loadStage === "reading" ? "Reading workbook…" : "Connecting report sections…" : report ? "Replace this customer report" : "Drop the customer Excel report here"}</strong><span>{report ? "Uploading another file replaces the current report." : "or click to choose an .xlsx file"}</span></div><button className="button primary" type="button" disabled={isLoading} onClick={() => inputRef.current?.click()}>{report ? "Choose another file" : "Choose Excel file"}</button></section>
    <section className="sharing-guide"><div className="sharing-icon"><Download size={19} /></div><div><strong>Share a report with its data already loaded</strong><p>1. Load and verify the Excel report. 2. Select <b>Export interactive report</b>.</p><small><ShieldCheck size={12} /> The exported file contains customer data. Share only with authorized recipients, and never assume worksheet totals are correct—verify them before sending.</small></div><button className="button primary" type="button" disabled={!report} onClick={prepareInteractiveReport}><Download size={15} /> Export interactive report</button></section>
    </div>
    {loadStage === "error" && <div className="notice error" role="alert"><FileSpreadsheet size={15} /><span>{notice}</span></div>}
    {activeLinks.length > 0 && <section className="link-strip"><div><Filter size={14} /><strong>Linked report view</strong></div><div className="link-chips">{activeLinks.map(([label, value]) => <button type="button" key={label} onClick={() => label === "Part number" ? setSelectedPart("") : label === "Warehouse receipt" ? setSelectedReceipt("") : setSelectedConsignee("")}><span>{label}</span><b>{value}</b><X size={12} /></button>)}</div><button type="button" onClick={clearSelections}>Clear linked filters</button></section>}
    <nav className="section-nav" aria-label="Report sections"><a href="#overview">Overview</a><a href="#workbook-overview">Monthly performance</a><a href="#safety">Safety & CO₂</a><a href="#inventory">Inventory movement</a><a href="#orders">Order KPI</a><a href="#aging">Aging</a><a href="#volumes">Volumes</a><a href="#details">Linked details</a></nav>
    <section id="overview" className="kpi-grid">
      <KpiCard icon={<Boxes size={18} />} label="Ending inventory" target="inventory" value={report ? formatNumber(totals.end) : "—"} detail={report ? `${totals.end - totals.start >= 0 ? "+" : ""}${formatNumber(totals.end - totals.start)} net change` : "Waiting for In Out Summary"} />
      <KpiCard icon={<PackageOpen size={18} />} label="Inbound cases" target="inventory" value={report ? formatNumber(totals.inbound) : "—"} detail={report ? `${formatNumber(report.movements.filter((row) => row.inbound > 0).length)} active part numbers` : "Waiting for movement data"} />
      <KpiCard icon={<Truck size={18} />} label="Outbound cases" target="volumes" value={report ? formatNumber(totals.outbound) : "—"} detail={report && priorOutboundValue ? `${outboundChange >= 0 ? "+" : ""}${(outboundChange * 100).toFixed(1)}% vs prior-year month` : "Waiting for volume comparison"} tone="gold" />
      <KpiCard icon={<ClipboardList size={18} />} label="Orders in KPI sheet" target="orders" value={report ? formatNumber(report.orders.length) : "—"} detail={report ? `${averageCycle.toFixed(1)} avg. business days to delivery` : "Waiting for KPI sheet"} />
      <KpiCard icon={<Warehouse size={18} />} label="Aged inventory" target="aging" value={report ? formatNumber(report.aging.reduce((sum, row) => sum + row.pieces, 0)) : "—"} detail={report ? `${new Set(report.aging.map((row) => row.receipt)).size} warehouse receipts` : "Waiting for aging sheet"} tone="gold" />
      <KpiCard icon={<ShieldCheck size={18} />} label="Days accident free" target="safety" value={report?.accidentDays === null || !report ? "—" : formatNumber(report.accidentDays)} detail={report?.accidentAsOf ? `As of ${report.accidentAsOf}` : "Waiting for safety data"} />
    </section>
    <section className="coverage-grid"><article><span>Report date</span><strong>{report ? formatDate(report.reportDate) : "—"}</strong></article><article><span>Part numbers</span><strong>{report ? formatNumber(report.movements.length) : "—"}</strong></article><article><span>Warehouse receipts</span><strong>{report ? formatNumber(new Set(report.details.map((row) => row.receipt)).size) : "—"}</strong></article></section>

    {report && <section className="section-block" id="workbook-overview"><div className="section-title"><div><p className="eyebrow">Workbook overview</p><h2>Monthly performance, accuracy and shrinkage</h2><p>Values from the Overview worksheet. Blank cells display as —; percentages and zero values are preserved as reported.</p></div></div><div className="panel table-panel"><div className="table-scroll"><table><thead><tr><th>Metric</th>{report.overview.months.map((month, index) => <th key={index}>{month}</th>)}</tr></thead><tbody>{report.overview.rows.filter(row => !HIDDEN_OVERVIEW_ROWS.has(row.label.trim())).map((row, index) => <tr key={index}><td><strong>{row.label}</strong></td>{report.overview.months.map((_, column) => <td key={column}>{row.values[column] == null ? "—" : typeof row.values[column] === "number" ? row.label.includes("%") ? `${(Number(row.values[column]) * 100).toFixed(1)}%` : formatNumber(Number(row.values[column])) : row.values[column]}</td>)}</tr>)}</tbody></table></div></div></section>}
    <section id="safety" className="section-block"><div className="section-title"><div><p className="eyebrow"><ShieldCheck size={13} /> Safety & sustainability</p><h2>Accident-free operations and CO₂ reporting</h2><p>Missing CO₂ inputs are identified clearly instead of being displayed as a misleading zero.</p></div></div><div className="safety-grid"><article className="panel safety-card"><div className="safety-icon"><ShieldCheck /></div><span>Days accident free</span><strong>{report?.accidentDays === null || !report ? "—" : formatNumber(report.accidentDays)}</strong><p>{report?.accidentDays ? `${(report.accidentDays / 365.25).toFixed(1)} years · as of ${report.accidentAsOf || formatDate(report.reportDate)}` : "Waiting for the Days Accident Free sheet"}</p></article><article className="panel sustainability-card"><div className="safety-icon"><Leaf /></div><span>CO₂ emissions reported</span><strong>{report?.co2.length ? `${totalCo2.toFixed(3)} t` : "Not supplied"}</strong><p>{report?.co2.length ? `${report.co2.length} completed transportation movements` : report ? "The CO₂ worksheet contains the calculation structure but no completed movement inputs." : "Waiting for the CO₂ sheet"}</p></article><article className="panel sustainability-note"><Factory size={22} /><div><span>Reporting readiness</span><strong>{report?.co2.length ? "CO₂ data available" : report ? "CO₂ inputs needed" : "Awaiting report"}</strong><p>Weight, distance, and emission factor are required to calculate emissions for each move.</p></div></article></div>{report?.co2.length ? <div className="panel table-panel"><div className="table-scroll"><table><thead><tr><th>Move</th><th>PO</th><th>Consignee</th><th className="numeric">Weight (lb)</th><th className="numeric">Distance (mi)</th><th className="numeric">Emission factor</th><th className="numeric">CO₂ metric tons</th></tr></thead><tbody>{report.co2.map((row, index) => <tr key={`${row.number}-${index}`}><td>{row.number || "—"}</td><td>{row.po || "—"}</td><td>{row.consignee || "—"}</td><td className="numeric">{formatNumber(row.weight)}</td><td className="numeric">{formatNumber(row.distance)}</td><td className="numeric">{row.factor}</td><td className="numeric"><strong>{row.emission.toFixed(4)}</strong></td></tr>)}</tbody></table></div></div> : null}</section>

    <section id="inventory" className="section-block"><div className="section-title"><div><p className="eyebrow"><Boxes size={13} /> Inventory movement</p><h2>Starting and ending inventory by part number</h2><p>Click a part number in the chart or table to connect its warehouse activity and aged inventory.</p></div><label className="search-box"><Search size={14} /><input value={movementSearch} disabled={!report} onChange={(event) => setMovementSearch(event.target.value)} placeholder="Search part number or description" /></label></div>
      <div className="two-column wide"><article className="panel chart-panel"><div className="panel-head"><div><span>Largest ending balances</span><strong>{movementRows.length} part numbers in view</strong></div><span className="click-hint">Select a bar to link reports</span></div><div className="chart-area">{movementChart.length ? <ResponsiveContainer width="100%" height="100%"><ComposedChart data={movementChart} margin={{ top: 12, right: 16, left: 0, bottom: 42 }} onClick={(state) => { if (state.activeLabel !== undefined) { const part = String(state.activeLabel); setSelectedPart((current) => current === part ? "" : part); setDetailPage(1); } }}><CartesianGrid stroke="#E7EBF0" vertical={false} /><XAxis dataKey="part" angle={-35} textAnchor="end" height={58} axisLine={false} tickLine={false} tick={{ fill: "#000", fontSize: 11 }} /><YAxis axisLine={false} tickLine={false} tick={{ fill: "#000", fontSize: 11 }} /><Tooltip formatter={(value, name) => [formatNumber(Number(value)), name]} /><Legend iconType="circle" wrapperStyle={{ fontSize: 12 }} /><Bar dataKey="start" name="Starting" fill={BRAND.navy} radius={[4, 4, 0, 0]} /><Bar dataKey="end" name="Ending" fill={BRAND.gold} radius={[4, 4, 0, 0]} /></ComposedChart></ResponsiveContainer> : <EmptyPanel icon={<Boxes />} text="Load a report to compare starting and ending inventory." />}</div></article>
        <article className="panel movement-flow"><div className="panel-head"><div><span>Period movement</span><strong>Inventory reconciliation</strong></div></div><div className="flow-metrics"><div><small>Start</small><strong>{report ? formatNumber(totals.start) : "—"}</strong></div><ArrowUpRight /><div className="flow-in"><small>In</small><strong>{report ? `+${formatNumber(totals.inbound)}` : "—"}</strong></div><ArrowDownRight /><div className="flow-out"><small>Out</small><strong>{report ? `−${formatNumber(totals.outbound)}` : "—"}</strong></div><ChevronDown /><div className="ending"><small>End</small><strong>{report ? formatNumber(totals.end) : "—"}</strong></div></div><p>{report ? totals.start + totals.inbound - totals.outbound === totals.end ? "Start + In − Out reconciles to the ending balance." : "The movement totals require reconciliation." : "The movement equation will appear after upload."}</p></article></div>
      <div className="panel table-panel"><div className="table-scroll"><table><thead><tr><th>Part number</th><th>Description</th><th className="numeric">Start</th><th className="numeric">In</th><th className="numeric">Out</th><th className="numeric">End</th><th className="numeric">Net change</th></tr></thead><tbody>{movementRows.length ? movementRows.map((row) => <tr key={row.part} className={selectedPart === row.part ? "selected-row" : ""}><td><button className="reference-link" type="button" onClick={() => { setSelectedPart((current) => current === row.part ? "" : row.part); setDetailPage(1); }}>{row.part}</button></td><td>{row.description}</td><td className="numeric">{formatNumber(row.start)}</td><td className="numeric">{formatNumber(row.inbound)}</td><td className="numeric">{formatNumber(row.outbound)}</td><td className="numeric"><strong>{formatNumber(row.end)}</strong></td><td className="numeric">{row.end - row.start >= 0 ? "+" : ""}{formatNumber(row.end - row.start)}</td></tr>) : <tr><td colSpan={7}><EmptyPanel icon={<FileSpreadsheet />} text="Inventory movement rows will appear here." /></td></tr>}</tbody></table></div></div>
    </section>

    <section id="orders" className="section-block"><div className="section-title"><div><p className="eyebrow"><CalendarClock size={13} /> Order KPI</p><h2>Creation-to-delivery service performance</h2><p>Stage intervals are read from the workbook, including negative values. Total cycle counts weekdays from Creation Date through Delivery Date.</p></div>{selectedConsignee && <button className="button clear" type="button" onClick={() => setSelectedConsignee("")}>Clear consignee</button>}</div>
      <div className="three-column"><article className="panel metric-feature"><span>Average cycle</span><strong>{report ? `${averageCycle.toFixed(1)} days` : "—"}</strong><p>{report ? `${orderRows.length} completed orders in view` : "Waiting for KPI data"}</p><div className="mini-rule"><i style={{ width: report ? `${Math.min(100, averageCycle / 7 * 100)}%` : "0%" }} /></div></article><article className="panel chart-panel"><div className="panel-head"><div><span>Cycle distribution</span><strong>Business days</strong></div></div><div className="chart-area short">{report ? <ResponsiveContainer width="100%" height="100%"><BarChart data={cycleData}><CartesianGrid stroke="#E7EBF0" vertical={false} /><XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fill: "#000", fontSize: 11 }} /><YAxis allowDecimals={false} axisLine={false} tickLine={false} tick={{ fill: "#000", fontSize: 11 }} /><Tooltip /><Bar dataKey="value" name="Orders" fill={BRAND.navy} radius={[4, 4, 0, 0]} /></BarChart></ResponsiveContainer> : <EmptyPanel icon={<CalendarClock />} text="Cycle-time distribution will appear here." />}</div></article><article className="panel chart-panel"><div className="panel-head"><div><span>Orders by consignee</span><strong>Click to filter the order table</strong></div></div><div className="chart-area short">{consigneeData.length ? <ResponsiveContainer width="100%" height="100%"><BarChart data={consigneeData} layout="vertical" margin={{ left: 10, right: 10 }}><XAxis type="number" hide /><YAxis type="category" dataKey="name" width={115} axisLine={false} tickLine={false} tickFormatter={(value) => value.length > 18 ? `${value.slice(0, 17)}…` : value} tick={{ fill: "#000", fontSize: 10 }} /><Tooltip /><Bar dataKey="value" name="Orders" onClick={(item) => setSelectedConsignee((current) => current === String(item.payload?.name ?? "") ? "" : String(item.payload?.name ?? ""))} radius={[0, 4, 4, 0]}>{consigneeData.map((item) => <Cell key={item.name} fill={selectedConsignee === item.name ? BRAND.gold : BRAND.navy} fillOpacity={selectedConsignee && selectedConsignee !== item.name ? 0.4 : 1} />)}</Bar></BarChart></ResponsiveContainer> : <EmptyPanel icon={<ClipboardList />} text="Consignee activity will appear here." />}</div></article></div>
      <div className="panel table-panel"><div className="table-scroll"><table><thead><tr><th>Cargo Release</th><th>PO number</th><th>PRO number</th><th>Consignee</th><th>Creation</th><th>Confirmation</th><th className="numeric">Create–confirm</th><th>Pick</th><th className="numeric">Confirm–pick</th><th>Delivery</th><th className="numeric">Pick–delivery</th><th className="numeric">Total cycle</th></tr></thead><tbody>{orderRows.length ? orderRows.map((row) => <tr key={row.number}><td><strong>{row.number}</strong></td><td>{row.po || "—"}</td><td>{row.pro || "—"}</td><td>{row.consignee}</td><td>{formatDate(row.creation)}</td><td>{formatDate(row.confirmation)}</td><td className="numeric">{row.creationToConfirmation ?? "—"}</td><td>{formatDate(row.pick)}</td><td className="numeric">{row.confirmationToPick ?? "—"}</td><td>{formatDate(row.pickup)}</td><td className="numeric">{row.pickToPickup ?? "—"}</td><td className="numeric"><strong>{row.cycleDays ?? "—"}</strong></td></tr>) : <tr><td colSpan={12}><EmptyPanel icon={<ClipboardList />} text="Order KPI rows will appear here." /></td></tr>}</tbody></table></div></div>
    </section>

    <section id="aging" className="section-block"><div className="section-title"><div><p className="eyebrow"><TimerReset size={13} /> Inventory aging</p><h2>Inventory older than one year</h2><p>Part numbers and warehouse receipts are clickable references that connect this section to inventory details.</p></div><label className="search-box"><Search size={14} /><input value={agingSearch} disabled={!report} onChange={(event) => setAgingSearch(event.target.value)} placeholder="Search part or warehouse receipt" /></label></div>
      <div className="two-column"><article className="panel aging-summary"><div className="panel-head"><div><span>Aged cases by age band</span><strong>{formatNumber(agingPieces)} cases in view</strong></div><span className="click-hint">Inner: age band · Outer: warehouse receipt</span></div><div className="donut-layout"><div className="donut"><ResponsiveContainer width="100%" height="100%">{agingBuckets.length ? <PieChart><Pie data={agingBuckets} dataKey="value" nameKey="name" innerRadius="42%" outerRadius="65%" paddingAngle={2} stroke="none">{agingBuckets.map((item, index) => <Cell key={item.name} fill={SERIES[index % SERIES.length]} />)}</Pie><Pie data={receiptBreakdown} dataKey="value" nameKey="name" innerRadius="72%" outerRadius="92%" paddingAngle={1} stroke="#fff" strokeWidth={1} onClick={(item) => { const receipt = String(item.name ?? ""); setSelectedReceipt((current) => current === receipt ? "" : receipt); setDetailPage(1); }}>{receiptBreakdown.map((item) => <Cell key={item.name} fill={receiptColorMap.get(item.name) ?? BRAND.navy} fillOpacity={selectedReceipt && selectedReceipt !== item.name ? 0.25 : 1} className="clickable-segment" />)}</Pie><Tooltip formatter={(value) => `${formatNumber(Number(value))} cases`} /></PieChart> : <PieChart />}</ResponsiveContainer><div><strong>{formatNumber(agingPieces)}</strong><span>Aged cases</span></div></div><div className="donut-legend"><div className="legend-section"><span className="legend-title">Age bands</span><div className="legend-list">{agingBuckets.map((item, index) => <div key={item.name}><span><i style={{ background: SERIES[index % SERIES.length] }} />{item.name}</span><strong>{formatNumber(item.value)}</strong></div>)}</div></div><div className="legend-section"><span className="legend-title">Warehouse receipts</span><div className="receipt-legend">{receiptBreakdown.map((item) => <button key={item.name} type="button" className={selectedReceipt === item.name ? "selected" : ""} onClick={() => { setSelectedReceipt((current) => current === item.name ? "" : item.name); setDetailPage(1); }}><span><i style={{ background: receiptColorMap.get(item.name) ?? BRAND.navy }} />{item.name}</span><strong>{formatNumber(item.value)}</strong></button>)}</div></div></div></div></article><article className="panel chart-panel"><div className="panel-head"><div><span>Cases by warehouse receipt</span><strong>Colors match the outer ring · click to filter</strong></div></div><div className="chart-area">{receiptData.length ? <ResponsiveContainer width="100%" height="100%"><BarChart data={receiptData} layout="vertical" margin={{ left: 8, right: 14 }}><XAxis type="number" axisLine={false} tickLine={false} tick={{ fill: "#000", fontSize: 11 }} /><YAxis type="category" dataKey="name" width={62} axisLine={false} tickLine={false} tick={{ fill: "#000", fontSize: 11 }} /><Tooltip formatter={(value) => `${formatNumber(Number(value))} cases`} /><Bar dataKey="value" name="Aged cases" onClick={(item) => { const receipt = String(item.payload?.name ?? ""); setSelectedReceipt((current) => current === receipt ? "" : receipt); setDetailPage(1); }} radius={[0, 4, 4, 0]}>{receiptData.map((item) => <Cell key={item.name} fill={receiptColorMap.get(item.name) ?? BRAND.navy} fillOpacity={selectedReceipt && selectedReceipt !== item.name ? 0.25 : 1} stroke={selectedReceipt === item.name ? "#000" : "none"} strokeWidth={selectedReceipt === item.name ? 2 : 0} />)}</Bar></BarChart></ResponsiveContainer> : <EmptyPanel icon={<Warehouse />} text="Warehouse receipt aging will appear here." />}</div></article></div>
      <div className="panel table-panel"><p className="aging-table-note">One row per part number. Cases are summed across the matching source rows; receipt dates and ages show the range when multiple receipts are included.</p><div className="table-scroll"><table><thead><tr><th>Status</th><th>Warehouse receipts</th><th>Part number</th><th>Description</th><th>Receipt date range</th><th className="numeric">Age range</th><th className="numeric">Cases</th><th>PO numbers</th></tr></thead><tbody>{groupedAgingRows.length ? groupedAgingRows.map(row => <tr key={row.part} className={selectedPart === row.part ? "selected-row" : ""}><td>{[...row.statuses].join(", ")}</td><td>{[...row.receipts].map(receipt => <button key={receipt} className="reference-link aging-receipt" type="button" onClick={() => { setSelectedReceipt(current => current === receipt ? "" : receipt); setDetailPage(1); }}>{receipt}</button>)}</td><td><button className="reference-link" type="button" onClick={() => { setSelectedPart(current => current === row.part ? "" : row.part); setDetailPage(1); }}>{row.part}</button></td><td>{[...row.descriptions].join(" / ")}</td><td>{formatDate(row.earliest)}{row.earliest.getTime() !== row.latest.getTime() ? ` – ${formatDate(row.latest)}` : ""}</td><td className="numeric">{row.minAge === row.maxAge ? row.maxAge : `${row.minAge}–${row.maxAge}`} days</td><td className="numeric"><strong>{formatNumber(row.pieces)}</strong></td><td>{[...row.pos].join(", ") || "—"}</td></tr>) : <tr><td colSpan={8}><EmptyPanel icon={<TimerReset />} text="Inventory older than one year will appear here." /></td></tr>}</tbody></table></div></div>
    </section>

    <section id="volumes" className="section-block"><div className="section-title"><div><p className="eyebrow"><TrendingUp size={13} /> Historical volumes</p><h2>Inbound and outbound cases by month and year</h2><p>Compare two years at a time, or select additional years. Hover over a month for exact case counts.</p></div></div><div className="volume-panels"><VolumeChart key={`in-${report?.fileName}`} title="Inbound volume" series={report?.inboundVolumes ?? []} /><VolumeChart key={`out-${report?.fileName}`} title="Outbound volume" series={report?.outboundVolumes ?? []} /></div></section>



    <section className="panel executive-summary"><div className="summary-icon"><Sparkles /></div><div><p className="eyebrow">Customer summary</p><h2>What the customer should know</h2><div className="observation-grid">{report ? observations.map((item, index) => <article key={item}><span>{index + 1}</span><p>{item}</p></article>) : [1, 2, 3, 4].map((item) => <article key={item}><span>{item}</span><p>—</p></article>)}</div></div></section>

    <section id="details" className="section-block"><div className="section-title"><div><p className="eyebrow"><Activity size={13} /> Linked detail explorer</p><h2>Warehouse movements connected by part number and receipt</h2><p>Use a part number or warehouse receipt elsewhere in the dashboard, or search this table directly.</p></div><label className="search-box"><Search size={14} /><input value={detailSearch} disabled={!report} onChange={(event) => { setDetailSearch(event.target.value); setDetailPage(1); }} placeholder="Search movement details" /></label></div><div className="panel table-panel detail-table"><div className="table-scroll"><table><thead><tr><th>Status</th><th>Part number</th><th>Consignee</th><th>Warehouse receipt</th><th>Purchase order</th><th>Description</th><th className="numeric">Pieces</th><th>Cargo release</th><th>Entry number</th><th>Entry date</th><th>Out date</th></tr></thead><tbody>{pagedDetails.length ? pagedDetails.map((row, index) => <tr key={`${row.part}-${row.receipt}-${index}`}><td><span className={`status ${row.status.toLowerCase().includes("deliver") ? "delivered" : "onhand"}`}>{row.status}</span></td><td><button className="reference-link" type="button" onClick={() => { setSelectedPart((current) => current === row.part ? "" : row.part); setDetailPage(1); }}>{row.part || "—"}</button></td><td>{row.consignee || "—"}</td><td><button className="reference-link" type="button" onClick={() => { setSelectedReceipt((current) => current === row.receipt ? "" : row.receipt); setDetailPage(1); }}>{row.receipt}</button></td><td>{row.location || "—"}</td><td>{row.description}</td><td className="numeric">{formatNumber(row.pieces)}</td><td>{row.packageType || "—"}</td><td>{row.entryNumber || "—"}</td><td>{formatDate(row.entryDate)}</td><td>{formatDate(row.customsDate)}</td></tr>) : <tr><td colSpan={11}><EmptyPanel icon={<Activity />} text="Detailed warehouse movements will appear here." /></td></tr>}</tbody></table></div><div className="pagination"><span>{report ? `${detailRows.length.toLocaleString()} linked rows · Page ${safePage} of ${pageCount}` : "No detail loaded"}</span><div><button type="button" disabled={safePage <= 1} onClick={() => setDetailPage((value) => Math.max(1, value - 1))} aria-label="Previous page"><ChevronLeft size={15} /></button><button type="button" disabled={safePage >= pageCount} onClick={() => setDetailPage((value) => Math.min(pageCount, value + 1))} aria-label="Next page"><ChevronRight size={15} /></button></div></div></div></section>
    <footer><span><ShieldCheck size={13} /> Workbook data never leaves this browser</span><span>Refreshing or closing the tab clears the report</span><span>WTDC Customer Reporting Dashboard</span></footer>
    {showExport && <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.currentTarget === event.target) closeInteractiveReport(); }}><section className="email-modal export-modal" role="dialog" aria-modal="true" aria-labelledby="export-title"><div className="modal-head"><div><p className="eyebrow">Portable offline dashboard</p><h2 id="export-title">Download interactive customer report</h2><p>The report is prepared locally in this browser. Nothing is uploaded.</p></div><button type="button" onClick={closeInteractiveReport} aria-label="Close interactive report export"><X size={18} /></button></div><div className="export-warning"><ShieldCheck size={20} /><div><strong>Confidential customer data will be embedded in this file</strong><p>The recipient can open and filter the report without the Excel workbook or dashboard link. Share it only with authorized recipients.</p></div></div><div className="export-details"><div><span>Customer</span><strong>{report?.customer ?? "—"}</strong></div><div><span>Reporting period</span><strong>{report?.period ?? "—"}</strong></div><div><span>File type</span><strong>Interactive HTML</strong></div><div><span>Approximate size</span><strong>{exportPackage ? `${Math.max(1, Math.round(exportPackage.size / 1024))} KB` : "Preparing…"}</strong></div></div>{exportStatus !== "idle" && <div className={`export-result ${exportStatus}`}><span>{exportStatus === "success" ? <CheckCircle2 size={18} /> : <X size={18} />}</span><p>{exportMessage}</p></div>}<div className="modal-actions"><button className="button clear" type="button" onClick={closeInteractiveReport}>Cancel</button>{exportPackage ? <a className="button primary" href={exportPackage.url} download={exportPackage.name} onClick={confirmInteractiveDownload}><Download size={15} /> {exportStatus === "success" ? "Download again" : "Download HTML report"}</a> : <button className="button primary" type="button" disabled><LoaderCircle className="spinner" size={15} /> Preparing file</button>}</div><p className="export-help">After downloading, attach the HTML file to your email. If your email service blocks HTML attachments, place the file inside a ZIP file first. The regular dashboard link always opens blank.</p></section></div>}
  </main>;
}
