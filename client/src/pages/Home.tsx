/**
 * تصميم هذا الملف: «دفتر الأداء التنفيذي» — كحلي قيادي، ذهب دلالي، ومساحة تشغيل RTL مخصصة للقرار السريع.
 */
import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "./v8-features.css";
import "./executive-refinement.css";
import {
  AlertCircle,
  ArrowDownUp,
  BarChart3,
  CalendarDays,
  Cloud,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  CircleDashed,
  Clock3,
  Copy,
  Download,
  ExternalLink,
  FileJson,
  FileSpreadsheet,
  Flag,
  GripVertical,
  Layers3,
  ListTodo,
  Loader2,
  PauseCircle,
  Pencil,
  Plus,
  Search,
  RotateCcw,
  Target,
  Trash2,
  Upload,
  UserRound,
  UsersRound,
  X,
} from "lucide-react";
import { toast } from "sonner";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  PolarAngleAxis,
  PolarGrid,
  Radar,
  RadarChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import {
  GROUP_ORDER_KEY,
  seedTasks,
  stageDefinitions,
  stageOptions,
  STAGE_STORAGE_KEY,
  STORAGE_KEY,
  TASK_TYPE_OPTIONS,
  Task,
  TaskPriority,
  TaskStatus,
} from "@/lib/task-data";
import { loadDriveBackup, requestDriveToken, saveDriveBackup } from "@/lib/drive-utils";

type View = "tasks" | "reports" | "stages";
type Filters = { query: string; status: string; priority: string; owner: string; type: string };
type StageMap = Record<string, string>;
type ImportMode = "replace" | "append";
type ImportKey = "taskType" | "client" | "rank" | "priority" | "owner" | "startDate" | "dueDate" | "progress" | "status" | "notes" | "link";
type ImportPreview = { tasks: Task[]; warnings: string[]; source: string; groupOrder?: string[]; stageMap?: StageMap };

const importHeaders: Record<ImportKey, string[]> = { taskType: ["نوع المهمة", "task type", "tasktype", "نوع", "type"], client: ["العميل", "اسم العميل", "الشركة", "client", "company"], rank: ["الترتيب", "rank"], priority: ["الأولوية", "priority"], owner: ["القائم بالعمل", "المسؤول", "owner", "assignee"], startDate: ["تاريخ البدء", "بداية", "start date", "startdate"], dueDate: ["تاريخ الانتهاء", "تاريخ الاستحقاق", "due date", "duedate", "deadline"], progress: ["الإنجاز%", "نسبة الإنجاز", "الإنجاز", "progress", "completion"], status: ["الحالة", "status"], notes: ["ملاحظات", "notes", "note"], link: ["رابط", "link", "url"] };
const statusAliases: Record<string, TaskStatus> = { "مكتمل": "Done", "تم": "Done", done: "Done", "قيد التنفيذ": "On Progress", "جاري": "On Progress", "on progress": "On Progress", "لم يبدأ": "Not Started", "لم تبدأ": "Not Started", "not started": "Not Started", "متأخر": "Late", late: "Late", "معلق": "Pending", "معلقة": "Pending", pending: "Pending" };
const priorityAliases: Record<string, TaskPriority> = { "عالية": "HIGH", "عالي": "HIGH", high: "HIGH", "متوسطة": "MEDIUM", "متوسط": "MEDIUM", medium: "MEDIUM", "منخفضة": "LOW", "منخفض": "LOW", low: "LOW" };

const blankTask = (): Omit<Task, "id"> => ({
  taskType: "",
  client: "",
  rank: 5,
  priority: "HIGH",
  owner: "",
  startDate: new Date().toISOString().slice(0, 10),
  dueDate: "",
  status: "Not Started",
  progress: 0,
  notes: "",
  link: "",
});

const statusMeta: Record<TaskStatus, { label: string; className: string; Icon: typeof CheckCircle2 }> = {
  Done: { label: "مكتمل", className: "status-done", Icon: CheckCircle2 },
  "On Progress": { label: "قيد التنفيذ", className: "status-progress", Icon: Loader2 },
  "Not Started": { label: "لم يبدأ", className: "status-notstarted", Icon: Clock3 },
  Late: { label: "متأخر", className: "status-late", Icon: AlertCircle },
  Pending: { label: "معلق", className: "status-pending", Icon: PauseCircle },
};

const priorityLabel: Record<TaskPriority, string> = { HIGH: "عالية", MEDIUM: "متوسطة", LOW: "منخفضة" };
const statusOrder: TaskStatus[] = ["Done", "On Progress", "Not Started", "Late", "Pending"];
const chartPalette = ["#2563eb", "#059669", "#d4a017", "#7c3aed", "#0d9488", "#ea580c"];

function safeRead<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

function isOverdue(task: Task) {
  if (task.status === "Done" || !task.dueDate) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return new Date(`${task.dueDate}T00:00:00`) < today;
}

function effectiveStatus(task: Task): TaskStatus {
  return isOverdue(task) || task.status === "Late" ? "Late" : task.status;
}

function dateText(date: string) {
  if (!date) return "—";
  return new Intl.DateTimeFormat("en-GB", { day: "2-digit", month: "short", year: "numeric" }).format(new Date(`${date}T00:00:00`));
}

function cleanTask(input: unknown): Task | null {
  if (!input || typeof input !== "object") return null;
  const raw = input as Partial<Task>;
  if (!raw.taskType || !raw.client || !raw.owner) return null;
  return {
    id: String(raw.id || crypto.randomUUID()),
    taskType: String(raw.taskType),
    client: String(raw.client),
    rank: Math.min(10, Math.max(1, Number(raw.rank) || 5)),
    priority: raw.priority === "LOW" || raw.priority === "MEDIUM" ? raw.priority : "HIGH",
    owner: String(raw.owner),
    startDate: String(raw.startDate || ""),
    dueDate: String(raw.dueDate || ""),
    status: statusOrder.includes(raw.status as TaskStatus) ? (raw.status as TaskStatus) : "Not Started",
    progress: Math.min(100, Math.max(0, Number(raw.progress) || 0)),
    notes: String(raw.notes || ""),
    link: String(raw.link || ""),
  };
}

function parseFlexibleDate(value: unknown) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return `${date.y}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
  }
  const text = String(value).trim();
  let match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
  match = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (match) return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString().slice(0, 10);
}

function normalizeImport(input: Partial<Task>): Task {
  const status = statusAliases[String(input.status ?? "").trim().toLocaleLowerCase()] ?? (statusOrder.includes(input.status as TaskStatus) ? input.status as TaskStatus : "Not Started");
  const priority = priorityAliases[String(input.priority ?? "").trim().toLocaleLowerCase()] ?? (input.priority === "HIGH" || input.priority === "MEDIUM" || input.priority === "LOW" ? input.priority : "MEDIUM");
  return { id: input.id ? String(input.id) : crypto.randomUUID(), taskType: String(input.taskType ?? "").trim() || "غير مصنف", client: String(input.client ?? "").trim() || "—", rank: Math.min(10, Math.max(1, Number(input.rank) || 5)), priority, owner: String(input.owner ?? "").trim() || "No One", startDate: parseFlexibleDate(input.startDate), dueDate: parseFlexibleDate(input.dueDate), status, progress: Math.min(100, Math.max(0, Number(input.progress) || 0)), notes: String(input.notes ?? "").trim(), link: String(input.link ?? "").trim() };
}

function keyFromHeader(header: unknown) {
  const text = String(header ?? "").trim().toLocaleLowerCase();
  return (Object.keys(importHeaders) as ImportKey[]).find((key) => importHeaders[key].some((alias) => alias.toLocaleLowerCase() === text));
}

function tasksFromWorkbook(workbook: XLSX.WorkBook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("الملف لا يحتوي على صفحة بيانات.");
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: true, defval: "" });
  const headerIndex = rows.slice(0, 5).findIndex((row) => row.filter((cell) => keyFromHeader(cell)).length >= 2);
  if (headerIndex < 0) throw new Error("تعذر التعرف على عناوين الأعمدة. استخدم عناوين مثل العميل، نوع المهمة، والحالة.");
  const mapped = new Map<number, ImportKey>();
  rows[headerIndex].forEach((cell, index) => { const key = keyFromHeader(cell); if (key) mapped.set(index, key); });
  const warnings: string[] = [];
  const tasks: Task[] = [];
  rows.slice(headerIndex + 1).forEach((row, offset) => {
    if (!row.some((cell) => String(cell ?? "").trim())) return;
    const record: Partial<Task> = {};
    mapped.forEach((key, index) => { (record as Record<string, unknown>)[key] = row[index]; });
    if (!record.client && !record.taskType) return;
    if (!record.client) warnings.push(`الصف ${headerIndex + offset + 2}: لا يوجد اسم عميل.`);
    tasks.push(normalizeImport(record));
  });
  if (!tasks.length) throw new Error("لم يتم العثور على صفوف مهام صالحة في الملف.");
  return { tasks, warnings };
}

function calcStageProgress(task: Task, map: StageMap) {
  const stages = stageDefinitions[task.taskType];
  if (!stages) return task.progress;
  const weighted = stages.filter((stage) => stage.weight > 0);
  const total = weighted.reduce((sum, stage) => {
    const value = map[`${task.id}_${stage.key}`] ?? "pending";
    const share = value === "done" ? 1 : ["collect", "process", "inprog", "send", "review", "edit"].includes(value) ? 0.5 : 0;
    return sum + stage.weight * share;
  }, 0);
  return Math.min(100, Math.round(total));
}

function StageBadge({ status }: { status: TaskStatus }) {
  const meta = statusMeta[status];
  const Icon = meta.Icon;
  return <span className={`status-badge ${meta.className}`}><Icon size={13} className={status === "On Progress" ? "spin-subtle" : ""} />{meta.label}</span>;
}

function StarRank({ rank }: { rank: number }) {
  const filled = Math.max(1, Math.round(rank / 2));
  return <span className="rank-stars" aria-label={`ترتيب ${rank} من 10`}>{Array.from({ length: 5 }, (_, i) => <span key={i}>{i < filled ? "★" : "☆"}</span>)}</span>;
}

function MetricCard({ icon: Icon, tone, value, label }: { icon: typeof ListTodo; tone: string; value: number; label: string }) {
  return <article className="metric-card">
    <div className={`metric-icon ${tone}`}><Icon size={18} /></div>
    <div><strong>{value}</strong><span>{label}</span></div>
  </article>;
}

function ChartTitle({ accent, icon: Icon, title, subtitle }: { accent: string; icon: typeof BarChart3; title: string; subtitle: string }) {
  return <div className="chart-title" style={{ "--accent": accent } as React.CSSProperties}><span><Icon size={18} /></span><div><h3>{title}</h3><p>{subtitle}</p></div></div>;
}

export default function Home() {
  const [view, setView] = useState<View>("tasks");
  const [tasks, setTasks] = useState<Task[]>(() => safeRead<Task[]>(STORAGE_KEY, seedTasks));
  const [groupOrder, setGroupOrder] = useState<string[]>(() => safeRead<string[]>(GROUP_ORDER_KEY, []));
  const [filters, setFilters] = useState<Filters>({ query: "", status: "", priority: "", owner: "", type: "" });
  const [dialogOpen, setDialogOpen] = useState(false);
  const [editing, setEditing] = useState<Task | null>(null);
  const [form, setForm] = useState<Omit<Task, "id">>(blankTask);
  const [deleteTarget, setDeleteTarget] = useState<Task | null>(null);
  const [resetOpen, setResetOpen] = useState(false);
  const [stageMap, setStageMap] = useState<StageMap>(() => safeRead<StageMap>(STAGE_STORAGE_KEY, {}));
  const [stageType, setStageType] = useState("إقرارات ضريبة كسب العمل");
  const importRef = useRef<HTMLInputElement>(null);
  const [importPreview, setImportPreview] = useState<ImportPreview | null>(null);
  const [importMode, setImportMode] = useState<ImportMode>("replace");
  const [driveOpen, setDriveOpen] = useState(false);
  const [driveClientId, setDriveClientId] = useState(() => localStorage.getItem("mIsmailGoogleClientId") ?? "");
  const [driveToken, setDriveToken] = useState<string | null>(null);
  const [driveFileId, setDriveFileId] = useState(() => localStorage.getItem("mIsmailDriveFileId") ?? "");
  const [driveAutoSync, setDriveAutoSync] = useState(() => localStorage.getItem("mIsmailDriveAutoSync") === "1");
  const [driveBusy, setDriveBusy] = useState(false);
  const [lastDriveSync, setLastDriveSync] = useState(() => localStorage.getItem("mIsmailDriveLastSync") ?? "");

  useEffect(() => { localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks)); }, [tasks]);
  useEffect(() => { localStorage.setItem(GROUP_ORDER_KEY, JSON.stringify(groupOrder)); }, [groupOrder]);
  useEffect(() => { localStorage.setItem(STAGE_STORAGE_KEY, JSON.stringify(stageMap)); }, [stageMap]);
  useEffect(() => { localStorage.setItem("mIsmailDriveAutoSync", driveAutoSync ? "1" : "0"); }, [driveAutoSync]);

  const stats = useMemo(() => {
    const total = tasks.length;
    const done = tasks.filter((task) => task.status === "Done").length;
    const progress = tasks.filter((task) => task.status === "On Progress").length;
    const late = tasks.filter(isOverdue).length;
    const pending = tasks.filter((task) => task.status === "Not Started" || task.status === "Pending").length;
    const high = tasks.filter((task) => task.priority === "HIGH").length;
    const average = total ? Math.round(tasks.reduce((sum, task) => sum + task.progress, 0) / total) : 0;
    return { total, done, progress, late, pending, high, average };
  }, [tasks]);

  const owners = useMemo(() => Array.from(new Set(tasks.map((task) => task.owner).filter(Boolean))).sort(), [tasks]);
  const taskTypes = useMemo(() => Array.from(new Set(tasks.map((task) => task.taskType).filter(Boolean))), [tasks]);
  const orderedTypes = useMemo(() => {
    const retained = groupOrder.filter((type) => taskTypes.includes(type));
    return [...retained, ...taskTypes.filter((type) => !retained.includes(type))];
  }, [groupOrder, taskTypes]);

  useEffect(() => {
    if (JSON.stringify(orderedTypes) !== JSON.stringify(groupOrder)) setGroupOrder(orderedTypes);
  }, [orderedTypes, groupOrder]);

  const filteredGroups = useMemo(() => {
    const query = filters.query.trim().toLocaleLowerCase();
    const filtered = tasks.filter((task) => {
      const matchesQuery = !query || [task.client, task.taskType, task.owner, task.notes].join(" ").toLocaleLowerCase().includes(query);
      return matchesQuery &&
        (!filters.status || effectiveStatus(task) === filters.status) &&
        (!filters.priority || task.priority === filters.priority) &&
        (!filters.owner || task.owner === filters.owner) &&
        (!filters.type || task.taskType === filters.type);
    });
    return orderedTypes.map((type) => ({ type, tasks: filtered.filter((task) => task.taskType === type).sort((a, b) => a.rank - b.rank) })).filter((group) => group.tasks.length);
  }, [filters, orderedTypes, tasks]);

  const chartData = useMemo(() => {
    const status = statusOrder.map((key) => ({ name: statusMeta[key].label, value: tasks.filter((task) => effectiveStatus(task) === key).length, color: key === "Done" ? "#059669" : key === "On Progress" ? "#2563eb" : key === "Late" ? "#dc2626" : key === "Pending" ? "#7c3aed" : "#94a3b8" }));
    const byOwner = owners.map((owner) => {
      const list = tasks.filter((task) => task.owner === owner);
      return { name: owner, الإجمالي: list.length, الإنجاز: list.length ? Math.round(list.reduce((sum, task) => sum + task.progress, 0) / list.length) : 0 };
    });
    const byType = orderedTypes.map((type) => ({ name: type.replace("إقرارات ضريبة ", ""), total: tasks.filter((task) => task.taskType === type).length, done: tasks.filter((task) => task.taskType === type && task.status === "Done").length }));
    const byPriority = ["HIGH", "MEDIUM", "LOW"].map((priority) => ({ name: priorityLabel[priority as TaskPriority], مكتمل: tasks.filter((task) => task.priority === priority && task.status === "Done").length, "قيد التنفيذ": tasks.filter((task) => task.priority === priority && task.status === "On Progress").length, "لم يبدأ": tasks.filter((task) => task.priority === priority && task.status === "Not Started").length }));
    return { status, byOwner, byType, byPriority };
  }, [owners, orderedTypes, tasks]);

  const updateFilter = (key: keyof Filters, value: string) => setFilters((current) => ({ ...current, [key]: value }));

  const openAdd = () => { setEditing(null); setForm(blankTask()); setDialogOpen(true); };
  const openEdit = (task: Task) => { setEditing(task); setForm({ ...task }); setDialogOpen(true); };
  const patchForm = <K extends keyof Omit<Task, "id">>(key: K, value: Omit<Task, "id">[K]) => setForm((current) => ({ ...current, [key]: value }));

  const submitTask = (event: FormEvent) => {
    event.preventDefault();
    if (!form.taskType.trim() || !form.client.trim() || !form.owner.trim()) { toast.error("يرجى إكمال نوع المهمة والعميل والقائم بالعمل."); return; }
    if (form.startDate && form.dueDate && form.dueDate < form.startDate) { toast.error("تاريخ الانتهاء يجب أن يأتي بعد تاريخ البدء."); return; }
    const normalized: Omit<Task, "id"> = { ...form, progress: Number(form.progress), rank: Math.max(1, Math.min(10, Number(form.rank))), status: form.progress >= 100 ? "Done" : form.progress > 0 && form.status === "Not Started" ? "On Progress" : form.status };
    if (editing) {
      setTasks((current) => current.map((task) => task.id === editing.id ? { ...normalized, id: task.id } : task));
      toast.success("تم تحديث المهمة.");
    } else {
      setTasks((current) => [...current, { ...normalized, id: crypto.randomUUID() }]);
      toast.success("تمت إضافة المهمة.");
    }
    setDialogOpen(false);
  };

  const duplicateTask = (task: Task) => {
    setTasks((current) => [...current, { ...task, id: crypto.randomUUID(), client: `${task.client} (نسخة)`, status: "Not Started", progress: 0 }]);
    toast.success("تم إنشاء نسخة من المهمة.");
  };

  const deleteTask = () => {
    if (!deleteTarget) return;
    setTasks((current) => current.filter((task) => task.id !== deleteTarget.id));
    toast.success("تم حذف المهمة.");
    setDeleteTarget(null);
  };

  const moveGroup = (type: string, direction: -1 | 1) => {
    const index = orderedTypes.indexOf(type);
    const target = index + direction;
    if (target < 0 || target >= orderedTypes.length) return;
    const next = [...orderedTypes];
    [next[index], next[target]] = [next[target], next[index]];
    setGroupOrder(next);
  };

  const moveTask = (task: Task, direction: -1 | 1) => {
    const siblingTasks = tasks.filter((item) => item.taskType === task.taskType).sort((a, b) => a.rank - b.rank);
    const index = siblingTasks.findIndex((item) => item.id === task.id);
    const target = index + direction;
    if (target < 0 || target >= siblingTasks.length) return;
    [siblingTasks[index], siblingTasks[target]] = [siblingTasks[target], siblingTasks[index]];
    const updatedRanks = new Map(siblingTasks.map((item, itemIndex) => [item.id, itemIndex + 1]));
    setTasks((current) => current.map((item) => updatedRanks.has(item.id) ? { ...item, rank: updatedRanks.get(item.id)! } : item));
  };

  const download = (content: string, fileName: string, type: string) => {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
  };
  const exportJSON = () => { download(JSON.stringify(tasks, null, 2), `tasks-${new Date().toISOString().slice(0, 10)}.json`, "application/json"); toast.success("تم تصدير ملف JSON."); };
  const exportCSV = () => {
    const headers = ["نوع المهمة", "العميل", "الترتيب", "الأولوية", "القائم بالعمل", "تاريخ البدء", "تاريخ الانتهاء", "الإنجاز%", "الحالة", "ملاحظات", "رابط"];
    const rows = tasks.map((task) => [task.taskType, task.client, task.rank, task.priority, task.owner, task.startDate, task.dueDate, task.progress, task.status, task.notes, task.link].map((value) => `"${String(value ?? "").replaceAll('"', '""')}"`).join(","));
    download(`\uFEFF${[headers.join(","), ...rows].join("\n")}`, `tasks-${new Date().toISOString().slice(0, 10)}.csv`, "text/csv;charset=utf-8");
    toast.success("تم تصدير ملف CSV.");
  };
  const inspectImportFile = async (event: ChangeEvent<HTMLInputElement>) => { const file = event.target.files?.[0]; if (!file) return; try { const extension = file.name.toLocaleLowerCase().split(".").pop(); if (extension === "json") { const payload: unknown = JSON.parse(await file.text()); const raw: unknown[] | null = Array.isArray(payload) ? payload : typeof payload === "object" && payload !== null && Array.isArray((payload as { tasks?: unknown }).tasks) ? (payload as { tasks: unknown[] }).tasks : null; if (!raw) throw new Error("ملف JSON لا يحتوي على قائمة مهام."); setImportPreview({ tasks: raw.map((task: unknown) => normalizeImport(task as Partial<Task>)), warnings: [], source: "JSON" }); } else if (extension === "xlsx" || extension === "xls" || extension === "csv") { const source = extension === "csv" ? await file.text() : await file.arrayBuffer(); const workbook = XLSX.read(source, { type: extension === "csv" ? "string" : "array", cellDates: false }); setImportPreview({ ...tasksFromWorkbook(workbook), source: extension.toLocaleUpperCase() }); } else throw new Error("الصيغة غير مدعومة. استخدم XLSX أو XLS أو CSV أو JSON."); } catch (error) { toast.error(error instanceof Error ? error.message : "تعذر قراءة الملف."); } finally { event.target.value = ""; } };
  const confirmImport = () => { if (!importPreview) return; setTasks((current) => importMode === "replace" ? importPreview.tasks : [...current, ...importPreview.tasks]); if (importMode === "replace") { if (importPreview.groupOrder) setGroupOrder(importPreview.groupOrder); if (importPreview.stageMap) setStageMap(importPreview.stageMap); } toast.success(`تم ${importMode === "replace" ? "استبدال" : "دمج"} ${importPreview.tasks.length} مهمة.`); setImportPreview(null); };
  const saveClientId = () => { if (!/^\d+-[a-z0-9]+\.apps\.googleusercontent\.com$/i.test(driveClientId.trim())) { toast.error("أدخل Google Client ID بصيغة صحيحة."); return false; } localStorage.setItem("mIsmailGoogleClientId", driveClientId.trim()); toast.success("تم حفظ Google Client ID على هذا المتصفح."); return true; };
  const connectDrive = async () => { if (!saveClientId()) return; setDriveBusy(true); try { setDriveToken(await requestDriveToken(driveClientId.trim(), driveToken ?? undefined)); toast.success("تم الاتصال بـ Google Drive."); } catch (error) { toast.error(error instanceof Error ? error.message : "تعذر الاتصال بـ Google Drive."); } finally { setDriveBusy(false); } };
  const persistDriveBackup = async () => { if (!driveToken) { toast.error("اتصل بـ Google Drive أولاً."); return; } setDriveBusy(true); try { const fileId = await saveDriveBackup(driveToken, { tasks, groupOrder, stageMap, exportedAt: new Date().toISOString() }, driveFileId || null); const timestamp = new Date().toISOString(); setDriveFileId(fileId); setLastDriveSync(timestamp); localStorage.setItem("mIsmailDriveFileId", fileId); localStorage.setItem("mIsmailDriveLastSync", timestamp); toast.success("تم حفظ النسخة الاحتياطية على Google Drive."); } catch (error) { toast.error(error instanceof Error ? error.message : "فشل الحفظ على Google Drive."); } finally { setDriveBusy(false); } };
  const restoreDriveBackup = async () => { if (!driveToken) { toast.error("اتصل بـ Google Drive أولاً."); return; } setDriveBusy(true); try { const { fileId, backup } = await loadDriveBackup(driveToken, driveFileId || null); const loaded = Array.isArray(backup.tasks) ? backup.tasks.map((task) => normalizeImport(task)) : []; if (!loaded.length) throw new Error("النسخة المحفوظة لا تحتوي على مهام صالحة."); setDriveFileId(fileId); localStorage.setItem("mIsmailDriveFileId", fileId); setDriveOpen(false); setImportPreview({ tasks: loaded, warnings: [], source: "Google Drive", groupOrder: Array.isArray(backup.groupOrder) ? backup.groupOrder : undefined, stageMap: backup.stageMap }); toast.message("تم تحميل النسخة من Drive. اختر أسلوب تطبيقها."); } catch (error) { toast.error(error instanceof Error ? error.message : "فشل تحميل النسخة من Google Drive."); } finally { setDriveBusy(false); } };
  useEffect(() => { if (!driveAutoSync || !driveToken) return; const timer = window.setTimeout(() => { void persistDriveBackup(); }, 4000); return () => window.clearTimeout(timer); }, [tasks, groupOrder, stageMap, driveAutoSync, driveToken]);

  const resetData = () => {
    setTasks(seedTasks);
    setGroupOrder([]);
    setStageMap({});
    setResetOpen(false);
    toast.success("تمت استعادة بيانات البداية.");
  };

  const updateStage = (task: Task, stageKey: string, value: string) => {
    const next = { ...stageMap, [`${task.id}_${stageKey}`]: value };
    const progress = calcStageProgress(task, next);
    setStageMap(next);
    setTasks((current) => current.map((item) => item.id === task.id ? { ...item, progress, status: progress === 100 ? "Done" : progress > 0 ? "On Progress" : "Not Started" } : item));
  };

  const stageTasks = tasks.filter((task) => task.taskType === stageType);

  return <main dir="rtl" className="app-shell">
    <header className="app-header">
      <div className="header-pattern" />
      <div className="header-inner">
        <div className="brand-block">
          <div className="brand-mark-wrap" aria-label="رمز M. Ismail Tasks"><BrandMark /></div>
          <div>
            <h1><span className="arabic-brand">مركز المتابعة التنفيذي</span><span className="brand-latin"><b>M.</b> ISMAIL <i>/ TASKS</i></span></h1>
            <p>مهام الضرائب والمحاسبة <span>•</span> راجع المتأخر قبل نهاية اليوم</p>
          </div>
        </div>
        <div className="header-actions">
          <Button className="btn-primary" onClick={openAdd}><Plus size={16} /><span>مهمة جديدة</span></Button>
          <Button className="btn-success" onClick={exportJSON}><FileJson size={15} /><span>تصدير JSON</span></Button>
          <Button className="btn-gold" onClick={exportCSV}><FileSpreadsheet size={15} /><span>تصدير CSV</span></Button>
          <Button variant="outline" className="btn-ghost-header" onClick={() => importRef.current?.click()}><Upload size={15} /><span>استيراد ملف</span></Button>
          <Button variant="outline" className="btn-ghost-header" onClick={() => setDriveOpen(true)}><Cloud size={15} /><span>Google Drive</span></Button>
          <input ref={importRef} type="file" className="sr-only" accept="application/json,.json,.csv,.xlsx,.xls" onChange={(event) => { void inspectImportFile(event); }} />
        </div>
      </div>
      <nav className="main-nav" aria-label="التنقل الرئيسي">
        <button className={view === "tasks" ? "active" : ""} onClick={() => setView("tasks")}><ListTodo size={16} />لوحة المهام</button>
        <button className={view === "reports" ? "active" : ""} onClick={() => setView("reports")}><BarChart3 size={16} />التقارير والمؤشرات</button>
        <button className={view === "stages" ? "active" : ""} onClick={() => setView("stages")}><Layers3 size={16} />مراحل الإنجاز</button>
      </nav>
    </header>

    {view === "tasks" && <>
      <section className="metrics-strip" aria-label="ملخص المهام">
        <MetricCard icon={ListTodo} tone="tone-blue" value={stats.total} label="إجمالي المهام" />
        <MetricCard icon={CheckCircle2} tone="tone-green" value={stats.done} label="مكتملة" />
        <MetricCard icon={Loader2} tone="tone-yellow" value={stats.progress} label="قيد التنفيذ" />
        <MetricCard icon={AlertCircle} tone="tone-red" value={stats.late} label="متأخرة" />
        <MetricCard icon={PauseCircle} tone="tone-purple" value={stats.pending} label="لم تبدأ / معلقة" />
      </section>
      <section className="workspace-hero">
        <div><span className="eyebrow"><Target size={14} /> مساحة التشغيل</span><h2>راجع المتأخر قبل نهاية اليوم.</h2><p>صفِّ المهام وحدّث التقدم، ثم اتخذ قرارك من مؤشرات الأداء الحية.</p></div>
        <div className="hero-readout"><strong>{stats.average}%</strong><span>متوسط الإنجاز</span></div>
      </section>
      <section className="filters-bar" aria-label="مرشحات البحث">
        <div className="search-field"><Search size={16} /><Input value={filters.query} onChange={(event) => updateFilter("query", event.target.value)} placeholder="بحث في المهام، العملاء، الملاحظات..." /></div>
        <select value={filters.status} onChange={(event) => updateFilter("status", event.target.value)}><option value="">كل الحالات</option>{statusOrder.map((status) => <option key={status} value={status}>{statusMeta[status].label}</option>)}</select>
        <select value={filters.priority} onChange={(event) => updateFilter("priority", event.target.value)}><option value="">كل الأولويات</option>{Object.entries(priorityLabel).map(([key, label]) => <option key={key} value={key}>{label}</option>)}</select>
        <select value={filters.owner} onChange={(event) => updateFilter("owner", event.target.value)}><option value="">كل المسؤولين</option>{owners.map((owner) => <option key={owner}>{owner}</option>)}</select>
        <select value={filters.type} onChange={(event) => updateFilter("type", event.target.value)}><option value="">كل أنواع المهام</option>{taskTypes.map((type) => <option key={type}>{type}</option>)}</select>
        <Button variant="outline" className="clear-filters" onClick={() => setFilters({ query: "", status: "", priority: "", owner: "", type: "" })}><X size={15} />مسح</Button>
      </section>
      <section className="tasks-area">
        {filteredGroups.length ? <div className="table-frame"><table className="tasks-table"><thead><tr><th aria-label="تحريك" /><th>#</th><th>العميل / الشركة</th><th>الترتيب</th><th>الأولوية</th><th>القائم بالعمل</th><th>تاريخ البدء</th><th>تاريخ الانتهاء</th><th>الإنجاز</th><th>الحالة</th><th>ملاحظات</th><th>إجراءات</th></tr></thead><tbody>{filteredGroups.map((group) => <GroupRows key={group.type} group={group} groupIndex={orderedTypes.indexOf(group.type)} groupCount={orderedTypes.length} onMoveGroup={moveGroup} onEdit={openEdit} onDuplicate={duplicateTask} onDelete={setDeleteTarget} onMoveTask={moveTask} />)}</tbody></table></div> : <EmptyState onReset={() => setFilters({ query: "", status: "", priority: "", owner: "", type: "" })} />}</section>
    </>}

    {view === "reports" && <section className="report-page">
      <div className="report-header"><div><span className="eyebrow"><BarChart3 size={14} /> تقرير حي</span><h2>الأداء والمؤشرات التحليلية</h2><p>يتحدث تلقائياً عند إضافة مهمة أو تعديل تقدمها أو تغيير حالتها.</p></div><div className="report-badge"><CalendarDays size={15} /> آخر تحديث محلي</div></div>
      <div className="report-kpis">
        <KpiCard tone="blue" icon={ListTodo} value={stats.total} label="إجمالي المهام" />
        <KpiCard tone="green" icon={CheckCircle2} value={stats.done} label="مهام مكتملة" />
        <KpiCard tone="gold" icon={Target} value={`${stats.average}%`} label="متوسط الإنجاز" />
        <KpiCard tone="red" icon={AlertCircle} value={stats.late} label="مهام متأخرة" />
        <KpiCard tone="purple" icon={Loader2} value={stats.progress} label="قيد التنفيذ" />
        <KpiCard tone="teal" icon={Flag} value={stats.high} label="أولوية عالية" />
      </div>
      <div className="charts-grid">
        <article className="chart-card"><ChartTitle accent="#2563eb" icon={BarChart3} title="توزيع الحالات" subtitle="Status Distribution" /><div className="chart-wrap donut-wrap"><ResponsiveContainer width="100%" height={260}><PieChart><Pie data={chartData.status} dataKey="value" nameKey="name" innerRadius={70} outerRadius={102} paddingAngle={3}>{chartData.status.map((entry) => <Cell key={entry.name} fill={entry.color} />)}</Pie><Tooltip /><Legend verticalAlign="bottom" iconType="circle" /></PieChart></ResponsiveContainer><div className="donut-label"><strong>{stats.total}</strong><span>مهمة</span></div></div></article>
        <article className="chart-card"><ChartTitle accent="#059669" icon={UsersRound} title="توزيع المهام حسب المسؤول" subtitle="Tasks by Owner" /><div className="chart-wrap"><ResponsiveContainer width="100%" height={280}><BarChart data={chartData.byOwner}><CartesianGrid vertical={false} stroke="#edf1f6" /><XAxis dataKey="name" tick={{ fontSize: 11 }} /><YAxis allowDecimals={false} /><Tooltip /><Legend /><Bar dataKey="الإجمالي" fill="#2563eb" radius={[5, 5, 0, 0]} /><Bar dataKey="الإنجاز" fill="#059669" radius={[5, 5, 0, 0]} /></BarChart></ResponsiveContainer></div></article>
        <article className="chart-card"><ChartTitle accent="#d4a017" icon={Layers3} title="المهام حسب النوع" subtitle="Tasks by Type" /><div className="chart-wrap"><ResponsiveContainer width="100%" height={280}><BarChart layout="vertical" data={chartData.byType} margin={{ right: 8, left: 8 }}><CartesianGrid horizontal={false} stroke="#edf1f6" /><XAxis type="number" allowDecimals={false} /><YAxis type="category" dataKey="name" width={120} tick={{ fontSize: 10 }} /><Tooltip /><Legend /><Bar dataKey="total" name="الإجمالي" fill="#d4a017" radius={[0, 5, 5, 0]} /><Bar dataKey="done" name="مكتمل" fill="#059669" radius={[0, 5, 5, 0]} /></BarChart></ResponsiveContainer></div></article>
        <article className="chart-card"><ChartTitle accent="#7c3aed" icon={Target} title="أداء الفريق" subtitle="Team Completion Rate" /><div className="chart-wrap"><ResponsiveContainer width="100%" height={280}><RadarChart data={chartData.byOwner}><Tooltip /><PolarGrid /><PolarAngleAxis dataKey="name" tick={{ fontSize: 11 }} /><Radar dataKey="الإنجاز" stroke="#7c3aed" fill="#7c3aed" fillOpacity={0.2} /></RadarChart></ResponsiveContainer></div></article>
        <article className="chart-card"><ChartTitle accent="#ea580c" icon={ArrowDownUp} title="الأولوية × الحالة" subtitle="Priority vs. Status" /><div className="chart-wrap"><ResponsiveContainer width="100%" height={280}><BarChart data={chartData.byPriority}><CartesianGrid vertical={false} stroke="#edf1f6" /><XAxis dataKey="name" /><YAxis allowDecimals={false} /><Tooltip /><Legend /><Bar dataKey="مكتمل" stackId="a" fill="#059669" /><Bar dataKey="قيد التنفيذ" stackId="a" fill="#2563eb" /><Bar dataKey="لم يبدأ" stackId="a" fill="#cbd5e1" radius={[5, 5, 0, 0]} /></BarChart></ResponsiveContainer></div></article>
        <article className="chart-card"><ChartTitle accent="#0d9488" icon={Target} title="متوسط الإنجاز حسب النوع" subtitle="Average Task-Type Progress" /><div className="chart-wrap"><ResponsiveContainer width="100%" height={280}><BarChart data={chartData.byType}><CartesianGrid vertical={false} stroke="#edf1f6" /><XAxis dataKey="name" tick={{ fontSize: 10 }} /><YAxis domain={[0, 100]} unit="%" /><Tooltip /><Bar dataKey="done" name="المكتمل" fill="#0d9488" radius={[5, 5, 0, 0]} /></BarChart></ResponsiveContainer></div></article>
      </div>
      <article className="team-table-card"><ChartTitle accent="#1e40af" icon={UsersRound} title="جدول أداء الفريق التفصيلي" subtitle="Detailed Team Performance" /><div className="team-table-scroll"><table className="team-table"><thead><tr><th>المسؤول</th><th>إجمالي</th><th>مكتملة</th><th>قيد التنفيذ</th><th>لم تبدأ</th><th>متأخرة</th><th>نسبة الإنجاز</th><th>مؤشر الأداء</th></tr></thead><tbody>{owners.map((owner, index) => { const list = tasks.filter((task) => task.owner === owner); const average = list.length ? Math.round(list.reduce((sum, task) => sum + task.progress, 0) / list.length) : 0; const late = list.filter(isOverdue).length; return <tr key={owner}><td><strong>{owner}</strong></td><td>{list.length}</td><td className="text-green">{list.filter((task) => task.status === "Done").length}</td><td className="text-blue">{list.filter((task) => task.status === "On Progress").length}</td><td>{list.filter((task) => task.status === "Not Started").length}</td><td className="text-red">{late}</td><td><div className="mini-progress"><span style={{ width: `${average}%`, backgroundColor: chartPalette[index % chartPalette.length] }} /><b>{average}%</b></div></td><td><Performance average={average} /></td></tr>; })}</tbody></table></div></article>
    </section>}

    {view === "stages" && <section className="stages-page">
      <div className="report-header stages-banner"><div><span className="eyebrow"><Layers3 size={14} /> متابعة تفصيلية</span><h2>تقرير مراحل الإنجاز</h2><p>حدّث مرحلة العمل، وسيُعاد احتساب تقدم المهمة وحالتها تلقائياً.</p></div></div>
      <div className="stage-toolbar"><Label htmlFor="stageType">نوع المهمة</Label><select id="stageType" value={stageType} onChange={(event) => setStageType(event.target.value)}>{Object.keys(stageDefinitions).map((type) => <option key={type}>{type}</option>)}</select><span>{stageTasks.length} مهمة</span></div>
      <div className="stage-list">{stageTasks.map((task) => <article className="stage-card" key={task.id}><div className="stage-card-head"><div><span className="stage-client">{task.client}</span><small><UserRound size={12} /> {task.owner}</small></div><StageBadge status={effectiveStatus(task)} /></div><div className="stage-progress"><span style={{ width: `${task.progress}%` }} /><b>{task.progress}%</b></div><div className="stage-fields">{stageDefinitions[stageType].map((stage) => <label key={stage.key}><span>{stage.label}{stage.weight > 0 && <em>{stage.weight}%</em>}</span><select value={stageMap[`${task.id}_${stage.key}`] ?? (task.progress >= 100 ? "done" : task.progress > 0 ? "inprog" : "pending")} onChange={(event) => updateStage(task, stage.key, event.target.value)}>{stageOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}</select></label>)}</div></article>)}</div>
    </section>}

    <footer className="app-footer"><span>تخزين محلي داخل المتصفح</span><button onClick={() => setResetOpen(true)}><RotateCcw size={14} />استعادة بيانات البداية</button></footer>

    <Dialog open={dialogOpen} onOpenChange={setDialogOpen}><DialogContent className="task-dialog" dir="rtl"><DialogHeader><DialogTitle>{editing ? "تعديل المهمة" : "إضافة مهمة جديدة"}</DialogTitle><DialogDescription>أدخل المعلومات التشغيلية ثم احفظ لتحديث الجدول والتقارير.</DialogDescription></DialogHeader><form onSubmit={submitTask} className="task-form"><div className="form-grid"><FormField label="نوع المهمة *"><Input list="task-types" value={form.taskType} onChange={(event) => patchForm("taskType", event.target.value)} placeholder="اختر أو اكتب نوع المهمة" /></FormField><datalist id="task-types">{TASK_TYPE_OPTIONS.map((type) => <option key={type} value={type} />)}</datalist><FormField label="اسم العميل / الشركة *"><Input value={form.client} onChange={(event) => patchForm("client", event.target.value)} /></FormField><FormField label="الأولوية *"><select value={form.priority} onChange={(event) => patchForm("priority", event.target.value as TaskPriority)}>{Object.entries(priorityLabel).map(([key, label]) => <option key={key} value={key}>{label}</option>)}</select></FormField><FormField label="الترتيب"><Input type="number" min="1" max="10" value={form.rank} onChange={(event) => patchForm("rank", Number(event.target.value))} /></FormField><FormField label="القائم بالعمل *"><Input list="owner-options" value={form.owner} onChange={(event) => patchForm("owner", event.target.value)} placeholder="اسم الموظف" /></FormField><datalist id="owner-options">{[...owners, "AHMED", "SARA"].map((owner) => <option key={owner} value={owner} />)}</datalist><FormField label="الحالة"><select value={form.status} onChange={(event) => patchForm("status", event.target.value as TaskStatus)}>{statusOrder.filter((status) => status !== "Late").map((status) => <option key={status} value={status}>{statusMeta[status].label}</option>)}</select></FormField><FormField label="تاريخ البدء"><Input type="date" value={form.startDate} onChange={(event) => patchForm("startDate", event.target.value)} /></FormField><FormField label="تاريخ الانتهاء"><Input type="date" value={form.dueDate} onChange={(event) => patchForm("dueDate", event.target.value)} /></FormField><div className="form-field full"><Label>نسبة الإنجاز: <strong>{form.progress}%</strong></Label><input type="range" min="0" max="100" value={form.progress} onChange={(event) => { const progress = Number(event.target.value); patchForm("progress", progress); if (progress === 100) patchForm("status", "Done"); else if (progress > 0 && form.status === "Not Started") patchForm("status", "On Progress"); }} /></div><div className="form-field full"><Label>ملاحظات</Label><Textarea value={form.notes} onChange={(event) => patchForm("notes", event.target.value)} /></div><div className="form-field full"><Label>رابط مرجعي</Label><Input type="url" dir="ltr" value={form.link} onChange={(event) => patchForm("link", event.target.value)} placeholder="https://" /></div></div><DialogFooter><Button type="button" variant="outline" onClick={() => setDialogOpen(false)}>إلغاء</Button><Button type="submit" className="btn-primary">{editing ? "حفظ التعديل" : "إضافة المهمة"}</Button></DialogFooter></form></DialogContent></Dialog>
    <Dialog open={Boolean(deleteTarget)} onOpenChange={(open) => !open && setDeleteTarget(null)}><DialogContent dir="rtl" className="confirm-dialog"><DialogHeader><DialogTitle>حذف المهمة؟</DialogTitle><DialogDescription>سيُحذف سجل <strong>{deleteTarget?.client}</strong> من لوحة المهام والتقارير. لا يمكن التراجع عن ذلك إلا من خلال استيراد نسخة محفوظة.</DialogDescription></DialogHeader><DialogFooter><Button variant="outline" onClick={() => setDeleteTarget(null)}>إلغاء</Button><Button variant="destructive" onClick={deleteTask}>حذف المهمة</Button></DialogFooter></DialogContent></Dialog>
    <Dialog open={resetOpen} onOpenChange={setResetOpen}><DialogContent dir="rtl" className="confirm-dialog"><DialogHeader><DialogTitle>استعادة بيانات البداية؟</DialogTitle><DialogDescription>سيتم استبدال المهام الحالية والفرز ومراحل العمل ببيانات البداية. صدّر ملف JSON إذا رغبت في الاحتفاظ بنسخة من الوضع الحالي.</DialogDescription></DialogHeader><DialogFooter><Button variant="outline" onClick={() => setResetOpen(false)}>إلغاء</Button><Button variant="destructive" onClick={resetData}>استعادة البيانات</Button></DialogFooter></DialogContent></Dialog>
    <Dialog open={Boolean(importPreview)} onOpenChange={(open) => !open && setImportPreview(null)}><DialogContent dir="rtl" className="confirm-dialog"><DialogHeader><DialogTitle>معاينة الاستيراد</DialogTitle><DialogDescription>تمت قراءة ملف {importPreview?.source}. راجع الملخص وحدد كيفية تطبيق البيانات قبل المتابعة.</DialogDescription></DialogHeader><div className="v8-dialog-section"><div className="import-summary-grid"><div>المهام المكتشفة<strong>{importPreview?.tasks.length ?? 0}</strong></div><div>أنواع المهام<strong>{new Set(importPreview?.tasks.map((task) => task.taskType)).size || 0}</strong></div><div>العملاء<strong>{new Set(importPreview?.tasks.map((task) => task.client)).size || 0}</strong></div><div>المهام الحالية<strong>{tasks.length}</strong></div></div>{importPreview?.warnings.length ? <div className="import-warnings"><strong>ملاحظات الاستيراد:</strong><br />{importPreview.warnings.slice(0, 5).join(" ")}{importPreview.warnings.length > 5 ? " …" : ""}</div> : null}<div className="choice-row"><label><input type="radio" name="importMode" checked={importMode === "replace"} onChange={() => setImportMode("replace")} /><span><b>استبدال البيانات الحالية</b><small>سيصبح الملف المستورد هو قائمة المهام الرئيسية.</small></span></label><label><input type="radio" name="importMode" checked={importMode === "append"} onChange={() => setImportMode("append")} /><span><b>دمج مع البيانات الحالية</b><small>تضاف المهام المستوردة إلى القائمة الحالية.</small></span></label></div></div><DialogFooter><Button variant="outline" onClick={() => setImportPreview(null)}>إلغاء</Button><Button className="btn-primary" onClick={confirmImport}>تأكيد الاستيراد</Button></DialogFooter></DialogContent></Dialog>
    <Dialog open={driveOpen} onOpenChange={setDriveOpen}><DialogContent dir="rtl" className="confirm-dialog"><DialogHeader><DialogTitle>التخزين على Google Drive</DialogTitle><DialogDescription>مزامنة اختيارية مباشرة مع حساب Google الخاص بك. لا تُرسل بيانات المهام إلى خادم التطبيق.</DialogDescription></DialogHeader><div className="v8-dialog-section"><div className={`drive-status ${driveToken ? "connected" : ""}`}><i />{driveToken ? "متصل بـ Google Drive" : "غير متصل بـ Google Drive"}</div><p className="drive-note">أدخل <strong>OAuth Client ID</strong> من Google Cloud Console، ثم أضف نطاق موقع Netlify ضمن <strong>Authorized JavaScript origins</strong>.</p><FormField label="Google Client ID"><Input dir="ltr" value={driveClientId} onChange={(event) => setDriveClientId(event.target.value)} placeholder="xxxxxxxx.apps.googleusercontent.com" /></FormField><div className="drive-actions"><Button variant="outline" onClick={saveClientId}>حفظ Client ID</Button><Button className="btn-primary" onClick={() => { void connectDrive(); }} disabled={driveBusy}>{driveBusy ? "جارٍ الاتصال…" : "اتصال / تسجيل الدخول"}</Button><Button className="btn-success" onClick={() => { void persistDriveBackup(); }} disabled={!driveToken || driveBusy}>حفظ الآن على Drive</Button><Button className="btn-gold" onClick={() => { void restoreDriveBackup(); }} disabled={!driveToken || driveBusy}>تحميل من Drive</Button></div><label className="drive-autosync"><input type="checkbox" checked={driveAutoSync} disabled={!driveToken} onChange={(event) => setDriveAutoSync(event.target.checked)} />مزامنة تلقائية بعد كل تعديل</label>{lastDriveSync && <p className="drive-last-sync">آخر مزامنة: {new Date(lastDriveSync).toLocaleString("ar-EG")}</p>}</div><DialogFooter><Button variant="outline" onClick={() => setDriveOpen(false)}>إغلاق</Button></DialogFooter></DialogContent></Dialog>
  </main>;
}

function GroupRows({ group, groupIndex, groupCount, onMoveGroup, onEdit, onDuplicate, onDelete, onMoveTask }: { group: { type: string; tasks: Task[] }; groupIndex: number; groupCount: number; onMoveGroup: (type: string, direction: -1 | 1) => void; onEdit: (task: Task) => void; onDuplicate: (task: Task) => void; onDelete: (task: Task) => void; onMoveTask: (task: Task, direction: -1 | 1) => void }) {
  return <>{<tr className="group-row"><td colSpan={12}><div><span><Layers3 size={15} />{group.type}<b>{group.tasks.length}</b></span><span className="group-actions"><button disabled={groupIndex === 0} onClick={() => onMoveGroup(group.type, -1)} title="نقل المجموعة للأعلى"><ChevronUp size={15} /></button><button disabled={groupIndex === groupCount - 1} onClick={() => onMoveGroup(group.type, 1)} title="نقل المجموعة للأسفل"><ChevronDown size={15} /></button></span></div></td></tr>}{group.tasks.map((task, index) => { const status = effectiveStatus(task); return <tr key={task.id} className={`${status === "Late" ? "is-overdue" : ""} ${task.status === "Done" ? "is-done" : ""}`}><td><GripVertical size={16} className="drag-icon" /></td><td className="serial">{index + 1}</td><td><div className="client-name">{task.client}{task.link && <a href={task.link} target="_blank" rel="noreferrer" aria-label={`فتح رابط ${task.client}`}><ExternalLink size={13} /></a>}</div></td><td><StarRank rank={task.rank} /><small>#{task.rank}</small></td><td><span className={`priority-badge priority-${task.priority}`}>{priorityLabel[task.priority]}</span></td><td><span className="owner-badge"><UserRound size={13} />{task.owner}</span></td><td className="date-cell">{dateText(task.startDate)}</td><td className={`date-cell ${status === "Late" ? "late-date" : ""}`}>{dateText(task.dueDate)}{status === "Late" && <small><AlertCircle size={11} />متأخر</small>}</td><td><div className="progress-cell"><div><span style={{ width: `${task.progress}%` }} className={task.progress === 100 ? "full" : task.progress < 30 ? "warn" : ""} /></div><b>{task.progress}%</b></div></td><td><StageBadge status={status} /></td><td><span className="notes-cell" title={task.notes}>{task.notes || "—"}</span></td><td><div className="row-actions"><Button size="icon" variant="outline" onClick={() => onEdit(task)} title="تعديل"><Pencil size={14} /></Button><Button size="icon" variant="outline" onClick={() => onDuplicate(task)} title="نسخ"><Copy size={14} /></Button><Button size="icon" variant="outline" onClick={() => onMoveTask(task, -1)} disabled={index === 0} title="نقل لأعلى"><ChevronUp size={14} /></Button><Button size="icon" variant="outline" onClick={() => onMoveTask(task, 1)} disabled={index === group.tasks.length - 1} title="نقل لأسفل"><ChevronDown size={14} /></Button><Button size="icon" variant="outline" className="delete-action" onClick={() => onDelete(task)} title="حذف"><Trash2 size={14} /></Button></div></td></tr>; })}</>;
}

function EmptyState({ onReset }: { onReset: () => void }) { return <div className="empty-state"><CircleDashed size={48} /><h3>لا توجد مهام تطابق معايير البحث.</h3><p>غيّر المرشحات أو أعد ضبطها لرؤية بقية قائمة المهام.</p><Button variant="outline" onClick={onReset}><RotateIcon />مسح المرشحات</Button></div>; }
function RotateIcon() { return <ArrowDownUp size={15} />; }
function KpiCard({ tone, icon: Icon, value, label }: { tone: string; icon: typeof ListTodo; value: number | string; label: string }) { return <article className={`kpi-card ${tone}`}><Icon size={26} /><strong>{value}</strong><span>{label}</span></article>; }
function Performance({ average }: { average: number }) { return <span className={`performance performance-${average >= 80 ? "great" : average >= 50 ? "good" : "follow"}`}>{average >= 80 ? "ممتاز" : average >= 50 ? "جيد" : "يحتاج متابعة"}</span>; }
function FormField({ label, children }: { label: string; children: React.ReactNode }) { return <div className="form-field"><Label>{label}</Label>{children}</div>; }
function BrandMark() { return <svg className="brand-logo" viewBox="0 0 64 64" role="img" aria-label="رمز تنظيم وإنجاز"><path d="M12 17a7 7 0 0 1 7-7h21l12 12v23a7 7 0 0 1-7 7H19a7 7 0 0 1-7-7V17Z" fill="#122b57" stroke="#f0c040" strokeWidth="2.5"/><path d="M39 10v12h13" fill="none" stroke="#f0c040" strokeWidth="2.5" strokeLinejoin="round"/><path d="m20 34 7 7 16-17" fill="none" stroke="#53d5ad" strokeWidth="4.5" strokeLinecap="round" strokeLinejoin="round"/><path d="M19 46h20" stroke="#7695c8" strokeWidth="2.5" strokeLinecap="round"/></svg>; }
