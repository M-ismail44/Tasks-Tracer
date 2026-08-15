/**
 * تصميم هذا الملف: «دفتر الأداء التنفيذي» — بيانات تشغيلية عربية أولاً، منظمة وقابلة للاستيراد والتصدير.
 */
export type TaskStatus = "Done" | "On Progress" | "Not Started" | "Late" | "Pending";
export type TaskPriority = "HIGH" | "MEDIUM" | "LOW";

export interface Task {
  id: string;
  taskType: string;
  client: string;
  rank: number;
  priority: TaskPriority;
  owner: string;
  startDate: string;
  dueDate: string;
  status: TaskStatus;
  progress: number;
  notes: string;
  link: string;
}

export const STORAGE_KEY = "mIsmailTasks_v1";
export const GROUP_ORDER_KEY = "mIsmailTaskGroupOrder_v1";
export const STAGE_STORAGE_KEY = "mIsmailTaskStages_v1";

export const TASK_TYPE_OPTIONS = [
  "إقرارات ضريبة كسب العمل",
  "إقرارات ضريبة القيمة المضافة",
  "إقرارات ضريبة الخصم من المنبع",
  "تسوية ربع سنوية (ضريبة كسب العمل)",
  "المراجعة الدورية",
  "إعداد إقرارات الدخل والقوائم المالية",
  "إعداد الميزانية",
  "مراجعة العقود",
  "تقديم الاعتراض",
  "إعداد التقارير المالية",
];

const coreClients = [
  "Spade Consulting",
  "Dyne Digital",
  "Strike Ads",
  "Granre Refrigration Limited",
  "Fufeng Packaging Technology",
  "Giza Power Industry",
  "Giza Egyptian For Trans.",
  "G B E Meters",
  "EWB",
  "Belleza",
  "King Steel",
  "Porcelina",
  "Minds And Machines",
  "Podu Media",
  "Calmera",
  "Mide Pride Hospital",
  "Reviera Medical Center",
  "Black Ox",
  "Ashraf Shabaan",
];

const owners = [
  "M.ISMAIL",
  "M.ISMAIL",
  "M.ISMAIL",
  "M.ISMAIL",
  "ABDULAH",
  "M.ISMAIL",
  "ABDUL-RAHMAN",
  "ABDUL-RAHMAN",
  "M.ISMAIL",
  "ABDUL-RAHMAN",
  "ABDUL-RAHMAN",
  "ABDUL-RAHMAN",
  "ABDULAH",
  "ABDULAH",
  "ABDULAH",
  "ABDULAH",
  "ABDUL-RAHMAN",
  "ABDULAH",
  "ABDUL-RAHMAN",
];

const rankForIndex = (index: number) => [5, 5, 4, 3, 2, 3, 4, 5, 6, 7, 8, 9, 2, 4, 3, 5, 4, 3, 2][index] ?? 4;
const priorityForIndex = (index: number): TaskPriority => (index < 13 ? "HIGH" : index < 17 ? "MEDIUM" : "LOW");

function makeTasks(
  type: string,
  count: number,
  dueDate: string,
  idStart: number,
  changes: Record<number, Pick<Task, "status" | "progress" | "notes">> = {},
): Task[] {
  return Array.from({ length: count }, (_, index) => {
    const specialClient = count === 20 && index === 16 ? "M H Saadeldin (GLC)" : coreClients[index % coreClients.length];
    const updated = changes[index] ?? { status: "Not Started" as TaskStatus, progress: 0, notes: "" };
    return {
      id: String(idStart + index),
      taskType: type,
      client: specialClient,
      rank: rankForIndex(index),
      priority: priorityForIndex(index),
      owner: owners[index % owners.length],
      startDate: "2025-04-01",
      dueDate,
      status: updated.status,
      progress: updated.progress,
      notes: updated.notes,
      link: "",
    };
  });
}

/** البيانات المرجعية مولدة من نفس مجموعات العملاء وأنواع العمل في ملف HTML المصدر. */
export const seedTasks: Task[] = [
  ...makeTasks("إقرارات ضريبة كسب العمل", 19, "2025-04-15", 1, {
    0: { status: "On Progress", progress: 25, notes: "" },
    3: { status: "Done", progress: 100, notes: "" },
    13: { status: "Done", progress: 100, notes: "بدون موظفين" },
    16: { status: "Done", progress: 0, notes: "غير مسجل" },
    18: { status: "Done", progress: 0, notes: "غير مسجل" },
  }),
  ...makeTasks("إقرارات ضريبة القيمة المضافة", 20, "2025-04-30", 20, {
    0: { status: "On Progress", progress: 25, notes: "" },
    2: { status: "On Progress", progress: 75, notes: "" },
    3: { status: "On Progress", progress: 75, notes: "" },
    4: { status: "On Progress", progress: 50, notes: "" },
    6: { status: "Done", progress: 100, notes: "" },
    7: { status: "Done", progress: 0, notes: "غير مسجل" },
    8: { status: "On Progress", progress: 75, notes: "غير مسجل" },
    9: { status: "On Progress", progress: 50, notes: "غير مسجل" },
    10: { status: "On Progress", progress: 75, notes: "غير مسجل" },
    13: { status: "On Progress", progress: 75, notes: "" },
    16: { status: "Done", progress: 0, notes: "غير مسجل" },
    17: { status: "Done", progress: 0, notes: "غير مسجل" },
    18: { status: "Done", progress: 0, notes: "غير مسجل" },
    19: { status: "Done", progress: 0, notes: "تيسيرات" },
  }),
  ...makeTasks("المراجعة الدورية", 19, "2025-04-30", 40, {
    0: { status: "On Progress", progress: 20, notes: "" },
    1: { status: "On Progress", progress: 20, notes: "" },
    2: { status: "On Progress", progress: 20, notes: "" },
    3: { status: "On Progress", progress: 20, notes: "" },
    12: { status: "On Progress", progress: 20, notes: "" },
    13: { status: "On Progress", progress: 20, notes: "" },
  }),
  ...makeTasks("إقرارات ضريبة الخصم من المنبع", 19, "2025-04-28", 59),
  ...makeTasks("تسوية ربع سنوية (ضريبة كسب العمل)", 19, "2025-04-28", 78),
  ...makeTasks("إعداد إقرارات الدخل والقوائم المالية", 20, "2025-04-30", 97),
];

export const stageDefinitions: Record<string, { key: string; label: string; weight: number }[]> = {
  "إقرارات ضريبة كسب العمل": [
    { key: "s1", label: "الحصول على شيت الشركة", weight: 25 },
    { key: "s2", label: "عمل التسويات (Gross Up)", weight: 25 },
    { key: "s3", label: "الحفظ على منظومة توحيد", weight: 25 },
    { key: "s4", label: "التقديم والإرسال", weight: 25 },
    { key: "s5", label: "لقطة ملخص", weight: 0 },
  ],
  "إقرارات ضريبة القيمة المضافة": [
    { key: "s1", label: "الحصول على شيت الشركة", weight: 25 },
    { key: "s2", label: "عمل التسويات", weight: 25 },
    { key: "s3", label: "الحفظ على منظومة الضرائب", weight: 25 },
    { key: "s4", label: "التقديم والإرسال", weight: 25 },
    { key: "s5", label: "تحميل نموذج 10", weight: 0 },
  ],
  "إقرارات ضريبة الخصم من المنبع": [
    { key: "s1", label: "الحصول على شيت الشركة", weight: 25 },
    { key: "s2", label: "الدفعات المقدمة", weight: 25 },
    { key: "s3", label: "عمل التسويات", weight: 25 },
    { key: "s4", label: "الحفظ على منظومة الضرائب", weight: 25 },
    { key: "s5", label: "تحميل نموذج 41", weight: 0 },
  ],
  "تسوية ربع سنوية (ضريبة كسب العمل)": [
    { key: "s1", label: "الحصول على إيصالات السداد", weight: 25 },
    { key: "s2", label: "مراجعة الإيصالات", weight: 25 },
    { key: "s3", label: "الحفظ على منظومة توحيد", weight: 25 },
    { key: "s4", label: "التقديم والإرسال", weight: 25 },
  ],
  "المراجعة الدورية": [
    { key: "s1", label: "الحصول على ميزان المراجعة", weight: 20 },
    { key: "s2", label: "الحصول على المصادقات والعقود", weight: 20 },
    { key: "s3", label: "إرسال الملاحظات والتعديلات", weight: 20 },
    { key: "s4", label: "الحصول على ميزان المراجعة المعدل", weight: 20 },
    { key: "s5", label: "إعداد القوائم المبدائية", weight: 20 },
  ],
  "إعداد إقرارات الدخل والقوائم المالية": [
    { key: "s1", label: "الحصول على ميزان المراجعة", weight: 20 },
    { key: "s2", label: "مراجعة وتسوية الحسابات", weight: 20 },
    { key: "s3", label: "إعداد القوائم المالية", weight: 20 },
    { key: "s4", label: "إعداد الإقرار الضريبي", weight: 20 },
    { key: "s5", label: "التقديم والإرسال", weight: 20 },
    { key: "s6", label: "سداد الضريبة المستحقة", weight: 0 },
  ],
};

export const stageOptions = [
  { value: "pending", label: "في الانتظار" },
  { value: "collect", label: "جمع البيانات" },
  { value: "process", label: "معالجة محاسبية" },
  { value: "inprog", label: "قيد التنفيذ" },
  { value: "send", label: "إرسال وتقديم" },
  { value: "review", label: "مراجعة" },
  { value: "edit", label: "تعديلات" },
  { value: "done", label: "تم الإنجاز" },
  { value: "na", label: "غير منطبق" },
];
