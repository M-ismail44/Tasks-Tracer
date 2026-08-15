/**
 * تصميم هذا الملف: «دفتر الأداء التنفيذي» — مزامنة اختيارية، شفافة، ومحصورة ببيانات المستخدم داخل Google Drive.
 */
import type { Task } from "@/lib/task-data";

export interface DriveBackup {
  tasks: Task[];
  groupOrder: string[];
  stageMap: Record<string, string>;
  exportedAt: string;
}

type GoogleTokenResponse = { access_token?: string; error?: string };
type TokenClient = { requestAccessToken: (options: { prompt: string }) => void };
type GsiWindow = Window & { google?: { accounts?: { oauth2?: { initTokenClient: (config: { client_id: string; scope: string; callback: (response: GoogleTokenResponse) => void }) => TokenClient } } } };

const BACKUP_NAME = "M_Ismail_Tasks_Backup.json";
const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";

export function loadGoogleIdentity() {
  if ((window as GsiWindow).google?.accounts?.oauth2) return Promise.resolve();
  return new Promise<void>((resolve, reject) => {
    const existing = document.querySelector<HTMLScriptElement>('script[data-google-identity="true"]');
    if (existing) {
      existing.addEventListener("load", () => resolve(), { once: true });
      existing.addEventListener("error", () => reject(new Error("تعذر تحميل Google Identity.")), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = "https://accounts.google.com/gsi/client";
    script.async = true;
    script.defer = true;
    script.dataset.googleIdentity = "true";
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("تعذر تحميل Google Identity."));
    document.head.appendChild(script);
  });
}

export async function requestDriveToken(clientId: string, previousToken?: string) {
  await loadGoogleIdentity();
  const googleIdentity = (window as GsiWindow).google;
  const oauth = googleIdentity?.accounts?.oauth2;
  if (!oauth) throw new Error("لم تتوفر خدمة Google Identity في المتصفح.");
  return new Promise<string>((resolve, reject) => {
    const tokenClient = oauth.initTokenClient({
      client_id: clientId,
      scope: DRIVE_SCOPE,
      callback: (response: GoogleTokenResponse) => response.access_token ? resolve(response.access_token) : reject(new Error(response.error || "تعذر إتمام تسجيل الدخول.")),
    });
    tokenClient.requestAccessToken({ prompt: previousToken ? "" : "consent" });
  });
}

async function driveFetch(token: string, url: string, init: RequestInit = {}) {
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token}`);
  const response = await fetch(url, { ...init, headers });
  if (response.status === 401) throw new Error("انتهت صلاحية جلسة Google. أعد الاتصال ثم حاول مجدداً.");
  return response;
}

async function findBackupFile(token: string, knownFileId?: string | null) {
  if (knownFileId) return knownFileId;
  const query = encodeURIComponent(`name='${BACKUP_NAME}' and trashed=false`);
  const response = await driveFetch(token, `https://www.googleapis.com/drive/v3/files?q=${query}&fields=files(id,name)`);
  if (!response.ok) throw new Error("تعذر البحث عن ملف النسخة الاحتياطية في Drive.");
  const payload = await response.json() as { files?: { id: string }[] };
  return payload.files?.[0]?.id ?? null;
}

export async function saveDriveBackup(token: string, backup: DriveBackup, knownFileId?: string | null) {
  const fileId = await findBackupFile(token, knownFileId);
  const boundary = `m_ismail_boundary_${Date.now()}`;
  const metadata = JSON.stringify({ name: BACKUP_NAME, mimeType: "application/json" });
  const content = JSON.stringify(backup, null, 2);
  const body = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${metadata}\r\n--${boundary}\r\nContent-Type: application/json\r\n\r\n${content}\r\n--${boundary}--`;
  const endpoint = fileId
    ? `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=multipart`
    : "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart";
  const response = await driveFetch(token, endpoint, {
    method: fileId ? "PATCH" : "POST",
    headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
    body,
  });
  if (!response.ok) throw new Error("تعذر رفع النسخة الاحتياطية إلى Google Drive.");
  return (await response.json() as { id: string }).id;
}

export async function loadDriveBackup(token: string, knownFileId?: string | null) {
  const fileId = await findBackupFile(token, knownFileId);
  if (!fileId) throw new Error("لا توجد نسخة محفوظة على Google Drive بعد.");
  const response = await driveFetch(token, `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`);
  if (!response.ok) throw new Error("تعذر تحميل النسخة الاحتياطية من Google Drive.");
  return { fileId, backup: await response.json() as Partial<DriveBackup> };
}
