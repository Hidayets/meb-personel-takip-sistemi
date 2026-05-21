import React, { useState, useEffect, useRef, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { 
  Users, LayoutDashboard, FileText, DollarSign, Network, Settings, 
  Search, Bell, LogOut, Plus, Download, Filter, FileSpreadsheet,
  Briefcase, Calendar, GraduationCap, Phone, Mail, MapPin, Building,
  ChevronRight, AlertCircle, CheckCircle2, Clock, BarChart2, PieChart as PieChartIcon,
  Trash2, X, Palette, ArrowRight, RefreshCw
} from 'lucide-react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LabelList
} from 'recharts';
import * as XLSX from 'xlsx';
import './index.css';

// --- WINDOW TYPES for Electron ---
interface UpdateStatus {
  status: 'checking' | 'available' | 'not-available' | 'downloading' | 'downloaded' | 'error' | 'dev-mode';
  message: string;
  version?: string;
  percent?: number;
}

interface BackupResult {
  success: boolean;
  path?: string;
  data?: string;
  error?: string;
  canceled?: boolean;
}

interface PrinterOptions {
  orientation?: 'portrait' | 'landscape';
  filename?: string;
}

interface PrinterResult {
  success: boolean;
  path?: string;
  error?: string;
  canceled?: boolean;
}

interface BackupFileInfo {
  name: string;
  size: number;
  mtime: number;
  path: string;
}

declare global {
  interface Window {
    windowControls?: {
      minimize: () => void;
      maximize: () => void;
      close: () => void;
    };
    updater?: {
      checkForUpdates: () => void;
      installUpdate: () => void;
      getAppVersion: () => Promise<string>;
      onUpdateStatus: (callback: (data: UpdateStatus) => void) => () => void;
    };
    backup?: {
      readAuto: () => Promise<string | null>;
      writeAuto: (data: string) => Promise<BackupResult>;
      writeDaily: (data: string) => Promise<BackupResult>;
      exportToDocuments: (data: string) => Promise<BackupResult>;
      saveAs: (data: string) => Promise<BackupResult>;
      import: () => Promise<BackupResult>;
      openFolder: () => Promise<string>;
      openDocumentsFolder: () => Promise<string>;
      list: () => Promise<BackupFileInfo[]>;
      readFile: (filePath: string) => Promise<BackupResult>;
    };
    printer?: {
      savePDF: (options?: PrinterOptions) => Promise<PrinterResult>;
      print: (options?: PrinterOptions) => Promise<PrinterResult>;
    };
  }
}

// Yazdırma orientation'unu form bazlı dinamik olarak ayarlamak için
// geçici bir <style> tag enjekte eder. Yazdırma/PDF işlemi bittiğinde
// kaldırmak için döndürdüğü cleanup fonksiyonu çağrılmalıdır.
//
// Electron printToPDF() preferCSSPageSize: true ile çağrıldığı için
// bu @page kuralı PDF kağıt boyutunu belirler. Margins burada minimum
// tutulur; form içindeki padding zaten boşluk sağlıyor.
const applyPrintOrientation = (orientation: 'portrait' | 'landscape'): (() => void) => {
  const styleId = '__pts_dynamic_print_style__';
  const existing = document.getElementById(styleId);
  if (existing) existing.remove();

  const style = document.createElement('style');
  style.id = styleId;
  style.innerHTML = `
    @media print {
      @page {
        size: A4 ${orientation};
        margin: 5mm 4mm;
      }
      html, body {
        background: white !important;
        margin: 0 !important;
        padding: 0 !important;
      }
      /* Form A4'e tam sığsın - Tailwind min-width/max-width override */
      .border-2.border-black {
        min-width: 0 !important;
        max-width: 100% !important;
        width: 100% !important;
        margin: 0 !important;
        overflow: visible !important;
        box-shadow: none !important;
      }
      /* Tablo sütun genişliklerini esnek yap, içerik taşmasın */
      .border-2.border-black table {
        table-layout: fixed !important;
        width: 100% !important;
      }
      .border-2.border-black th,
      .border-2.border-black td {
        word-break: break-word !important;
        overflow-wrap: anywhere !important;
        white-space: normal !important;
      }
    }
  `;
  document.head.appendChild(style);

  return () => {
    const el = document.getElementById(styleId);
    if (el) el.remove();
  };
};

// localStorage'da uygulamaya ait tüm anahtarların ortak prefix'i.
// Yedekleme/geri yükleme yardımcıları bu prefix'i kullanır.
const PTS_KEY_PREFIX = 'pts_';
const LAST_DAILY_BACKUP_KEY = '__pts_last_daily_backup__';

// localStorage'daki tüm pts_* anahtarlarını yedek formatında topla
const collectBackupPayload = (): string => {
  const dump: Record<string, string> = {};
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith(PTS_KEY_PREFIX)) continue;
      const val = localStorage.getItem(key);
      if (val !== null) dump[key] = val;
    }
  } catch {}
  return JSON.stringify({
    version: 1,
    appName: 'Okul Takip Sistemi',
    timestamp: new Date().toISOString(),
    data: dump,
  });
};

// Yedek payload'unu localStorage'a uygula (geri yükleme)
const applyBackupPayload = (json: string): { count: number; ok: boolean } => {
  try {
    const parsed = JSON.parse(json);
    const data: Record<string, unknown> =
      parsed && typeof parsed === 'object' && parsed.data && typeof parsed.data === 'object'
        ? parsed.data
        : parsed;
    let count = 0;
    Object.entries(data).forEach(([k, v]) => {
      if (typeof k === 'string' && k.startsWith(PTS_KEY_PREFIX) && typeof v === 'string') {
        localStorage.setItem(k, v);
        count++;
      }
    });
    return { count, ok: count > 0 };
  } catch {
    return { count: 0, ok: false };
  }
};

// localStorage'da herhangi bir pts_* verisi var mı?
const hasLocalData = (): boolean => {
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith(PTS_KEY_PREFIX)) return true;
    }
  } catch {}
  return false;
};

// Uygulama açılmadan ÖNCE çalışacak senkron geri yükleme:
// localStorage boşsa ve disk'te otomatik yedek varsa onu yükle.
// (Async olmasına rağmen React render başlamadan önce await ederiz.)
const ensureAutoRestoreBeforeRender = async (): Promise<{ restored: boolean; count: number }> => {
  if (typeof window === 'undefined' || !window.backup) return { restored: false, count: 0 };
  if (hasLocalData()) return { restored: false, count: 0 };
  try {
    const json = await window.backup.readAuto();
    if (!json) return { restored: false, count: 0 };
    const res = applyBackupPayload(json);
    return { restored: res.ok, count: res.count };
  } catch {
    return { restored: false, count: 0 };
  }
};

// --- TYPES ---

interface SalaryRecord {
  id: string;
  year: number;
  month: number;
  type: 'Maas' | 'EkDers';
  amount: number;
}

interface LeaveRecord {
  id: number;
  type: string;
  subType?: string;
  startDate: string;
  endDate: string;
  duration: number;
  year: number;
  description?: string;
}

interface ScheduleRecord {
  id: number;
  personnelId: number;
  type: 'Ders Programı' | 'Nöbet Çizelgesi';
  term: string;
  fileName: string;
  fileData: string;
  uploadDate: string;
}

interface Personnel {
  id: number;
  name: string;
  degree: number;
  level: number;
  branch: string;
  startDate: string;
  performance: number;
  leaveTotal: number;
  leaveHistory: LeaveRecord[];
  tc: string;
  phone: string;
  email: string;
  address: string;
  personnelNo: string;
  iban: string;
  maritalStatus: string;
  childrenCount: number;
  trainings: string[];
  education: string;
  role: string;
  title: string;
  employmentType: string;
  union?: string; // Sendika
  salaryHistory: SalaryRecord[];
  lastPromotionDate?: string; // Son terfi tarihi (mevcut kademeyi aldığı tarih)
}

interface Announcement {
  id: number;
  title: string;
  type: 'text' | 'link' | 'file';
  content?: string;
  url?: string;
  fileName?: string;
  fileData?: string; // Base64
  fileType?: string;
  date: string;
}

interface Notification {
  id: number;
  title: string;
  message: string;
  type: 'info' | 'warning' | 'success' | 'error';
  read: boolean;
  timestamp: string;
}

interface DocumentRecord {
  id: number;
  documentNumber: string;
  subject: string;
  date: string;
  type: 'Gelen' | 'Giden';
  senderReceiver: string;
  status: 'Bekliyor' | 'İşlemde' | 'Tamamlandı';
  relatedPersonnelId?: number;
}

interface AppSettings {
  username: string;
  password?: string;
  provinceTitle: string;
  districtTitle: string;
  schoolTitle: string;
  principalName: string;
  headVicePrincipals: string[]; // 2 adet
  vicePrincipals: string[]; // Dinamik sayıda
  theme: 'blue' | 'dark' | 'gray';
  primaryColor?: string;
  dashboardLayout?: 'grid' | 'list';
  calendarStyle?: 'modern' | 'classic';
  users?: AppUser[];
  printOrientation?: 'portrait' | 'landscape';
}

type UserRole = 'admin' | 'manager' | 'user';

interface User {
  username: string;
  role: UserRole;
  personnelId?: number; // öğretmen hesabı belirli personele bağlanır
}

interface AppUser {
  username: string;
  password: string;
  role: UserRole;
  personnelId?: number;
}

interface ForeignStudent {
  id: number;
  country: string;
  gender: 'Kız' | 'Erkek';
  count: number;
}

interface GradeClassInfo {
  id: number;
  grade: string; // "5. Sınıf", "6. Sınıf" vb.
  classCount: number;
  studentCount: number;
  femaleCount: number;
  maleCount: number;
}

interface GraduateStudent {
  id: number;
  name: string;
  gender: 'Kız' | 'Erkek';
  tookExam: boolean;
  passed?: boolean;
  schoolName?: string; // Kazandığı okul
}

interface AbsenteeStudent {
  id: number;
  name: string;
  gender: 'Kız' | 'Erkek';
  grade: string;
  absentDays: number;
  reason?: string;
}

interface StudentStatistics {
  id: string; // Eğitim-öğretim yılı (örn: "2025-2026")
  academicYear: string; // "2025-2026 Eğitim-Öğretim Yılı"
  totalClasses: number;
  totalStudents: number;
  femaleStudents: number;
  maleStudents: number;
  gradeClassInfo: GradeClassInfo[];
  foreignStudents: ForeignStudent[];
  graduates: GraduateStudent[];
  absentees: AbsenteeStudent[];
  createdDate: string;
  lastModified: string;
}

// --- CONSTANTS ---

const MONTHS = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];

// ISO tarihini (YYYY-MM-DD) Türkçe formata (GG.AA.YYYY) çevir
const formatDateTR = (dateStr: string | undefined): string => {
  if (!dateStr) return '';
  // Zaten gün.ay.yıl formatındaysa dokunma
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(dateStr)) return dateStr;
  // ISO formatı (YYYY-MM-DD) ise çevir
  const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return `${match[3]}.${match[2]}.${match[1]}`;
  }
  return dateStr;
};
const ROLES = ["Okul Müdürü", "Müdür Başyardımcısı", "Müdür Yardımcısı", "Öğretmen", "Memur", "Teknisyen", "Sürekli İşçi", "Geçici İşçi", "Hizmetli", "Güvenlik"];
const BRANCHES = ["Sınıf Öğretmenliği", "Türkçe", "Matematik", "Fen Bilimleri", "Sosyal Bilgiler", "İngilizce", "Bilişim Teknolojileri", "Okul Öncesi", "Özel Eğitim", "Yok (İdari)"];
const TITLES = ["Aday Öğretmen", "Öğretmen", "Uzman Öğretmen", "Başöğretmen", "VHKİ", "Memur", "Hizmetli", "Teknisyen", "İşçi", "Güvenlik Görevlisi"];
const EDUCATIONS = ["İlkokul", "Ortaokul", "Lise", "Önlisans", "Lisans", "Yüksek Lisans", "Doktora"];
const EMPLOYMENT_TYPES = ["Kadrolu", "Sözleşmeli", "Ücretli", "Geçici Görevlendirme"];
const NON_ACADEMIC_ROLES = ["Sürekli İşçi", "Geçici İşçi", "Teknisyen", "Hizmetli", "Güvenlik"];
const MARITAL_STATUSES = ["Evli", "Bekar", "Dul"];

const DEFAULT_SETTINGS: AppSettings = {
  username: "admin",
  password: "123456",
  provinceTitle: "MAMAK MALMÜDÜRLÜĞÜNE",
  districtTitle: "T.C. ANKARA VALİLİĞİ",
  schoolTitle: "XXX ORTAOKULU MÜDÜRLÜĞÜ",
  principalName: "Okul Müdürü Adı Soyadı",
  headVicePrincipals: ["", ""],
  vicePrincipals: [""],
  theme: 'blue',
  primaryColor: 'blue',
  dashboardLayout: 'grid',
  calendarStyle: 'modern',
  printOrientation: 'landscape',
  users: [
    { username: 'admin', password: '123456', role: 'admin' },
    { username: 'manager', password: '123456', role: 'manager' },
  ]
};

// --- HELPER FUNCTIONS ---

const getPrimaryColorClass = (color: string | undefined, type: 'bg' | 'text' | 'border' | 'ring' | 'bg-light' | 'hover-bg') => {
  const c = color || 'blue';
  const colorMap: Record<string, Record<string, string>> = {
    blue: { bg: 'bg-blue-600', text: 'text-blue-600', border: 'border-blue-600', ring: 'ring-blue-500', 'bg-light': 'bg-blue-50', 'hover-bg': 'hover:bg-blue-700' },
    indigo: { bg: 'bg-indigo-600', text: 'text-indigo-600', border: 'border-indigo-600', ring: 'ring-indigo-500', 'bg-light': 'bg-indigo-50', 'hover-bg': 'hover:bg-indigo-700' },
    emerald: { bg: 'bg-emerald-600', text: 'text-emerald-600', border: 'border-emerald-600', ring: 'ring-emerald-500', 'bg-light': 'bg-emerald-50', 'hover-bg': 'hover:bg-emerald-700' },
    rose: { bg: 'bg-rose-600', text: 'text-rose-600', border: 'border-rose-600', ring: 'ring-rose-500', 'bg-light': 'bg-rose-50', 'hover-bg': 'hover:bg-rose-700' },
    amber: { bg: 'bg-amber-600', text: 'text-amber-600', border: 'border-amber-600', ring: 'ring-amber-500', 'bg-light': 'bg-amber-50', 'hover-bg': 'hover:bg-amber-700' }
  };
  return colorMap[c]?.[type] || colorMap['blue'][type];
};

const calculateAutomaticPromotion = (startDateStr: string, education: string) => {
  if (!startDateStr) return { degree: 9, level: 1 };
  const start = new Date(startDateStr);
  const now = new Date();
  
  let totalYears = now.getFullYear() - start.getFullYear();
  const monthDiff = now.getMonth() - start.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && now.getDate() < start.getDate())) {
    totalYears--;
  }
  
  if (totalYears < 0) totalYears = 0;

  let startDegree = 9;
  let startLevel = 1;

  switch (education) {
    case "Lise":
      startDegree = 13;
      startLevel = 1;
      break;
    case "Önlisans":
      startDegree = 10;
      startLevel = 2;
      break;
    case "Lisans":
      startDegree = 9;
      startLevel = 1;
      break;
    case "Yüksek Lisans":
      startDegree = 9;
      startLevel = 3;
      break;
    case "Doktora":
      startDegree = 8;
      startLevel = 3;
      break;
    default:
      startDegree = 9;
      startLevel = 1;
  }

  let currentLevel = startLevel + totalYears;
  let currentDegree = startDegree;

  while (currentLevel > 3) {
    currentLevel -= 3;
    currentDegree -= 1;
  }

  if (currentDegree < 1) {
    currentDegree = 1;
    // Derece 1 için kademe 4'e kadar çıkabilir
    if (currentLevel > 4) {
      currentLevel = 4;
    }
  }

  return { degree: currentDegree, level: currentLevel };
};

const getSalaryPeriod = () => {
  const now = new Date();
  let startYear = now.getFullYear();
  let startMonth = now.getMonth();
  
  if (now.getDate() < 15) {
    startMonth -= 1;
    if (startMonth < 0) { startMonth = 11; startYear -= 1; }
  }
  
  const start = new Date(startYear, startMonth, 15);
  const end = new Date(startYear, startMonth, 15);
  end.setMonth(end.getMonth() + 1);
  end.setDate(14);
  
  return { start, end };
};

const isAnniversaryInPeriod = (startDateStr: string, pStart: Date, pEnd: Date) => {
  if (!startDateStr) return null;
  const start = new Date(startDateStr);
  const annivThisYear = new Date(pStart.getFullYear(), start.getMonth(), start.getDate());
  const annivNextYear = new Date(pStart.getFullYear() + 1, start.getMonth(), start.getDate());

  if (annivThisYear >= pStart && annivThisYear <= pEnd) return annivThisYear;
  if (annivNextYear >= pStart && annivNextYear <= pEnd) return annivNextYear;
  return null;
};

// --- CONSTANTS ---

const CHART_COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#14b8a6'];

// --- COMPONENTS ---

const WindowControlBtn = ({ onClick, children, hoverColor }: { onClick: () => void, children?: React.ReactNode, hoverColor?: string }) => {
  const [hover, setHover] = useState(false);
  return (
    <button onClick={onClick} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)} className="title-bar-no-drag w-[45px] h-[32px] flex items-center justify-center text-sm text-white border-none outline-none cursor-pointer transition-colors"
      style={{
        background: hover ? (hoverColor || 'rgba(255,255,255,0.1)') : 'transparent',
      }}>{children}</button>
  );
};

const TitleBar = ({ title, printOrientation: _printOrientation }: { title: string; printOrientation?: 'portrait' | 'landscape' }) => {
  // NOT: @page orientation kuralı artık global olarak burada tanımlanmıyor.
  // Bunun yerine PDF/Yazdır butonlarına basıldığında dinamik olarak
  // applyPrintOrientation() ile o forma özel orientation enjekte edilir.
  return (
    <>
      <style>
        {`
        .title-bar-drag { -webkit-app-region: drag; }
        .title-bar-no-drag { -webkit-app-region: no-drag; }
        @media print { 
          .no-print { display: none !important; }
          body { background: white !important; }
        }
      `}
      </style>
      <div className="title-bar-drag no-print h-8 bg-slate-800 flex items-center justify-between pl-4 text-white font-sans text-xs select-none relative z-10">
        <div className="flex items-center gap-2 pointer-events-none">
          <span>🏫 {title}</span>
        </div>
        <div className="title-bar-no-drag flex h-full">
          <WindowControlBtn onClick={() => (window as any).windowControls?.minimize()}>─</WindowControlBtn>
          <WindowControlBtn onClick={() => (window as any).windowControls?.maximize()}>◻</WindowControlBtn>
          <WindowControlBtn onClick={() => (window as any).windowControls?.close()} hoverColor="#e74c3c">✕</WindowControlBtn>
        </div>
      </div>
    </>
  );
};

// --- MAIN APP ---

const App = () => {
  const [activeTab, setActiveTab] = useState<'dashboard' | 'personnel' | 'official' | 'leave' | 'documents' | 'org' | 'settings' | 'schedules' | 'students'>('dashboard');
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [loginUsername, setLoginUsername] = useState('admin');
  const [loginPassword, setLoginPassword] = useState('123456');
  const [personnel, setPersonnel] = useState<Personnel[]>([]);
  const [announcements, setAnnouncements] = useState<Announcement[]>([]);
  const [documents, setDocuments] = useState<DocumentRecord[]>([]);
  const [schedules, setSchedules] = useState<ScheduleRecord[]>([]);
  const [settings, setSettings] = useState<AppSettings>(DEFAULT_SETTINGS);
  const [notifications, setNotifications] = useState<Notification[]>([]);
  const [toasts, setToasts] = useState<{id:number, title:string, message:string, type:string}[]>([]);
  const [search, setSearch] = useState("");
  const [aiLoading, setAiLoading] = useState(false);
  const [aiResult, setAiResult] = useState("");
  const [pdfLoading, setPdfLoading] = useState(false);
  const [showAddModal, setShowAddModal] = useState(false);
  const [showDetailModal, setShowDetailModal] = useState(false);
  const [showNotificationPanel, setShowNotificationPanel] = useState(false);
  const [selectedPersonForDetail, setSelectedPersonForDetail] = useState<Personnel | null>(null);
  const [showAnnounceModal, setShowAnnounceModal] = useState(false);
  const [showPreviewModal, setShowPreviewModal] = useState(false);
  const [showSchedulePreview, setShowSchedulePreview] = useState(false);
  const [previewSchedule, setPreviewSchedule] = useState<ScheduleRecord | null>(null);
  const [unionChanges, setUnionChanges] = useState<{id: number, personnelId: number, type: 'Üye Olma' | 'Ayrılma', unionName: string, date: string}[]>([]);
  const [otherChanges, setOtherChanges] = useState<{id: number, personnelId: number, processType: string, amount: string, description: string, date: string}[]>([]);
  const [archives, setArchives] = useState<{id: string, eduYear: string, archivedAt: string, data: Record<string, string>, autoArchived: boolean, note?: string}[]>([]);
  const [dutyChanges, setDutyChanges] = useState<{id: number, personnelId: number, type: 'Göreve Başlama' | 'Görevden Ayrılma' | 'Aylıksız İzin', date: string, destination?: string, description?: string}[]>([]);
  const [manualPromotions, setManualPromotions] = useState<{id: number, personnelNo: string, name: string, tc?: string, education?: string, title?: string, branch?: string, oldDegree?: number, oldLevel?: number, oldPromotionDate?: string, newDegree: number, newLevel: number, date: string, description: string}[]>([]);
  const [excludedAutoPromotions, setExcludedAutoPromotions] = useState<number[]>([]); // Otomatik terfilerden hariç tutulan personel ID'leri
  const [excludedDutyStarters, setExcludedDutyStarters] = useState<number[]>([]); // Otomatik göreve başlamalardan hariç tutulan personel ID'leri
  // Rapor sınırı aşımları için "Gönderildi" işaretlemeleri.
  // Her kayıt: bir personelin belirli bir yıldaki kesinti günlerinin
  // resmi makama gönderildiğini belirtir. Aynı personel için yıl içinde
  // birden çok gönderim olabilir (her yeni rapor sonrası ayrı kayıt).
  const [reportSentRecords, setReportSentRecords] = useState<{
    id: string;
    personnelId: number;
    year: number;
    sentDays: number;
    sentAt: string;
  }[]>([]);
  const [promotionDateOverrides, setPromotionDateOverrides] = useState<{[key: string]: string}>({}); // Terfi tarihlerini özelleştirmek için
  const [oldPromotionDateOverrides, setOldPromotionDateOverrides] = useState<{[key: string]: string}>({}); // Eski kademe tarihlerini özelleştirmek için
  const [officialForm, setOfficialForm] = useState<'salaryChange' | 'stepPromotion'>('salaryChange');
  const [selectedLeaveFilter, setSelectedLeaveFilter] = useState<string>('Hastalık');
  const [selectedPersonnelIds, setSelectedPersonnelIds] = useState<number[]>([]);
  const [personnelSort, setPersonnelSort] = useState<{ key: 'name' | 'tc' | 'branch' | 'degree' | 'leave'; dir: 'asc' | 'desc' }>({ key: 'name', dir: 'asc' });
  
  // Öğrenci İstatistikleri
  const [studentStats, setStudentStats] = useState<StudentStatistics[]>([]);
  const [selectedAcademicYear, setSelectedAcademicYear] = useState<string>('');
  const [showStudentStatsModal, setShowStudentStatsModal] = useState(false);
  const [studentStatsModalType, setStudentStatsModalType] = useState<'general' | 'gradeClass' | 'foreign' | 'graduate' | 'absentee'>('general');
  const [showNewYearDialog, setShowNewYearDialog] = useState(false);
  const [newYearInput, setNewYearInput] = useState('');
  const [yearSavedIndicator, setYearSavedIndicator] = useState(false);
  
  const [expandedLeavePersonId, setExpandedLeavePersonId] = useState<number | null>(null);
  const [expandedLeaveManageId, setExpandedLeaveManageId] = useState<number | null>(null);
  const [editingLeaveId, setEditingLeaveId] = useState<number | null>(null);
  const [editingLeaveData, setEditingLeaveData] = useState<{type: string; startDate: string; endDate: string; description: string}>({type: '', startDate: '', endDate: '', description: ''});

  // Düzenleme durumu için state
  const [editingPersonId, setEditingPersonId] = useState<number | null>(null);

  const [newPerson, setNewPerson] = useState<Partial<Personnel>>({
    role: ROLES[3],
    education: EDUCATIONS[2],
    employmentType: EMPLOYMENT_TYPES[0],
    maritalStatus: MARITAL_STATUSES[1],
    degree: 9,
    level: 1,
    childrenCount: 0,
    branch: "",
    title: TITLES[1],
    startDate: new Date().toISOString().split('T')[0]
  });
  const [newAnnounce, setNewAnnounce] = useState<Partial<Announcement>>({ type: 'text' });
  const [newLeave, setNewLeave] = useState<{personnelId: number | '', type: string, startDate: string, endDate: string, description: string}>({
    personnelId: '', type: 'Yıllık', startDate: '', endDate: '', description: ''
  });

  // Auto Updater State
  const [updateStatus, setUpdateStatus] = useState<UpdateStatus | null>(null);
  const [appVersion, setAppVersion] = useState<string>('1.0.0');
  const [isCheckingUpdate, setIsCheckingUpdate] = useState(false);

  // Auto Updater Effect
  useEffect(() => {
    if (window.updater) {
      // Uygulama versiyonunu al
      window.updater.getAppVersion().then(version => {
        setAppVersion(version);
      }).catch(() => {
        setAppVersion('1.0.0');
      });

      // Güncelleme durumu dinleyicisi
      const unsubscribe = window.updater.onUpdateStatus((data) => {
        setUpdateStatus(data);
        setIsCheckingUpdate(data.status === 'checking');
        
        // Toast bildirimleri
        if (data.status === 'available') {
          addToast('Güncelleme Mevcut', data.message, 'info');
        } else if (data.status === 'downloaded') {
          addToast('Güncelleme Hazır', data.message, 'success');
        } else if (data.status === 'error') {
          addToast('Güncelleme Hatası', data.message, 'error');
        }
      });

      return () => unsubscribe();
    }
  }, []);

  useEffect(() => {
    const s = localStorage.getItem('pts_settings');
    if (s) {
      try {
        const parsed = JSON.parse(s);
        // Eski sürüm uyumluluğu: users yoksa varsayılan admin/manager ile oluştur
        if (!parsed.users || !Array.isArray(parsed.users) || parsed.users.length === 0) {
          parsed.users = [
            { username: parsed.username || 'admin', password: parsed.password || '123456', role: 'admin' },
            { username: 'manager', password: '123456', role: 'manager' },
          ];
        }
        setSettings(parsed);
      } catch {
        // ignore
      }
    }
    const p = localStorage.getItem('pts_personnel'); if (p) setPersonnel(JSON.parse(p));
    const a = localStorage.getItem('pts_announcements'); if (a) setAnnouncements(JSON.parse(a));
    const n = localStorage.getItem('pts_notifs'); if (n) setNotifications(JSON.parse(n));
    const d = localStorage.getItem('pts_documents'); if (d) setDocuments(JSON.parse(d));
    const sch = localStorage.getItem('pts_schedules'); if (sch) setSchedules(JSON.parse(sch));
    const u = localStorage.getItem('pts_union_changes'); if (u) setUnionChanges(JSON.parse(u));
    const oc = localStorage.getItem('pts_other_changes'); if (oc) setOtherChanges(JSON.parse(oc));
    const arc = localStorage.getItem('pts_archives'); if (arc) { try { setArchives(JSON.parse(arc)); } catch {} }
    const dc = localStorage.getItem('pts_duty_changes'); if (dc) setDutyChanges(JSON.parse(dc));
    const pdo = localStorage.getItem('pts_promotion_date_overrides'); if (pdo) setPromotionDateOverrides(JSON.parse(pdo));
    const opdo = localStorage.getItem('pts_old_promotion_date_overrides'); if (opdo) setOldPromotionDateOverrides(JSON.parse(opdo));
    const mp = localStorage.getItem('pts_manual_promotions');
    if (mp) {
      try {
        const parsed = JSON.parse(mp);
        const migrated = (Array.isArray(parsed) ? parsed : []).map((r: any) => {
          const oldDegree = r.oldDegree ?? r.degree ?? '';
          const oldLevel = r.oldLevel ?? r.level ?? '';
          const newDegree = r.newDegree ?? r.degree ?? oldDegree ?? '';
          const newLevel = r.newLevel ?? r.level ?? oldLevel ?? '';
          return { ...r, oldDegree, oldLevel, newDegree, newLevel };
        });
        setManualPromotions(migrated);
      } catch {}
    }
    const eap = localStorage.getItem('pts_excluded_auto_promotions'); if (eap) setExcludedAutoPromotions(JSON.parse(eap));
    const eds = localStorage.getItem('pts_excluded_duty_starters'); if (eds) setExcludedDutyStarters(JSON.parse(eds));
    const ss = localStorage.getItem('pts_student_stats'); if (ss) setStudentStats(JSON.parse(ss));
    const say = localStorage.getItem('pts_selected_academic_year'); if (say) setSelectedAcademicYear(say);
    const rsr = localStorage.getItem('pts_report_sent_records');
    if (rsr) {
      try { setReportSentRecords(JSON.parse(rsr)); } catch {}
    }
  }, []);

  // Öğrenci istatistiklerini localStorage'a kaydet
  useEffect(() => {
    localStorage.setItem('pts_student_stats', JSON.stringify(studentStats));
  }, [studentStats]);

  // Otomatik Eğitim-Öğretim Yılı Sonu Arşivleme
  // 15 Haziran ve sonrasında, mevcut eduYear için otomatik arşiv yoksa snapshot al.
  // (Veriler silinmez, yalnızca arşiv kopyası alınır.)
  useEffect(() => {
    const today = new Date();
    const month = today.getMonth(); // 0 = Ocak
    const day = today.getDate();
    // 15 Haziran (5,15) ile 31 Ağustos arasında veya yıl sonuna yakın tarihlerde tetikle
    const isYearEndWindow = (month === 5 && day >= 15) || month === 6 || month === 7;
    if (!isYearEndWindow) return;
    // Personel veya başka veri varsa anlamlı (boş bir arşiv almayalım)
    if (personnel.length === 0) return;
    const eduYear = getEduYear();
    const stored = JSON.parse(localStorage.getItem('pts_archives') || '[]') as typeof archives;
    const alreadyAuto = stored.some(a => a.eduYear === eduYear && a.autoArchived);
    if (alreadyAuto) return;
    const ok = archiveCurrentEduYear(true, 'Otomatik yıl sonu arşivi');
    if (ok) {
      addToast('Yıl Sonu Arşivi', `${eduYear} eğitim-öğretim yılı için otomatik arşiv oluşturuldu.`, 'success');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [personnel.length]);

  useEffect(() => {
    localStorage.setItem('pts_selected_academic_year', selectedAcademicYear);
  }, [selectedAcademicYear]);

  // --- OTOMATİK YEDEKLEME ---
  // Veriler güncelleme/yeniden yükleme sonrası silinmesin diye, her veri
  // değişiminden 2 sn sonra userData/backups klasörüne JSON yedek alınır.
  // userData klasörü electron-updater tarafından silinmediği için yedek
  // güncellemeden sonra da korunur. Bootstrap aşamasında localStorage boşsa
  // bu dosyadan otomatik geri yükleme yapılır.
  const autoBackupTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const lastBackupPayloadRef = useRef<string>('');

  useEffect(() => {
    if (typeof window === 'undefined' || !window.backup) return;
    if (autoBackupTimerRef.current) clearTimeout(autoBackupTimerRef.current);
    autoBackupTimerRef.current = setTimeout(() => {
      const payload = collectBackupPayload();
      // Aynı içerik tekrar yazılmasın
      if (payload === lastBackupPayloadRef.current) return;
      lastBackupPayloadRef.current = payload;

      window.backup!.writeAuto(payload).catch(() => {});

      // Günde 1 kez günlük rotasyon yedeği (max 7 gün tutulur)
      try {
        const today = new Date().toISOString().slice(0, 10);
        const lastDaily = localStorage.getItem(LAST_DAILY_BACKUP_KEY);
        if (lastDaily !== today) {
          window.backup!.writeDaily(payload).catch(() => {});
          localStorage.setItem(LAST_DAILY_BACKUP_KEY, today);
        }
      } catch {}
    }, 2000);

    return () => {
      if (autoBackupTimerRef.current) clearTimeout(autoBackupTimerRef.current);
    };
  }, [
    personnel,
    settings,
    announcements,
    notifications,
    documents,
    schedules,
    unionChanges,
    otherChanges,
    archives,
    dutyChanges,
    promotionDateOverrides,
    oldPromotionDateOverrides,
    manualPromotions,
    excludedAutoPromotions,
    excludedDutyStarters,
    studentStats,
    selectedAcademicYear,
    reportSentRecords,
  ]);

  // Bootstrap sırasında yedekten geri yükleme yapıldıysa kullanıcıya bildir
  useEffect(() => {
    try {
      const flag = sessionStorage.getItem('__pts_just_restored__');
      if (flag) {
        sessionStorage.removeItem('__pts_just_restored__');
        const count = parseInt(flag, 10) || 0;
        setTimeout(() => {
          addToast(
            'Veriler Geri Yüklendi',
            `Otomatik yedekten ${count} veri kümesi başarıyla geri yüklendi. Verileriniz korundu.`,
            'success'
          );
        }, 800);
      }
    } catch {}
    // Yalnızca ilk açılışta
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // personnelFiltered ve personnelSorted, visiblePersonnelBase tanımlandıktan sonra geliyor (aşağıda)

  // Otomatik Derece/Kademe Tetikleyici
  useEffect(() => {
    // Sadece yeni kayıt eklerken veya düzenleme modunda tarih değiştiğinde otomatik hesapla
    // Ama eğer düzenleme modunda zaten bir derece/kademe varsa ve manuel değiştirildiyse ezmemek için kontrol edilebilir
    // Burada basitlik adına her tarih değişiminde hesaplatıyoruz (kullanıcı manuel düzeltebilir)
    if (newPerson.startDate && newPerson.education && showAddModal && !editingPersonId) {
      const { degree, level } = calculateAutomaticPromotion(newPerson.startDate, newPerson.education);
      setNewPerson(prev => ({ ...prev, degree, level }));
    }
  }, [newPerson.startDate, newPerson.education, showAddModal, editingPersonId]);

  const addToast = (title: string, message: string, type: 'info'|'success'|'warning'|'error') => {
    const id = Date.now();
    setToasts(prev => [...prev, { id, title, message, type }]);
    setTimeout(() => setToasts(curr => curr.filter(t => t.id !== id)), 4000);
  };

  const removeToast = (id: number) => {
    setToasts(curr => curr.filter(t => t.id !== id));
  };

  // Rapor sınırı aşımı için "Gönderildi" işaretleme.
  // Mevcut gönderilmemiş kesinti günlerini kalıcı olarak gönderildi olarak
  // kaydeder. Bundan sonra alınan rapor günleri 7 günlük muafiyet zaten
  // dolduğu için doğrudan resmi çizelgeye eklenir.
  const markReportAsSent = (personnelId: number, days: number, personName: string) => {
    if (days <= 0) return;
    const currentYear = new Date().getFullYear();
    if (!window.confirm(
      `${personName} için ${days} günlük rapor kesintisi "Gönderildi" olarak işaretlenecek.\n\n` +
      `Bu işlemden sonra resmi çizelgeden bu personel kaldırılacaktır. ` +
      `Yeni rapor alındığında, 7 günlük muafiyet beklenmeden, alınan tüm günler ` +
      `doğrudan resmi çizelgeye işlenecektir.\n\nDevam edilsin mi?`
    )) return;

    const newRecord = {
      id: `${personnelId}-${currentYear}-${Date.now()}`,
      personnelId,
      year: currentYear,
      sentDays: days,
      sentAt: new Date().toISOString(),
    };
    const updated = [...reportSentRecords, newRecord];
    setReportSentRecords(updated);
    localStorage.setItem('pts_report_sent_records', JSON.stringify(updated));
    addToast(
      'Gönderildi',
      `${personName} için ${days} gün rapor kesintisi gönderildi olarak işaretlendi.`,
      'success'
    );
  };

  // Yanlışlıkla gönderildi işaretini geri al (en son kaydı sil)
  const undoReportSent = (personnelId: number, personName: string) => {
    const currentYear = new Date().getFullYear();
    const personRecords = reportSentRecords
      .filter(r => r.personnelId === personnelId && r.year === currentYear)
      .sort((a, b) => new Date(b.sentAt).getTime() - new Date(a.sentAt).getTime());
    if (personRecords.length === 0) {
      addToast('Bilgi', 'Geri alınacak gönderim kaydı bulunamadı.', 'info');
      return;
    }
    const latest = personRecords[0];
    if (!window.confirm(
      `${personName} için en son yapılan ${latest.sentDays} günlük "Gönderildi" işareti geri alınacak. Devam edilsin mi?`
    )) return;
    const updated = reportSentRecords.filter(r => r.id !== latest.id);
    setReportSentRecords(updated);
    localStorage.setItem('pts_report_sent_records', JSON.stringify(updated));
    addToast('Geri Alındı', `${personName} için son gönderim işareti geri alındı.`, 'info');
  };

  // --- Eğitim-Öğretim Yılı Arşiv Yardımcıları ---
  // Türkiye eğitim-öğretim yılı: Eylül başında başlar, Haziran sonunda biter.
  // Eylül - Aralık ise: yıl_X-yıl_X+1 (örn 2025-2026)
  // Ocak - Ağustos ise: yıl_X-1-yıl_X (örn 2024-2025)
  const getEduYear = (date: Date = new Date()): string => {
    const y = date.getFullYear();
    const m = date.getMonth(); // 0-11
    if (m >= 8) return `${y}-${y + 1}`; // Eylül (8) ve sonrası
    return `${y - 1}-${y}`;
  };

  const collectAllPtsData = (): Record<string, string> => {
    const data: Record<string, string> = {};
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith('pts_')) continue;
      if (key === 'pts_archives') continue; // arşivlerin kendisi arşive girmez
      const v = localStorage.getItem(key);
      if (v !== null) data[key] = v;
    }
    return data;
  };

  const archiveCurrentEduYear = (auto: boolean = false, note?: string): boolean => {
    const eduYear = getEduYear();
    const current = JSON.parse(localStorage.getItem('pts_archives') || '[]') as typeof archives;
    if (auto && current.some(a => a.eduYear === eduYear && a.autoArchived)) {
      return false; // bu yıl için otomatik arşiv zaten alınmış
    }
    const data = collectAllPtsData();
    const entry = {
      id: `${eduYear}-${Date.now()}`,
      eduYear,
      archivedAt: new Date().toISOString(),
      data,
      autoArchived: auto,
      note
    };
    const updated = [entry, ...current];
    setArchives(updated);
    localStorage.setItem('pts_archives', JSON.stringify(updated));
    return true;
  };

  // Yeni eğitim yılına temiz başlama: dönemsel verileri sıfırlar, kalıcı verileri korur
  const clearPeriodicData = () => {
    // Dönemsel anahtarları sıfırla
    const keysToReset = [
      'pts_duty_changes',
      'pts_union_changes',
      'pts_other_changes',
      'pts_manual_promotions',
      'pts_promotion_date_overrides',
      'pts_old_promotion_date_overrides',
      'pts_excluded_auto_promotions',
      'pts_excluded_duty_starters',
      'pts_notifs',
      'pts_announcements',
    ];
    keysToReset.forEach(k => localStorage.removeItem(k));
    // State'leri de sıfırla
    setDutyChanges([]);
    setUnionChanges([]);
    setOtherChanges([]);
    setManualPromotions([]);
    setPromotionDateOverrides({});
    setOldPromotionDateOverrides({});
    setExcludedAutoPromotions([]);
    setExcludedDutyStarters([]);
    setNotifications([]);
    setAnnouncements([]);
    // Personel listesinde leaveHistory'yi temizle (personel kaydı kalsın, izinler sıfırlansın)
    const cleanedPersonnel = personnel.map(p => ({ ...p, leaveHistory: [], leaveTotal: 0 }));
    setPersonnel(cleanedPersonnel);
    localStorage.setItem('pts_personnel', JSON.stringify(cleanedPersonnel));
  };

  const addNotif = (title: string, message: string, type: 'info'|'warning'|'success'|'error') => {
    const n: Notification = { id: Date.now(), title, message, type, read: false, timestamp: new Date().toLocaleTimeString() };
    const updated = [n, ...notifications].slice(0, 50);
    setNotifications(updated);
    localStorage.setItem('pts_notifs', JSON.stringify(updated));
    addToast(title, message, type);
  };

  const { periodPromotions, reportExceeders, periodDutyStarters } = useMemo(() => {
    const period = getSalaryPeriod();
    const currentYear = new Date().getFullYear();
    const promos: any[] = [];
    const reports: any[] = [];
    const dutyStarters: any[] = [];

    personnel.forEach(p => {
      // Terfi hesaplama: Göreve başlama yıl dönümü 15-14 dönemine denk gelirse kademe terfisi
      const anniv = isAnniversaryInPeriod(p.startDate, period.start, period.end);
      if (anniv) {
        // Hariç tutulmuşsa ekleme
        if (excludedAutoPromotions.includes(p.id)) return;
        
        let nD = p.degree; let nL = p.level + 1;
        // Derece 1 için kademe 4'e kadar çıkabilir, diğer dereceler için 3'te sıfırlanır
        if (nD === 1) {
          if (nL > 4) { 
            // 1/4 zaten son kademe, terfi edemez, çizelgeye ekleme
            return;
          }
        } else {
          if (nL > 3) { nL = 1; nD = Math.max(1, nD - 1); }
        }
        promos.push({ person: p, nextDegree: nD, nextLevel: nL, date: anniv.toLocaleDateString('tr-TR'), type: 'terfi' });
      }

      // Rapor kesinti hesaplama: İlk 7 gün kesintisiz, sonrası kesintiye yansır
      const sickLeaves = (p.leaveHistory || [])
        .filter(l => l.year === currentYear && l.type === 'Hastalık')
        .sort((a, b) => new Date(a.startDate).getTime() - new Date(b.startDate).getTime());
      
      let totalUsedDays = 0;
      let totalDeductionDays = 0;
      const deductionDetails: any[] = [];
      
      sickLeaves.forEach(leave => {
        const leaveDays = leave.duration;
        const remainingFreeDays = Math.max(0, 7 - totalUsedDays);
        
        if (remainingFreeDays > 0) {
          // Hala kesintisiz hak varsa
          const freeDaysUsed = Math.min(leaveDays, remainingFreeDays);
          const deductedDays = leaveDays - freeDaysUsed;
          totalUsedDays += freeDaysUsed;
          
          if (deductedDays > 0) {
            totalDeductionDays += deductedDays;
            deductionDetails.push({
              startDate: leave.startDate,
              endDate: leave.endDate,
              totalDays: leaveDays,
              deductedDays: deductedDays
            });
          }
        } else {
          // 7 günlük hak bitti, tüm rapor kesintiye yansır
          totalDeductionDays += leaveDays;
          deductionDetails.push({
            startDate: leave.startDate,
            endDate: leave.endDate,
            totalDays: leaveDays,
            deductedDays: leaveDays
          });
        }
      });
      
      if (totalDeductionDays > 0) {
        // Bu personel için bu yıl daha önce "Gönderildi" işareti konmuş gün sayısı
        const sentDays = reportSentRecords
          .filter(r => r.personnelId === p.id && r.year === currentYear)
          .reduce((sum, r) => sum + (r.sentDays || 0), 0);
        // Sadece henüz gönderilmemiş kesinti günlerini çizelgeye yansıt.
        // Gönderildi butonuna basıldıktan sonra mevcut kesintiler 0 olur;
        // yeni alınan rapor günleri (7 günlük muafiyet zaten dolduğu için)
        // doğrudan resmi çizelgeye eklenir.
        const unsentDays = Math.max(0, totalDeductionDays - sentDays);
        if (unsentDays > 0) {
          reports.push({
            person: p,
            days: unsentDays,
            excessDays: unsentDays,
            totalDeductionDays,
            sentDays,
            details: deductionDetails,
          });
        }
      }
    });

    // Sadece manuel eklenen göreve başlama kayıtlarını al (otomatik ekleme yok)
    dutyChanges.filter(dc => dc.type === 'Göreve Başlama').forEach(dc => {
      const parts = dc.date.split('.');
      if (parts.length === 3) {
        const dcDate = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
        if (dcDate >= period.start && dcDate <= period.end) {
          const p = personnel.find(x => x.id === dc.personnelId);
          if (p) {
            dutyStarters.push({ person: p, degree: p.degree, level: p.level, date: dc.date });
          }
        }
      }
    });

    return { periodPromotions: promos, reportExceeders: reports, periodDutyStarters: dutyStarters };
  }, [personnel, dutyChanges, excludedAutoPromotions, excludedDutyStarters, reportSentRecords]);

  const runAnalysis = (
    personnelList: Personnel[],
    promos: { person: Personnel; nextDegree: number; nextLevel: number; date: string }[],
    exceeders: { person: Personnel; days: number; excessDays: number }[]
  ): string => {
    const currentYear = new Date().getFullYear();
    const risks: string[] = [];
    const suggestions: string[] = [];

    personnelList.forEach(p => {
      const reportDays = (p.leaveHistory || [])
        .filter(l => l.year === currentYear && l.type === 'Hastalık')
        .reduce((sum, curr) => sum + curr.duration, 0);
      if (reportDays > 7) risks.push(`<strong>${p.name}</strong>: Tek hekim hastalık raporu bu yıl ${reportDays} gün (7 gün sınırı aşıldı).`);
      if (p.performance < 70) risks.push(`<strong>${p.name}</strong>: Performans düşük (${p.performance}%).`);
      if (p.leaveTotal > 20) risks.push(`<strong>${p.name}</strong>: Yıllık izin kullanımı yüksek (${p.leaveTotal} gün).`);
    });

    if (promos.length > 0) {
      suggestions.push(`Maaş döneminde (15-14) kademe/derece değişikliği olacak ${promos.length} personel var; bordro ve resmi formları kontrol edin.`);
    }
    if (exceeders.length > 0) {
      suggestions.push(`Tek hekim hastalık raporu 7 günü aşan ${exceeders.length} personel için sağlık raporu / idari işlem kontrolü önerilir.`);
    }
    const lowPerf = personnelList.filter(p => p.performance < 70);
    if (lowPerf.length > 0) suggestions.push(`Düşük performanslı ${lowPerf.length} personel için performans geliştirme veya eğitim planı değerlendirilebilir.`);
    if (personnelList.length > 0 && personnelList.filter(p => p.leaveTotal === 0).length === personnelList.length) {
      suggestions.push('İzin kullanımı düşük; personeli yıllık izin kullanımı konusunda bilgilendirin.');
    }

    const now = new Date().toLocaleString('tr-TR');
    let html = `<div class="space-y-4"><p class="text-xs text-slate-400 mb-2">Son güncelleme: ${now}</p>`;
    if (risks.length > 0) {
      html += '<p class="font-semibold text-amber-700">⚠️ Tespit edilen riskler / dikkat edilmesi gerekenler:</p><ul class="list-disc pl-5 space-y-1">';
      risks.forEach(r => { html += `<li>${r}</li>`; });
      html += '</ul>';
    } else html += '<p class="text-emerald-600 font-medium">✓ Risk olarak işaretlenen personel bulunmuyor.</p>';
    if (suggestions.length > 0) {
      html += '<p class="font-semibold text-slate-700 mt-3">💡 İyileştirme önerileri:</p><ul class="list-disc pl-5 space-y-1">';
      suggestions.forEach(s => { html += `<li>${s}</li>`; });
      html += '</ul>';
    }
    html += '</div>';
    return html || '<p>Veri yetersiz; personel listesini doldurup tekrar analiz edin.</p>';
  };

  const handleAIAnalysis = () => {
    setAiLoading(true);
    setAiResult('');
    setTimeout(() => {
      try {
        setAiResult(runAnalysis(personnel, periodPromotions, reportExceeders));
      } catch {
        setAiResult('<p>Analiz sırasında hata oluştu.</p>');
      }
      setAiLoading(false);
    }, 300);
  };

  // Dashboard'a geçildiğinde veya personel verisi değiştiğinde analizi otomatik güncelle
  useEffect(() => {
    if (activeTab !== 'dashboard' || personnel.length === 0) return;
    const html = runAnalysis(personnel, periodPromotions, reportExceeders);
    setAiResult(html);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTab, personnel]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = () => {
        setNewAnnounce({
          ...newAnnounce,
          fileName: file.name,
          fileType: file.type,
          fileData: reader.result as string
        });
      };
      reader.readAsDataURL(file);
    }
  };

  const handleAddAnnouncement = () => {
    if (!newAnnounce.title) return addToast("Hata", "Duyuru başlığı boş olamaz.", "error");
    const a: Announcement = { 
      ...newAnnounce as Announcement, 
      id: Date.now(), 
      date: new Date().toLocaleDateString('tr-TR') 
    };
    const updated = [a, ...announcements];
    setAnnouncements(updated);
    localStorage.setItem('pts_announcements', JSON.stringify(updated));
    setShowAnnounceModal(false);
    setNewAnnounce({ type: 'text' });
    addNotif("Duyuru Eklendi", a.title, "info");
  };

  const downloadFile = (fileData: string, fileName: string) => {
    const link = document.createElement('a');
    link.href = fileData;
    link.download = fileName;
    link.click();
  };

  // Düzenleme Modunu Aç
  const handleEditPersonnel = (p: Personnel) => {
    setEditingPersonId(p.id);
    setNewPerson({ ...p });
    setShowAddModal(true);
  };

  const handleViewDetails = (p: Personnel) => {
    setSelectedPersonForDetail(p);
    setShowDetailModal(true);
  };

  // Yeni Kayıt Modunu Aç
  const handleAddNewPersonnel = () => {
    setEditingPersonId(null);
    setNewPerson({
        role: ROLES[3],
        education: EDUCATIONS[2],
        employmentType: EMPLOYMENT_TYPES[0],
        maritalStatus: MARITAL_STATUSES[1],
        degree: 9,
        level: 1,
        childrenCount: 0,
        branch: "",
        title: TITLES[1],
        startDate: new Date().toISOString().split('T')[0]
    });
    setShowAddModal(true);
  };

  const handleSavePersonnel = () => {
    if (!newPerson.name || !newPerson.tc) return addToast("Hata", "Lütfen gerekli alanları (Ad Soyad ve TC) doldurun.", "error");
    
    if (editingPersonId) {
        // GÜNCELLEME
        const updatedPersonnel = personnel.map(p => p.id === editingPersonId ? {
            ...p,
            ...newPerson as Personnel,
            id: editingPersonId // ID değişmesin
        } : p);
        setPersonnel(updatedPersonnel);
        localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
        addNotif("Personel Güncellendi", newPerson.name!, "success");
    } else {
        // YENİ KAYIT
        const p: Personnel = { 
          id: Date.now(), 
          name: newPerson.name!, 
          tc: newPerson.tc!, 
          startDate: newPerson.startDate || new Date().toISOString().split('T')[0], 
          degree: Number(newPerson.degree) || 9, 
          level: Number(newPerson.level) || 1, 
          branch: newPerson.branch || "Belirtilmedi", 
          performance: 100, 
          leaveTotal: 0, 
          leaveHistory: [], 
          salaryHistory: [], 
          phone: newPerson.phone || "", 
          email: newPerson.email || "", 
          address: newPerson.address || "", 
          personnelNo: newPerson.personnelNo || "SİCİL-"+Date.now().toString().slice(-4), 
          iban: newPerson.iban || "", 
          maritalStatus: newPerson.maritalStatus || MARITAL_STATUSES[1], 
          childrenCount: Number(newPerson.childrenCount) || 0, 
          trainings: [], 
          education: newPerson.education || EDUCATIONS[2], 
          role: newPerson.role || ROLES[3], 
          title: newPerson.title || TITLES[1], 
          employmentType: newPerson.employmentType || EMPLOYMENT_TYPES[0],
          union: newPerson.union || ""
        };
        const updated = [...personnel, p]; 
        setPersonnel(updated); 
        localStorage.setItem('pts_personnel', JSON.stringify(updated)); 
        addNotif("Personel Eklendi", p.name, "success");
    }

    setShowAddModal(false); 
    setEditingPersonId(null);
  };

  const branchData = useMemo(() => {
    const counts: Record<string, number> = {};
    personnel.forEach(p => {
      const raw = (p.branch || '').trim();
      const key = raw || 'Branş Belirtilmemiş';
      counts[key] = (counts[key] || 0) + 1;
    });
    return Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
  }, [personnel]);

  const renderBranchLabel = (props: any) => {
    const { x, y, width, height, value } = props;
    const label = value as string | undefined;
    if (!label) return null;
    const cx = x + width / 2;
    const cy = y + height / 2;
    // Çok kısa sütunlarda yazıyı yukarıya, uzunlarda içeri dikey yaz
    if (height < 24) {
      return (
        <text
          x={cx}
          y={y - 4}
          textAnchor="middle"
          dominantBaseline="auto"
          fill="#475569"
          fontSize={10}
        >
          {label}
        </text>
      );
    }
    return (
      <text
        x={cx}
        y={cy}
        textAnchor="middle"
        dominantBaseline="central"
        fill="#ffffff"
        fontSize={10}
        transform={`rotate(-90 ${cx} ${cy})`}
      >
        {label}
      </text>
    );
  };

  // Toplam personel sayısı - sadece personel listesindeki kişileri say
  const totalPersonnelCount = useMemo(() => {
    return personnel.length;
  }, [personnel.length]);

  const roleData = useMemo(() => {
    const counts: Record<string, number> = {};

    // Personel listesinden rolleri say
    personnel.forEach(p => {
      const key = (p.role && p.role.trim()) ? p.role.trim() : 'Belirtilmemiş';
      counts[key] = (counts[key] || 0) + 1;
    });

    return Object.entries(counts)
      .filter(([, v]) => v > 0)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
  }, [personnel]);

  const visiblePersonnelBase = useMemo(() => {
    if (currentUser?.role === 'user' && currentUser.personnelId) {
      return personnel.filter(p => p.id === currentUser.personnelId);
    }
    return personnel;
  }, [personnel, currentUser]);

  const personnelFiltered = useMemo(() => {
    const q = (search || '').trim().toLocaleLowerCase('tr-TR');
    if (!q) return visiblePersonnelBase;
    return visiblePersonnelBase.filter(p =>
      p.name.toLocaleLowerCase('tr-TR').includes(q) ||
      p.tc.includes(search) ||
      p.personnelNo.includes(search)
    );
  }, [visiblePersonnelBase, search]);

  const getLeaveDaysForYear = (p: Personnel) => {
    const currentYear = new Date().getFullYear();
    return (p.leaveHistory || [])
      .filter(l => l.year === currentYear && l.type === selectedLeaveFilter)
      .reduce((sum, curr) => sum + curr.duration, 0);
  };

  const personnelSorted = useMemo(() => {
    const dir = personnelSort.dir === 'asc' ? 1 : -1;
    const collator = new Intl.Collator('tr-TR', { sensitivity: 'base', numeric: true });
    const arr = [...personnelFiltered];
    arr.sort((a, b) => {
      switch (personnelSort.key) {
        case 'tc':
          return dir * collator.compare(a.tc, b.tc);
        case 'branch': {
          const aKey = `${(a.branch || '').trim()} ${(a.title || '').trim()}`.trim();
          const bKey = `${(b.branch || '').trim()} ${(b.title || '').trim()}`.trim();
          return dir * collator.compare(aKey, bKey) || collator.compare(a.name, b.name);
        }
        case 'degree':
          return dir * ((a.degree - b.degree) || (a.level - b.level) || collator.compare(a.name, b.name));
        case 'leave':
          return dir * ((getLeaveDaysForYear(a) - getLeaveDaysForYear(b)) || collator.compare(a.name, b.name));
        case 'name':
        default:
          return dir * collator.compare(a.name, b.name);
      }
    });
    return arr;
  }, [personnelFiltered, personnelSort, selectedLeaveFilter]);

  const handleLogin = () => {
    const users = (settings.users && Array.isArray(settings.users) ? settings.users : DEFAULT_SETTINGS.users!) as AppUser[];
    const found = users.find(u => u.username === loginUsername);
    if (!found || found.password !== loginPassword) return addToast('Hata', 'Hatalı kullanıcı adı veya şifre', 'error');
    setCurrentUser({ username: found.username, role: found.role, personnelId: found.personnelId });
  };

  if (!currentUser) {
    return (
      <div className="h-screen bg-gradient-to-br from-slate-100 via-slate-50 to-blue-50 flex flex-col">
        <TitleBar title="Giriş Paneli" printOrientation="landscape" />
        <div className="flex-1 flex items-center justify-center p-4">
        <div className="bg-white p-10 rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-100" style={{width: '400px', maxWidth: '100%'}}>
          <div className="flex justify-center mb-6">
            <div className={`w-16 h-16 ${getPrimaryColorClass(settings.primaryColor, 'bg')} rounded-2xl flex items-center justify-center shadow-lg text-white`}>
              <Users size={32} />
            </div>
          </div>
          <h2 className="text-center text-2xl font-bold text-slate-800 mb-2 tracking-tight">Personel Takip Sistemi</h2>
          <p className="text-center text-sm text-slate-400 mb-8">{settings.schoolTitle}</p>
          <form onSubmit={(e) => { e.preventDefault(); handleLogin(); }} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-500 mb-1.5 ml-1">Kullanıcı Adı</label>
              <input 
                type="text" 
                placeholder="Kullanıcı Adı" 
                value={loginUsername} 
                onChange={e => setLoginUsername(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleLogin()}
                autoFocus
                className={`w-full px-4 py-3 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 ${getPrimaryColorClass(settings.primaryColor, 'ring')} transition-shadow`} 
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-500 mb-1.5 ml-1">Şifre</label>
              <input 
                type="password" 
                placeholder="Şifre" 
                value={loginPassword} 
                onChange={e => setLoginPassword(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleLogin()}
                className={`w-full px-4 py-3 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 ${getPrimaryColorClass(settings.primaryColor, 'ring')} transition-shadow`} 
              />
            </div>
            <button type="submit" className={`w-full py-3.5 mt-4 ${getPrimaryColorClass(settings.primaryColor, 'bg')} text-white ${getPrimaryColorClass(settings.primaryColor, 'hover-bg')} rounded-xl font-semibold transition-colors shadow-sm`}>Giriş Yap</button>
          </form>
          <div className="mt-6 text-xs text-center text-slate-400">
            <p>Roller: admin, manager, user (Şifre: 123456)</p>
          </div>
        </div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen bg-slate-50 font-sans text-slate-800">
      <TitleBar title={`${settings.schoolTitle}`} printOrientation={settings.printOrientation} />
      
      <div className="flex flex-1 overflow-hidden">
        {/* Sidebar */}
        <div className="no-print w-64 bg-slate-900 text-slate-300 p-5 flex flex-col border-r border-slate-800">
          <div className="text-center mb-6 pb-5 border-b border-slate-800">
             <div className="text-lg font-bold text-white tracking-wide">{settings.schoolTitle}</div>
          </div>
          
          {/* Global Search Bar */}
          <div className="mb-6 relative">
            <input 
              type="text" 
              placeholder="Hızlı Ara..." 
              value={search} 
              onChange={e => setSearch(e.target.value)}
              className="w-full py-2.5 pl-4 pr-10 rounded-xl bg-slate-800 border border-slate-700 text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all"
            />
            {search ? (
              <button 
                onClick={() => setSearch("")}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white transition-colors"
              ><AlertCircle size={14} /></button>
            ) : (
              <Search size={14} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-500" />
            )}
          </div>

          <div className="flex-1 flex flex-col gap-1.5">
            {[
              { id: 'dashboard', label: 'Dashboard', icon: <LayoutDashboard size={18} />, roles: ['admin', 'manager', 'user'] },
              { id: 'personnel', label: 'Personel Listesi', icon: <Users size={18} />, roles: ['admin', 'manager', 'user'] },
              { id: 'students', label: 'Öğrenci İstatistiği', icon: <GraduationCap size={18} />, roles: ['admin', 'manager', 'user'] },
              { id: 'leave', label: 'İzin Yönetimi', icon: <Calendar size={18} />, roles: ['admin', 'manager', 'user'] },
              { id: 'documents', label: 'Evrak Kayıt', icon: <FileText size={18} />, roles: ['admin', 'manager'] },
              { id: 'official', label: 'Resmi Çizelge', icon: <FileSpreadsheet size={18} />, roles: ['admin', 'manager'] },
              { id: 'schedules', label: 'Ders & Nöbet', icon: <Clock size={18} />, roles: ['admin', 'manager', 'user'] },
              { id: 'org', label: 'Organizasyon', icon: <Network size={18} />, roles: ['admin', 'manager'] },
              { id: 'settings', label: 'Ayarlar', icon: <Settings size={18} />, roles: ['admin', 'manager'] }
            ].filter(tab => tab.roles.includes(currentUser.role)).map(tab => (
              <button key={tab.id} onClick={() => setActiveTab(tab.id as any)} 
                className={`w-full text-left px-4 py-3 rounded-xl flex items-center gap-3 transition-all duration-200 text-sm ${
                  activeTab === tab.id 
                    ? `${getPrimaryColorClass(settings.primaryColor, 'bg')} text-white font-medium shadow-md` 
                    : 'text-slate-400 hover:bg-slate-800 hover:text-slate-200'
                }`}>
                {tab.icon}
                {tab.label}
              </button>
            ))}
          </div>
          <div className="text-[11px] text-slate-500 text-center mt-6 pt-4 border-t border-slate-800">
            <div className="flex justify-center gap-4 mb-3">
              <button 
                onClick={() => setShowNotificationPanel(true)} 
                className="text-slate-400 hover:text-white transition-colors relative"
                title="Bildirimler"
              >
                <Bell size={18} />
                {notifications.filter(n => !n.read).length > 0 && (
                  <span className="absolute -top-1 -right-1 w-2.5 h-2.5 bg-red-500 rounded-full border-2 border-slate-900 animate-pulse"></span>
                )}
              </button>
              <button onClick={() => setCurrentUser(null)} className="text-slate-400 hover:text-red-400 transition-colors" title="Çıkış Yap"><LogOut size={18} /></button>
            </div>
            <div className="mb-2 font-medium text-slate-300">
              {currentUser.username} ({currentUser.role})
            </div>
            <div>© 2026 MEB Personel Takip Sistemi</div>
            <div className="text-[10px] text-slate-500 mt-0.5">Geliştirici: Hidayet SEVDİ</div>
          </div>
        </div>

        {/* Main Content Area */}
        <div className={`flex-1 overflow-y-auto ${activeTab === 'official' ? 'p-0' : 'p-8'}`}>
          
          {activeTab === 'dashboard' && (
            <div className="flex flex-col gap-8 max-w-7xl mx-auto">
               {/* Stats Row */}
               <div className={settings.dashboardLayout === 'list' ? "flex flex-col gap-4" : "grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6"}>
                  <div className={`bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between group hover:shadow-md transition-all ${settings.dashboardLayout === 'list' ? 'flex-row' : ''}`}>
                     <div>
                       <div className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Toplam Personel</div>
                       <div className="text-3xl font-bold mt-2 text-slate-800">{totalPersonnelCount}</div>
                     </div>
                     <div className={`${getPrimaryColorClass(settings.primaryColor, 'bg-light')} p-4 rounded-2xl ${getPrimaryColorClass(settings.primaryColor, 'text')} group-hover:scale-110 transition-transform`}><Users size={24} /></div>
                  </div>
                  <div className={`bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between group hover:shadow-md transition-all ${settings.dashboardLayout === 'list' ? 'flex-row' : ''}`}>
                     <div>
                       <div className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Dönem Terfisi</div>
                       <div className="text-3xl font-bold mt-2 text-slate-800">{periodPromotions.length}</div>
                     </div>
                     <div className="bg-amber-50 p-4 rounded-2xl text-amber-500 group-hover:scale-110 transition-transform"><GraduationCap size={24} /></div>
                  </div>
                  <div className={`bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between group hover:shadow-md transition-all ${settings.dashboardLayout === 'list' ? 'flex-row' : ''}`}>
                     <div>
                       <div className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Rapor Sınırı Aşan</div>
                       <div className="text-3xl font-bold mt-2 text-slate-800">{reportExceeders.length}</div>
                     </div>
                     <div className="bg-rose-50 p-4 rounded-2xl text-rose-500 group-hover:scale-110 transition-transform"><AlertCircle size={24} /></div>
                  </div>
                  <div className={`bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between group hover:shadow-md transition-all ${settings.dashboardLayout === 'list' ? 'flex-row' : ''}`}>
                     <div>
                       <div className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Okul Performansı</div>
                       <div className="text-3xl font-bold mt-2 text-slate-800">%{personnel.length ? Math.round(personnel.reduce((a,b)=>a+b.performance,0)/personnel.length) : 0}</div>
                     </div>
                     <div className="bg-emerald-50 p-4 rounded-2xl text-emerald-500 group-hover:scale-110 transition-transform"><CheckCircle2 size={24} /></div>
                  </div>
               </div>

               {/* 7 Gün Sınırını Aşan Personel - Gönderildi İşaretleme */}
               {(reportExceeders.length > 0 || reportSentRecords.filter(r => r.year === new Date().getFullYear()).length > 0) && currentUser.role !== 'user' && (
                 <div className="bg-white p-6 rounded-2xl border border-rose-100 shadow-sm">
                   <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
                     <h3 className="text-base font-semibold text-slate-800 flex items-center gap-2">
                       <AlertCircle size={20} className="text-rose-500" />
                       Rapor Sınırı (7 Gün) Aşan Personel
                     </h3>
                     <span className="text-xs text-slate-500">
                       Resmi çizelgeye gönderdikten sonra "Gönderildi" butonuna basın. Yeni rapor alındığında 7 gün muafiyeti beklenmeden tüm günler çizelgeye işlenir.
                     </span>
                   </div>
                   {reportExceeders.length === 0 ? (
                     <div className="text-center text-slate-400 text-sm py-6">
                       Şu anda sınırı aşan personel bulunmuyor. ✓
                     </div>
                   ) : (
                     <div className="overflow-x-auto">
                       <table className="w-full text-sm">
                         <thead>
                           <tr className="border-b border-slate-200">
                             <th className="text-left py-2 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Adı Soyadı</th>
                             <th className="text-left py-2 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Sicil</th>
                             <th className="text-center py-2 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Çizelgeye İşlenecek</th>
                             <th className="text-center py-2 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Toplam Kesinti</th>
                             <th className="text-center py-2 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Daha Önce Gönderilen</th>
                             <th className="text-right py-2 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">İşlem</th>
                           </tr>
                         </thead>
                         <tbody>
                           {reportExceeders.map((item: any, idx: number) => (
                             <tr key={`exceeder-${item.person.id}-${idx}`} className="border-b border-slate-100 hover:bg-rose-50/30">
                               <td className="py-3 px-3 font-semibold text-slate-800">{item.person.name}</td>
                               <td className="py-3 px-3 text-slate-600 font-mono text-xs">{item.person.personnelNo}</td>
                               <td className="py-3 px-3 text-center">
                                 <span className="inline-flex items-center justify-center px-2.5 py-1 rounded-full bg-rose-100 text-rose-700 text-xs font-bold border border-rose-200">
                                   {item.days} gün
                                 </span>
                               </td>
                               <td className="py-3 px-3 text-center text-slate-600 text-xs">
                                 {item.totalDeductionDays || item.days} gün
                               </td>
                               <td className="py-3 px-3 text-center text-slate-600 text-xs">
                                 {item.sentDays || 0} gün
                               </td>
                               <td className="py-3 px-3 text-right">
                                 <button
                                   onClick={() => markReportAsSent(item.person.id, item.days, item.person.name)}
                                   className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-semibold transition-colors shadow-sm shadow-emerald-600/20 inline-flex items-center gap-1.5"
                                   title="Bu personelin mevcut kesinti günlerini resmi çizelgeye gönderildi olarak işaretle"
                                 >
                                   <CheckCircle2 size={14} />
                                   Gönderildi
                                 </button>
                               </td>
                             </tr>
                           ))}
                         </tbody>
                       </table>
                     </div>
                   )}
                   {/* Bu yıl gönderim yapılmış personellerin geçmişi - geri alma için */}
                   {(() => {
                     const cy = new Date().getFullYear();
                     const sentByPerson: Record<number, { name: string; total: number; lastAt: string }> = {};
                     reportSentRecords.filter(r => r.year === cy).forEach(r => {
                       const p = personnel.find(x => x.id === r.personnelId);
                       if (!p) return;
                       const cur = sentByPerson[r.personnelId] || { name: p.name, total: 0, lastAt: r.sentAt };
                       cur.total += r.sentDays;
                       if (new Date(r.sentAt).getTime() > new Date(cur.lastAt).getTime()) cur.lastAt = r.sentAt;
                       sentByPerson[r.personnelId] = cur;
                     });
                     const entries = Object.entries(sentByPerson);
                     if (entries.length === 0) return null;
                     return (
                       <details className="mt-4 border-t border-slate-100 pt-4">
                         <summary className="cursor-pointer text-xs font-semibold text-slate-500 uppercase tracking-wider hover:text-slate-700">
                           Bu Yıl Gönderilmiş Kesintiler ({entries.length} personel)
                         </summary>
                         <div className="mt-3 space-y-2">
                           {entries.map(([pidStr, info]) => {
                             const pid = Number(pidStr);
                             return (
                               <div key={pid} className="flex items-center justify-between bg-slate-50 rounded-lg px-3 py-2">
                                 <div>
                                   <div className="text-sm font-semibold text-slate-700">{info.name}</div>
                                   <div className="text-xs text-slate-500">
                                     Toplam <span className="font-bold text-slate-700">{info.total} gün</span> gönderildi
                                     {info.lastAt && ` • Son: ${new Date(info.lastAt).toLocaleDateString('tr-TR')}`}
                                   </div>
                                 </div>
                                 <button
                                   onClick={() => undoReportSent(pid, info.name)}
                                   className="px-2.5 py-1 bg-white border border-slate-200 hover:border-amber-300 hover:bg-amber-50 text-slate-600 hover:text-amber-700 rounded-md text-xs font-medium transition-colors"
                                   title="En son yapılan gönderim işaretini geri al"
                                 >
                                   ↶ Son işareti geri al
                                 </button>
                               </div>
                             );
                           })}
                         </div>
                       </details>
                     );
                   })()}
                 </div>
               )}

               {/* Öğrenci İstatistikleri Özet */}
               {selectedAcademicYear && studentStats.find(s => s.id === selectedAcademicYear) && (
                 <div className="bg-gradient-to-br from-blue-50 to-indigo-50 p-6 rounded-2xl border border-blue-100 shadow-sm">
                   <div className="flex items-center justify-between mb-4">
                     <h3 className="text-base font-semibold text-slate-800 flex items-center gap-2">
                       <GraduationCap size={20} className="text-blue-600" /> 
                       Öğrenci İstatistikleri ({studentStats.find(s => s.id === selectedAcademicYear)?.academicYear})
                     </h3>
                     <button 
                       onClick={() => setActiveTab('students')}
                       className="text-sm text-blue-600 hover:text-blue-700 font-semibold flex items-center gap-1"
                     >
                       Detaylar <ChevronRight size={16} />
                     </button>
                   </div>
                   <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                     <div className="bg-white p-4 rounded-xl shadow-sm">
                       <div className="text-xs text-slate-500 font-medium mb-1">Toplam Şube</div>
                       <div className="text-2xl font-bold text-slate-800">{studentStats.find(s => s.id === selectedAcademicYear)?.totalClasses || 0}</div>
                     </div>
                     <div className="bg-white p-4 rounded-xl shadow-sm">
                       <div className="text-xs text-slate-500 font-medium mb-1">Toplam Öğrenci</div>
                       <div className="text-2xl font-bold text-slate-800">{studentStats.find(s => s.id === selectedAcademicYear)?.totalStudents || 0}</div>
                     </div>
                     <div className="bg-white p-4 rounded-xl shadow-sm">
                       <div className="text-xs text-pink-500 font-medium mb-1">Kız Öğrenci</div>
                       <div className="text-2xl font-bold text-pink-600">{studentStats.find(s => s.id === selectedAcademicYear)?.femaleStudents || 0}</div>
                     </div>
                     <div className="bg-white p-4 rounded-xl shadow-sm">
                       <div className="text-xs text-cyan-500 font-medium mb-1">Erkek Öğrenci</div>
                       <div className="text-2xl font-bold text-cyan-600">{studentStats.find(s => s.id === selectedAcademicYear)?.maleStudents || 0}</div>
                     </div>
                   </div>
                   <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mt-4">
                     <div className="bg-white p-4 rounded-xl shadow-sm">
                       <div className="text-xs text-slate-500 font-medium mb-1">Yabancı Uyruklu</div>
                       <div className="text-2xl font-bold text-purple-600">
                         {studentStats.find(s => s.id === selectedAcademicYear)?.foreignStudents.reduce((sum, f) => sum + f.count, 0) || 0}
                       </div>
                     </div>
                     <div className="bg-white p-4 rounded-xl shadow-sm">
                       <div className="text-xs text-slate-500 font-medium mb-1">Mezun Sayısı</div>
                       <div className="text-2xl font-bold text-amber-600">{studentStats.find(s => s.id === selectedAcademicYear)?.graduates.length || 0}</div>
                     </div>
                     <div className="bg-white p-4 rounded-xl shadow-sm">
                       <div className="text-xs text-slate-500 font-medium mb-1">Sınav Kazanan</div>
                       <div className="text-2xl font-bold text-emerald-600">
                         {studentStats.find(s => s.id === selectedAcademicYear)?.graduates.filter(g => g.passed).length || 0}
                       </div>
                     </div>
                     <div className="bg-white p-4 rounded-xl shadow-sm">
                       <div className="text-xs text-slate-500 font-medium mb-1">Sürekli Devamsız</div>
                       <div className="text-2xl font-bold text-red-600">{studentStats.find(s => s.id === selectedAcademicYear)?.absentees.length || 0}</div>
                     </div>
                   </div>
                 </div>
               )}

               {/* Charts Section */}
               <div className={settings.dashboardLayout === 'list' ? "flex flex-col gap-6" : "grid grid-cols-1 lg:grid-cols-2 gap-6"}>
                  <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                    <h3 className="text-base font-semibold text-slate-800 mb-6 flex items-center gap-2"><BarChart2 size={18} className={getPrimaryColorClass(settings.primaryColor, 'text')} /> Branş Dağılımı</h3>
                    <div className="h-64">
                      {branchData.length > 0 ? (
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={branchData} margin={{ top: 10, right: 20, bottom: 5, left: 0 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="name" axisLine={false} tickLine={false} tick={false} />
                            <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                            <RechartsTooltip cursor={{ fill: '#f8fafc' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.08)' }} formatter={(value: number | undefined) => [`${value || 0} kişi`, 'Kişi sayısı']} labelFormatter={(label) => label} />
                            <Bar
                              dataKey="value"
                              name="Kişi sayısı"
                              fill={settings.primaryColor === 'indigo' ? '#4f46e5' : settings.primaryColor === 'emerald' ? '#059669' : settings.primaryColor === 'rose' ? '#e11d48' : settings.primaryColor === 'amber' ? '#d97706' : '#3b82f6'}
                              radius={[6, 6, 0, 0]}
                              barSize={28}
                            >
                              <LabelList dataKey="name" content={renderBranchLabel} />
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      ) : (
                        <div className="h-full flex items-center justify-center text-slate-400 text-sm">Veri bulunmuyor</div>
                      )}
                    </div>
                  </div>
                  <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                    <h3 className="text-base font-semibold text-slate-800 mb-6 flex items-center gap-2"><PieChartIcon size={18} className={getPrimaryColorClass(settings.primaryColor, 'text')} /> Görev Dağılımı</h3>
                    {roleData.length > 0 ? (
                      <>
                        <div className="w-full flex justify-center" style={{ marginBottom: '16px' }}>
                          <PieChart width={300} height={280}>
                            <Pie 
                              data={roleData} 
                              cx={150} 
                              cy={140} 
                              innerRadius={50} 
                              outerRadius={80} 
                              paddingAngle={3} 
                              dataKey="value"
                              label={(entry) => `${entry.name}: ${entry.value}`}
                              labelLine={false}
                              stroke="#fff"
                              strokeWidth={2}
                            >
                              {roleData.map((entry, index) => (
                                <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} stroke="#fff" strokeWidth={2} />
                              ))}
                            </Pie>
                            <RechartsTooltip 
                              contentStyle={{ 
                                borderRadius: '12px', 
                                border: 'none', 
                                boxShadow: '0 4px 20px rgba(0,0,0,0.08)',
                                padding: '8px 12px'
                              }} 
                            />
                          </PieChart>
                        </div>
                        <div className="border-t border-slate-100 pt-4 space-y-2">
                          {roleData.map((item, idx) => (
                            <div key={idx} className="flex items-center justify-between text-sm py-1">
                              <div className="flex items-center gap-2">
                                <div className="w-3 h-3 rounded-full flex-shrink-0" style={{ backgroundColor: CHART_COLORS[idx % CHART_COLORS.length] }}></div>
                                <span className="text-slate-700">{item.name}</span>
                              </div>
                              <span className="font-semibold text-slate-800">{item.value}</span>
                            </div>
                          ))}
                        </div>
                      </>
                    ) : (
                      <div className="h-64 flex items-center justify-center text-slate-400 text-sm">Veri bulunmuyor</div>
                    )}
                  </div>
               </div>

               <div className={settings.dashboardLayout === 'list' ? "flex flex-col gap-6" : "grid grid-cols-1 lg:grid-cols-5 gap-6"}>
                  <div className={`${settings.dashboardLayout === 'list' ? '' : 'lg:col-span-3'} bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex flex-col`}>
                     <div className="flex justify-between items-center mb-6">
                        <h3 className="text-base font-semibold text-slate-800 flex items-center gap-2">🤖 AI Personel Analiz Raporu</h3>
                        {currentUser.role !== 'user' && (
                          <button onClick={handleAIAnalysis} disabled={aiLoading} className={`px-4 py-2 ${getPrimaryColorClass(settings.primaryColor, 'bg-light')} ${getPrimaryColorClass(settings.primaryColor, 'text')} hover:bg-opacity-80 rounded-xl text-sm font-semibold transition-colors disabled:opacity-50`}>
                             {aiLoading ? "Analiz Ediliyor..." : "Yeni Analiz Başlat"}
                          </button>
                        )}
                     </div>
                     <div className="flex-1 bg-slate-50 rounded-xl p-5 border border-slate-100 overflow-y-auto max-h-[300px]">
                       {aiResult ? (
                         <div className="text-sm text-slate-700 leading-relaxed prose prose-sm max-w-none" dangerouslySetInnerHTML={{ __html: aiResult }} />
                       ) : (
                         <div className="h-full flex flex-col items-center justify-center text-slate-400 text-sm py-10">
                           <div className="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center mb-4"><Search size={24} className="text-slate-300" /></div>
                           Henüz bir analiz yapılmadı. Personel verilerini analiz etmek için yukarıdaki butona tıklayın.
                         </div>
                       )}
                     </div>
                  </div>
                  <div className={`${settings.dashboardLayout === 'list' ? '' : 'lg:col-span-2'} bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex flex-col`}>
                     <div className="flex justify-between items-center mb-6">
                        <h3 className="text-base font-semibold text-slate-800 flex items-center gap-2">📢 Dijital Duyuru Panosu</h3>
                        {currentUser.role !== 'user' && (
                          <button onClick={() => setShowAnnounceModal(true)} className={`w-8 h-8 ${getPrimaryColorClass(settings.primaryColor, 'bg-light')} ${getPrimaryColorClass(settings.primaryColor, 'text')} hover:bg-opacity-80 rounded-full flex items-center justify-center transition-colors`}><Plus size={16} /></button>
                        )}
                     </div>
                     <div className="flex-1 overflow-y-auto max-h-[300px] pr-2 space-y-4">
                        {announcements.filter(a => 
                          !search || 
                          a.title.toLowerCase().includes(search.toLowerCase()) || 
                          (a.content && a.content.toLowerCase().includes(search.toLowerCase()))
                        ).length === 0 ? (
                          <div className="text-center text-slate-400 text-sm py-10">
                            {search ? "Arama kriterine uygun duyuru bulunamadı." : "Duyuru bulunmuyor."}
                          </div>
                        ) : announcements.filter(a => 
                          !search || 
                          a.title.toLowerCase().includes(search.toLowerCase()) || 
                          (a.content && a.content.toLowerCase().includes(search.toLowerCase()))
                        ).map(a => (
                          <div key={a.id} className="p-4 rounded-xl border border-slate-100 bg-slate-50/50 hover:bg-slate-50 transition-colors">
                             <div className="flex justify-between items-start mb-2">
                                <div className="font-semibold text-sm text-slate-800">{a.title}</div>
                                <div className="text-[10px] font-bold bg-white border border-slate-200 px-2 py-0.5 rounded-full text-slate-500">{a.type.toUpperCase()}</div>
                             </div>
                             <div className="text-xs text-slate-400 mb-2 flex items-center gap-1"><Clock size={12} /> {a.date}</div>
                             {a.type === 'text' && a.content && <p className="text-sm text-slate-600 line-clamp-2">{a.content}</p>}
                             {a.type === 'link' && a.url && (
                               <a href={a.url} target="_blank" rel="noopener noreferrer" className="inline-flex items-center gap-1 mt-2 text-sm text-blue-600 hover:text-blue-700 font-medium">🔗 Bağlantıya Git <ChevronRight size={14} /></a>
                             )}
                             {a.type === 'file' && a.fileData && (
                               <div className="mt-3 bg-white border border-slate-200 p-2.5 rounded-lg flex justify-between items-center">
                                 <span className="text-xs text-slate-700 truncate max-w-[70%] flex items-center gap-2"><FileText size={14} className="text-slate-400" /> {a.fileName}</span>
                                 <button onClick={() => downloadFile(a.fileData!, a.fileName!)} className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-md text-xs font-medium transition-colors flex items-center gap-1"><Download size={12} /> İndir</button>
                               </div>
                             )}
                          </div>
                        ))}
                     </div>
                  </div>
               </div>
            </div>
          )}

          {activeTab === 'personnel' && (
            <div className="max-w-7xl mx-auto">
               <div className="flex justify-between items-center mb-6">
                  <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2"><Users size={24} className="text-blue-500" /> Personel Listesi</h2>
                   <div className="flex gap-3">
                    <button
                      onClick={() => {
                        const list = (personnelSorted.length ? personnelSorted : personnel);

                        const safe = (v: any) => (v === undefined || v === null ? '' : String(v));
                        const joinArr = (arr: any) => Array.isArray(arr) ? arr.filter(Boolean).join(', ') : '';
                        const json = (v: any) => {
                          try { return v ? JSON.stringify(v) : ''; } catch { return ''; }
                        };

                        // İzin geçmişini okunabilir formata çevir
                        const formatLeaveHistory = (leaveHistory: LeaveRecord[]) => {
                          if (!leaveHistory || leaveHistory.length === 0) return '';
                          return leaveHistory.map(l => 
                            `${l.type} | ${formatDateTR(l.startDate)} - ${formatDateTR(l.endDate)} | ${l.duration} gün`
                          ).join('\n');
                        };

                        // Maaş geçmişini okunabilir formata çevir
                        const formatSalaryHistory = (salaryHistory: SalaryRecord[]) => {
                          if (!salaryHistory || salaryHistory.length === 0) return '';
                          return salaryHistory.map(s => 
                            `${s.year}/${s.month} - ${s.type}: ${s.amount} TL`
                          ).join('\n');
                        };

                        const aoa: any[][] = [
                          [
                            "Ad Soyad",
                            "TC",
                            "Sicil",
                            "Telefon",
                            "E-posta",
                            "Adres",
                            "IBAN",
                            "Medeni Durum",
                            "Çocuk Sayısı",
                            "Öğrenim",
                            "İstihdam Türü",
                            "Görev (Title)",
                            "Kadro/Rol",
                            "Branş",
                            "Derece",
                            "Kademe",
                            "Başlama Tarihi",
                            "Performans",
                            "Toplam İzin",
                            "Sendika",
                            "Eğitimler",
                            "İzin Geçmişi",
                            "Maaş Geçmişi",
                          ],
                          ...list.map(p => ([
                            safe(p.name),
                            safe(p.tc),
                            safe(p.personnelNo),
                            safe(p.phone),
                            safe(p.email),
                            safe(p.address),
                            safe(p.iban),
                            safe(p.maritalStatus),
                            safe(p.childrenCount),
                            safe(p.education),
                            safe(p.employmentType),
                            safe(p.title),
                            safe(p.role),
                            safe(p.branch),
                            p.degree ?? '',
                            p.level ?? '',
                            formatDateTR(p.startDate),
                            p.performance ?? '',
                            p.leaveTotal ?? '',
                            safe((p as any).union),
                            joinArr((p as any).trainings),
                            formatLeaveHistory((p as any).leaveHistory),
                            formatSalaryHistory((p as any).salaryHistory),
                          ])),
                        ];

                        const ws = XLSX.utils.aoa_to_sheet(aoa);
                        ws['!freeze'] = { xSplit: 0, ySplit: 1 };
                        ws['!autofilter'] = { ref: `A1:W${aoa.length}` };

                        // Basit otomatik genişlik: her sütunda max karaktere göre
                        const colCount = aoa[0]?.length || 0;
                        const maxLens = Array.from({ length: colCount }, () => 10);
                        aoa.forEach(r => {
                          r.forEach((cell: any, idx: number) => {
                            const len = safe(cell).length;
                            if (len > maxLens[idx]) maxLens[idx] = Math.min(60, len);
                          });
                        });
                        ws['!cols'] = maxLens.map(wch => ({ wch }));

                        const wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, "Personel");
                        XLSX.writeFile(wb, "personel_listesi.xlsx");
                      }}
                      className="px-4 py-2 bg-white border border-slate-200 text-slate-700 hover:bg-slate-50 rounded-xl text-sm font-semibold transition-colors flex items-center gap-2 shadow-sm"
                    >
                      <Download size={16} /> Dışa Aktar
                    </button>
                    {currentUser.role !== 'user' && (
                      <>
                        <label className="px-4 py-2 bg-white border border-slate-200 text-slate-700 hover:bg-slate-50 rounded-xl text-sm font-semibold transition-colors flex items-center gap-2 shadow-sm cursor-pointer">
                          <FileSpreadsheet size={16} /> Excel'den İçe Aktar
                          <input type="file" accept=".xlsx, .xls" className="hidden" onChange={(e) => {
                            const file = e.target.files?.[0];
                            if (!file) return;
                            const reader = new FileReader();
                            reader.onload = (evt) => {
                              try {
                                const bstr = evt.target?.result;
                                const wb = XLSX.read(bstr, { type: 'array', cellDates: true });
                                const wsname = wb.SheetNames[0];
                                const ws = wb.Sheets[wsname];
                                const data = XLSX.utils.sheet_to_json(ws, { defval: '' });

                                const normalizeKey = (k: unknown) => {
                                  const s = String(k ?? '').trim();
                                  // Turkish locale lowercasing (İ/I issue) + remove diacritics/combining marks
                                  return s
                                    .toLocaleLowerCase('tr-TR')
                                    .normalize('NFKD')
                                    .replace(/[\u0300-\u036f]/g, '') // combining marks
                                    .replace(/[^\p{L}\p{N}]+/gu, ' ') // punctuation -> space
                                    .replace(/\s+/g, ' ')
                                    .trim();
                                };

                                const toIsoDate = (v: any) => {
                                  if (!v) return '';
                                  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toISOString().slice(0, 10);
                                  if (typeof v === 'number') {
                                    const d = (XLSX as any)?.SSF?.parse_date_code?.(v);
                                    if (d?.y && d?.m && d?.d) {
                                      const mm = String(d.m).padStart(2, '0');
                                      const dd = String(d.d).padStart(2, '0');
                                      return `${d.y}-${mm}-${dd}`;
                                    }
                                  }
                                  // dd.mm.yyyy or dd/mm/yyyy or yyyy-mm-dd
                                  const s = String(v).trim();
                                  const m1 = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
                                  if (m1) return `${m1[3]}-${String(m1[2]).padStart(2,'0')}-${String(m1[1]).padStart(2,'0')}`;
                                  const m2 = s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/);
                                  if (m2) return `${m2[1]}-${String(m2[2]).padStart(2,'0')}-${String(m2[3]).padStart(2,'0')}`;
                                  const d2 = new Date(s);
                                  return Number.isNaN(d2.getTime()) ? '' : d2.toISOString().slice(0, 10);
                                };
                                
                                const newPersonnelList = data.map((row: any, index: number) => {
                                  // Normalize keys: remove leading/trailing spaces and make lowercase for robust matching
                                  const normalizedRow: any = {};
                                  for (const key in row) {
                                    if (Object.prototype.hasOwnProperty.call(row, key)) {
                                      normalizedRow[normalizeKey(key)] = row[key];
                                    }
                                  }

                                  const get = (...keys: string[]) => {
                                    for (const k of keys) {
                                      const v = normalizedRow[normalizeKey(k)];
                                      if (v !== undefined && v !== null && String(v).trim() !== '') return v;
                                    }
                                    return '';
                                  };

                                  const dkRaw = String(get('derece/kademe', 'derece kademe', 'd/k', 'd k', 'derece-kademe')).trim();
                                  const [dkDeg, dkLev] = dkRaw.includes('/') ? dkRaw.split('/').map((x: string) => x.trim()) : ['', ''];

                                  return {
                                    id: Date.now() + Math.floor(Math.random() * 1000000) + index,
                                    name: String(get('ad soyad', 'adı soyadı', 'ad soyadı', 'isim soyisim', 'isim', 'adı', 'personel adı') || 'İsimsiz Personel'),
                                    tc: String(get('tc', 'tc kimlik', 'tc kimlik no', 'tc kimlik numarası', 'tckn') || Math.floor(10000000000 + Math.random() * 90000000000)),
                                    personnelNo: String(get('sicil', 'sicil no', 'sicil numarası', 'personel no', 'personel numarası') || Math.floor(10000 + Math.random() * 90000)),
                                    branch: String(get('branş', 'brans', 'alan', 'branşı') || ''),
                                    title: String(get('görev', 'görevi', 'unvan', 'ünvan', 'kadro görevi') || 'Öğretmen'),
                                    role: (() => {
                                      // Önce özel rol sütunlarına bak
                                      const rawRole = String(get('rol', 'kadro türü', 'görev türü', 'kadro') || '').trim();
                                      if (rawRole) {
                                        const matched = ROLES.find(r => r.toLocaleLowerCase('tr-TR') === rawRole.toLocaleLowerCase('tr-TR'));
                                        if (matched) return matched;
                                      }
                                      // Rol sütunu yoksa 'görev' sütunundaki değeri ROLES ile eşleştir
                                      const rawGörev = String(get('görev', 'görevi', 'unvan', 'ünvan', 'kadro görevi') || '').trim();
                                      if (rawGörev) {
                                        const exact = ROLES.find(r => r.toLocaleLowerCase('tr-TR') === rawGörev.toLocaleLowerCase('tr-TR'));
                                        if (exact) return exact;
                                        const partial = ROLES.find(r => rawGörev.toLocaleLowerCase('tr-TR').includes(r.toLocaleLowerCase('tr-TR').split(' ')[0]));
                                        if (partial) return partial;
                                      }
                                      return 'Öğretmen';
                                    })(),
                                    degree: Number(get('derece') || dkDeg) || 9,
                                    level: Number(get('kademe') || dkLev) || 1,
                                    startDate: toIsoDate(get('başlama tarihi', 'göreve başlama', 'işe giriş tarihi', 'ise giris tarihi')) || new Date().toISOString().split('T')[0],
                                    phone: String(get('telefon', 'tel', 'gsm') || ''),
                                    email: String(get('e-posta', 'eposta', 'email', 'e mail', 'mail') || ''),
                                    address: String(get('adres', 'ikamet adresi') || ''),
                                    education: String(get('öğrenim', 'ogrenim', 'eğitim', 'egitim') || 'Lisans'),
                                    employmentType: String(get('istihdam türü', 'istihdam turu', 'çalışma şekli', 'calisma sekli') || 'Kadrolu'),
                                    maritalStatus: String(get('medeni durum', 'medeni hali') || 'Bekar'),
                                    childrenCount: Number(get('çocuk sayısı', 'cocuk sayisi') || 0) || 0,
                                    iban: String(get('iban') || ''),
                                    union: String(get('sendika') || ''),
                                    performance: 100,
                                    leaveTotal: 0,
                                    leaveHistory: [],
                                    salaries: [],
                                    trainings: [],
                                    salaryHistory: []
                                  };
                                });

                                if (newPersonnelList.length > 0) {
                                  const withRealName = newPersonnelList.filter(p => p.name && p.name !== 'İsimsiz Personel').length;
                                  setPersonnel(prev => {
                                    const updatedPersonnel = [...prev, ...newPersonnelList];
                                    localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
                                    return updatedPersonnel;
                                  });
                                  if (withRealName === 0 && data.length > 0) {
                                    const headers = Object.keys((data as any[])[0] || {}).slice(0, 8).join(', ');
                                    addToast("Uyarı", `Sütun başlıkları eşleşmedi. Bulunan başlıklar: ${headers || '(bulunamadı)'}`, "warning");
                                  }
                                  addToast("Başarılı", `${newPersonnelList.length} personel eklendi.`, "success");
                                } else {
                                  addToast("Hata", "Excel dosyasında geçerli veri bulunamadı.", "error");
                                }
                              } catch (error) {
                                console.error(error);
                                addToast("Hata", "Excel dosyası okunamadı.", "error");
                              }
                              e.target.value = ''; // Reset input
                            };
                            reader.readAsArrayBuffer(file);
                          }} />
                        </label>
                        <button onClick={handleAddNewPersonnel} className={`px-4 py-2 ${getPrimaryColorClass(settings.primaryColor, 'bg')} text-white ${getPrimaryColorClass(settings.primaryColor, 'hover-bg')} rounded-xl text-sm font-semibold transition-colors flex items-center gap-2 shadow-sm`}>
                          <Plus size={16} /> Personel Ekle
                        </button>
                      </>
                    )}
                  </div>
               </div>
               
               {selectedPersonnelIds.length > 0 && currentUser.role !== 'user' && (
                 <div className="mb-4 p-3 bg-red-50 border border-red-100 rounded-xl flex justify-between items-center">
                   <span className="text-sm font-semibold text-red-700">{selectedPersonnelIds.length} personel seçildi</span>
                   <button onClick={() => {
                     if (window.confirm(`Seçili ${selectedPersonnelIds.length} personeli silmek istediğinize emin misiniz?`)) {
                       const updated = personnel.filter(p => !selectedPersonnelIds.includes(p.id));
                       setPersonnel(updated);
                       localStorage.setItem('pts_personnel', JSON.stringify(updated));
                       setSelectedPersonnelIds([]);
                       addToast("Silindi", "Seçili personeller başarıyla silindi.", "info");
                     }
                   }} className="px-4 py-2 bg-red-600 text-white hover:bg-red-700 rounded-lg text-sm font-semibold transition-colors flex items-center gap-2 shadow-sm">
                     <Trash2 size={16} /> Seçilenleri Sil
                   </button>
                 </div>
               )}

               <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
                 <table className="w-full text-left border-collapse">
                    <thead className="bg-slate-50 border-b border-slate-100">
                      <tr>
                        {currentUser.role !== 'user' && (
                          <th className="py-4 px-6 w-12">
                            <input 
                              type="checkbox" 
                              className="rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                              checked={personnelFiltered.length > 0 && selectedPersonnelIds.length === personnelFiltered.length}
                              onChange={(e) => {
                                if (e.target.checked) {
                                  setSelectedPersonnelIds(personnelFiltered.map(p => p.id));
                                } else {
                                  setSelectedPersonnelIds([]);
                                }
                              }}
                            />
                          </th>
                        )}
                        <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider">
                          <button
                            className="inline-flex items-center gap-1 hover:text-slate-700 transition-colors"
                            onClick={() => setPersonnelSort(s => ({ key: 'name', dir: s.key === 'name' && s.dir === 'asc' ? 'desc' : 'asc' }))}
                            title="Ada göre sırala"
                          >
                            Personel Bilgisi {personnelSort.key === 'name' ? (personnelSort.dir === 'asc' ? '▲' : '▼') : ''}
                          </button>
                        </th>
                        <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider">
                          <button
                            className="inline-flex items-center gap-1 hover:text-slate-700 transition-colors"
                            onClick={() => setPersonnelSort(s => ({ key: 'tc', dir: s.key === 'tc' && s.dir === 'asc' ? 'desc' : 'asc' }))}
                            title="TC kimliğe göre sırala"
                          >
                            TC Kimlik {personnelSort.key === 'tc' ? (personnelSort.dir === 'asc' ? '▲' : '▼') : ''}
                          </button>
                        </th>
                        <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider">
                          <button
                            className="inline-flex items-center gap-1 hover:text-slate-700 transition-colors"
                            onClick={() => setPersonnelSort(s => ({ key: 'branch', dir: s.key === 'branch' && s.dir === 'asc' ? 'desc' : 'asc' }))}
                            title="Branş/göreve göre sırala"
                          >
                            Branş / Görev {personnelSort.key === 'branch' ? (personnelSort.dir === 'asc' ? '▲' : '▼') : ''}
                          </button>
                        </th>
                        <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider text-center">
                          <button
                            className="inline-flex items-center gap-1 hover:text-slate-700 transition-colors"
                            onClick={() => setPersonnelSort(s => ({ key: 'degree', dir: s.key === 'degree' && s.dir === 'asc' ? 'desc' : 'asc' }))}
                            title="Derece/kademeye göre sırala"
                          >
                            D / K {personnelSort.key === 'degree' ? (personnelSort.dir === 'asc' ? '▲' : '▼') : ''}
                          </button>
                        </th>
                        <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider text-center">
                          <select value={selectedLeaveFilter} onChange={e => setSelectedLeaveFilter(e.target.value)} className="bg-transparent font-semibold uppercase focus:outline-none cursor-pointer">
                            <option value="Hastalık">Hastalık (Tek Hekim)</option>
                            <option value="Heyet Raporu">Heyet Raporu</option>
                            <option value="Refakat">Refakat İzni</option>
                            <option value="Yıllık">Yıllık İzin</option>
                            <option value="Mazeret">Mazeret İzni</option>
                            <option value="Ücretsiz">Ücretsiz İzin</option>
                          </select>
                          <div className="text-[10px] mt-1 text-slate-400 normal-case">Toplam ({new Date().getFullYear()})</div>
                          <button
                            className="mt-1 text-[10px] text-slate-400 hover:text-slate-600 transition-colors"
                            onClick={() => setPersonnelSort(s => ({ key: 'leave', dir: s.key === 'leave' && s.dir === 'asc' ? 'desc' : 'asc' }))}
                            title="İzin gününe göre sırala"
                            type="button"
                          >
                            Sırala {personnelSort.key === 'leave' ? (personnelSort.dir === 'asc' ? '▲' : '▼') : ''}
                          </button>
                        </th>
                        <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider text-right">İşlemler</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {personnelSorted.length === 0 ? (
                        <tr><td colSpan={6} className="py-10 text-center text-slate-400 text-sm">Personel bulunamadı.</td></tr>
                      ) : personnelSorted.map(p => {
                        const currentYear = new Date().getFullYear();
                        const reportDays = (p.leaveHistory || [])
                          .filter(l => l.year === currentYear && l.type === selectedLeaveFilter)
                          .reduce((sum, curr) => sum + curr.duration, 0);
                        
                        const isExpanded = expandedLeavePersonId === p.id;
                        const yearLeaves = (p.leaveHistory || []).filter(l => l.year === currentYear).sort((a, b) => new Date(b.startDate).getTime() - new Date(a.startDate).getTime());
                        const leaveSummary: Record<string, number> = {};
                        yearLeaves.forEach(l => { leaveSummary[l.type] = (leaveSummary[l.type] || 0) + l.duration; });
                        
                        return (
                        <React.Fragment key={p.id}>
                        <tr className={`hover:bg-slate-50/50 transition-colors ${isExpanded ? 'bg-blue-50/30' : ''}`}>
                          {currentUser.role !== 'user' && (
                            <td className="py-3 px-6 w-12">
                              <input 
                                type="checkbox" 
                                className="rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                                checked={selectedPersonnelIds.includes(p.id)}
                                onChange={(e) => {
                                  if (e.target.checked) {
                                    setSelectedPersonnelIds([...selectedPersonnelIds, p.id]);
                                  } else {
                                    setSelectedPersonnelIds(selectedPersonnelIds.filter(id => id !== p.id));
                                  }
                                }}
                              />
                            </td>
                          )}
                          <td className="py-3 px-6">
                             <button 
                               onClick={() => setExpandedLeavePersonId(isExpanded ? null : p.id)}
                               className="text-left group"
                               title="İzin detaylarını göster/gizle"
                             >
                               <div className="font-semibold text-slate-800 group-hover:text-blue-600 transition-colors flex items-center gap-1.5">
                                 {p.name}
                                 <ChevronRight size={14} className={`text-slate-400 group-hover:text-blue-500 transition-transform ${isExpanded ? 'rotate-90' : ''}`} />
                               </div>
                               <div className="text-xs text-slate-500 mt-0.5">Sicil: {p.personnelNo}</div>
                             </button>
                          </td>
                          <td className="py-3 px-6 text-sm text-slate-600 font-mono">{p.tc}</td>
                          <td className="py-3 px-6">
                             <div className="text-sm font-medium text-slate-800">{p.title}</div>
                             <div className="text-xs text-slate-600 mt-0.5">
                               <span className="font-semibold text-slate-500">Branş:</span>{' '}
                               <span className="text-blue-700 font-semibold">{p.branch || '—'}</span>
                             </div>
                          </td>
                          <td className="py-3 px-6 text-center">
                             <span className="inline-flex items-center justify-center px-2.5 py-1 rounded-full bg-slate-100 text-slate-700 text-xs font-bold border border-slate-200">{p.degree} / {p.level}</span>
                          </td>
                          <td className="py-3 px-6 text-center">
                             <span className={`inline-flex items-center justify-center px-2.5 py-1 rounded-full text-xs font-bold border ${selectedLeaveFilter === 'Hastalık' && reportDays > 7 ? 'bg-rose-100 text-rose-700 border-rose-200' : 'bg-emerald-50 text-emerald-700 border-emerald-200'}`}>
                               {reportDays} Gün
                             </span>
                          </td>
                          <td className="py-3 px-6 text-right">
                            <div className="flex justify-end gap-2">
                              <button onClick={() => handleViewDetails(p)} className="px-3 py-1.5 bg-indigo-50 text-indigo-600 hover:bg-indigo-100 rounded-lg text-xs font-semibold transition-colors">Detay</button>
                              {currentUser.role !== 'user' && (
                                <>
                                  <button onClick={() => handleEditPersonnel(p)} className="px-3 py-1.5 bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 hover:text-slate-800 rounded-lg text-xs font-semibold transition-colors shadow-sm">Düzenle</button>
                                  <button onClick={() => {
                                    if (window.confirm(`${p.name} isimli personeli silmek istediğinize emin misiniz?`)) {
                                      const updated = personnel.filter(person => person.id !== p.id);
                                      setPersonnel(updated);
                                      localStorage.setItem('pts_personnel', JSON.stringify(updated));
                                      addToast("Silindi", "Personel başarıyla silindi.", "info");
                                    }
                                  }} className="px-2 py-1.5 bg-white border border-red-200 text-red-500 hover:bg-red-50 hover:text-red-700 rounded-lg transition-colors shadow-sm" title="Sil">
                                    <Trash2 size={16} />
                                  </button>
                                </>
                              )}
                            </div>
                          </td>
                        </tr>
                        {isExpanded && (
                          <tr>
                            <td colSpan={currentUser.role !== 'user' ? 7 : 6} className="px-6 py-0">
                              <div className="bg-gradient-to-r from-blue-50 to-indigo-50 rounded-xl border border-blue-100 p-4 mb-3 animate-in">
                                <div className="flex items-center justify-between mb-3">
                                  <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                                    <Calendar size={16} className="text-blue-500" />
                                    {p.name} — {currentYear} Yılı İzin Detayları
                                  </h4>
                                  {Object.keys(leaveSummary).length > 0 && (
                                    <div className="flex gap-2 flex-wrap">
                                      {Object.entries(leaveSummary).map(([type, days]) => (
                                        <span key={type} className={`text-[11px] font-bold px-2 py-0.5 rounded-full border ${
                                          type === 'Hastalık' ? 'bg-rose-50 text-rose-700 border-rose-200' :
                                          type === 'Heyet Raporu' ? 'bg-fuchsia-50 text-fuchsia-700 border-fuchsia-200' :
                                          type === 'Refakat' ? 'bg-sky-50 text-sky-700 border-sky-200' :
                                          type === 'Yıllık' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' :
                                          type === 'Mazeret' ? 'bg-amber-50 text-amber-700 border-amber-200' :
                                          'bg-slate-100 text-slate-700 border-slate-200'
                                        }`}>
                                          {type}: {days} gün
                                        </span>
                                      ))}
                                    </div>
                                  )}
                                </div>
                                {yearLeaves.length === 0 ? (
                                  <p className="text-sm text-slate-400 italic">Bu yıl henüz izin kaydı bulunmuyor.</p>
                                ) : (
                                  <table className="w-full text-sm">
                                    <thead>
                                      <tr className="border-b border-blue-200/60">
                                        <th className="text-left py-2 px-3 text-xs font-semibold text-slate-500">İzin Türü</th>
                                        <th className="text-left py-2 px-3 text-xs font-semibold text-slate-500">Başlangıç</th>
                                        <th className="text-left py-2 px-3 text-xs font-semibold text-slate-500">Bitiş</th>
                                        <th className="text-center py-2 px-3 text-xs font-semibold text-slate-500">Süre</th>
                                        <th className="text-left py-2 px-3 text-xs font-semibold text-slate-500">Açıklama</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {yearLeaves.map(l => (
                                        <tr key={l.id} className="border-b border-blue-100/40 hover:bg-white/50">
                                          <td className="py-2 px-3">
                                            <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-bold ${
                                              l.type === 'Hastalık' ? 'bg-rose-100 text-rose-700' :
                                              l.type === 'Heyet Raporu' ? 'bg-fuchsia-100 text-fuchsia-700' :
                                              l.type === 'Refakat' ? 'bg-sky-100 text-sky-700' :
                                              l.type === 'Yıllık' ? 'bg-emerald-100 text-emerald-700' :
                                              l.type === 'Mazeret' ? 'bg-amber-100 text-amber-700' :
                                              'bg-slate-100 text-slate-700'
                                            }`}>
                                              {l.type}
                                            </span>
                                          </td>
                                          <td className="py-2 px-3 text-slate-700">{formatDateTR(l.startDate)}</td>
                                          <td className="py-2 px-3 text-slate-700">{formatDateTR(l.endDate)}</td>
                                          <td className="py-2 px-3 text-center font-bold text-slate-800">{l.duration} gün</td>
                                          <td className="py-2 px-3 text-slate-500 text-xs">{l.description || '—'}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                    <tfoot>
                                      <tr className="border-t-2 border-blue-200">
                                        <td colSpan={3} className="py-2 px-3 text-xs font-bold text-slate-600 text-right">Toplam:</td>
                                        <td className="py-2 px-3 text-center font-bold text-blue-700">{yearLeaves.reduce((s, l) => s + l.duration, 0)} gün</td>
                                        <td></td>
                                      </tr>
                                    </tfoot>
                                  </table>
                                )}
                              </div>
                            </td>
                          </tr>
                        )}
                        </React.Fragment>
                      )})}
                    </tbody>
                 </table>
               </div>
            </div>
          )}

          {activeTab === 'documents' && (
            <div className="max-w-7xl mx-auto">
              <div className="flex justify-between items-center mb-6">
                <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2"><FileText size={24} className="text-blue-500" /> Evrak Kayıt ve Takip</h2>
                {currentUser.role !== 'user' && (
                  <button onClick={() => {
                    const newDoc: DocumentRecord = {
                      id: Date.now(),
                      documentNumber: `EVR-${Math.floor(Math.random() * 10000)}`,
                      subject: 'Yeni Evrak Konusu',
                      date: new Date().toISOString().split('T')[0],
                      type: 'Gelen',
                      senderReceiver: 'İlçe Milli Eğitim',
                      status: 'Bekliyor'
                    };
                    const updated = [...documents, newDoc];
                    setDocuments(updated);
                    localStorage.setItem('pts_documents', JSON.stringify(updated));
                    addToast("Başarılı", "Yeni evrak kaydı oluşturuldu.", "success");
                  }} className="px-4 py-2 bg-emerald-600 text-white hover:bg-emerald-700 rounded-xl text-sm font-semibold transition-colors flex items-center gap-2 shadow-sm shadow-emerald-600/20">
                    <Plus size={16} /> Yeni Evrak Ekle
                  </button>
                )}
              </div>
              
              <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
                <table className="w-full text-left border-collapse">
                  <thead className="bg-slate-50 border-b border-slate-100">
                    <tr>
                      <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider">Evrak No</th>
                      <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider">Tür</th>
                      <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider">Gönderen / Alıcı</th>
                      <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider">Konu</th>
                      <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider">Tarih</th>
                      <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider text-center">Durum</th>
                      <th className="py-4 px-6 text-xs font-semibold text-slate-500 uppercase tracking-wider text-right">İşlemler</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {documents.length === 0 ? (
                      <tr><td colSpan={7} className="py-10 text-center text-slate-400 text-sm">Henüz evrak kaydı bulunmuyor.</td></tr>
                    ) : documents.map(doc => (
                      <tr key={doc.id} className="hover:bg-slate-50/50 transition-colors">
                        <td className="py-3 px-6 font-semibold text-slate-800 text-sm">{doc.documentNumber}</td>
                        <td className="py-3 px-6">
                          <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-bold ${doc.type === 'Gelen' ? 'bg-emerald-50 text-emerald-600 border border-emerald-200' : 'bg-blue-50 text-blue-600 border border-blue-200'}`}>
                            {doc.type}
                          </span>
                        </td>
                        <td className="py-3 px-6 text-sm text-slate-600">{doc.senderReceiver}</td>
                        <td className="py-3 px-6 text-sm text-slate-600">{doc.subject}</td>
                        <td className="py-3 px-6 text-sm text-slate-500 flex items-center gap-1"><Calendar size={14} /> {formatDateTR(doc.date)}</td>
                        <td className="py-3 px-6 text-center">
                          <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-xs font-bold border ${
                            doc.status === 'Bekliyor' ? 'bg-amber-50 text-amber-600 border-amber-200' : 
                            doc.status === 'İşlemde' ? 'bg-blue-50 text-blue-600 border-blue-200' : 
                            'bg-emerald-50 text-emerald-600 border-emerald-200'
                          }`}>
                            {doc.status}
                          </span>
                        </td>
                        <td className="py-3 px-6 text-right">
                          {currentUser.role !== 'user' && (
                            <button onClick={() => {
                              const updated = documents.map(d => d.id === doc.id ? { ...d, status: d.status === 'Bekliyor' ? 'İşlemde' : 'Tamamlandı' as any } : d);
                              setDocuments(updated);
                              localStorage.setItem('pts_documents', JSON.stringify(updated));
                            }} className="px-3 py-1.5 bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 hover:text-slate-800 rounded-lg text-xs font-semibold transition-colors shadow-sm">Durum Güncelle</button>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {activeTab === 'students' && (
            <div className="max-w-7xl mx-auto">
              <div className="flex justify-between items-center mb-6">
                <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
                  <GraduationCap size={24} className="text-blue-500" /> Öğrenci İstatistiği
                </h2>
                <div className="flex gap-3 items-center">
                  <div className="relative flex items-center gap-2">
                    <select 
                      value={selectedAcademicYear} 
                      onChange={(e) => {
                        setSelectedAcademicYear(e.target.value);
                        setYearSavedIndicator(true);
                        setTimeout(() => setYearSavedIndicator(false), 2000);
                      }}
                      className="px-4 py-2 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Eğitim-Öğretim Yılı Seçin</option>
                      {studentStats.slice().sort((a, b) => b.id.localeCompare(a.id)).map(stat => (
                        <option key={stat.id} value={stat.id}>{stat.academicYear}</option>
                      ))}
                    </select>
                    {yearSavedIndicator && (
                      <span className="text-xs text-emerald-600 font-semibold flex items-center gap-1 animate-pulse">
                        ✓ Kaydedildi
                      </span>
                    )}
                  </div>
                  {selectedAcademicYear && currentUser.role !== 'user' && (
                    <>
                    <button
                      onClick={() => {
                        if (window.confirm(`"${studentStats.find(s => s.id === selectedAcademicYear)?.academicYear}" yılını ve tüm verilerini silmek istediğinizden emin misiniz? Bu işlem geri alınamaz!`)) {
                          const updated = studentStats.filter(s => s.id !== selectedAcademicYear);
                          setStudentStats(updated);
                          setSelectedAcademicYear('');
                          addToast('Başarılı', 'Eğitim-öğretim yılı silindi.', 'success');
                        }
                      }}
                      className="px-4 py-2 bg-red-600 text-white hover:bg-red-700 rounded-xl text-sm font-semibold transition-colors shadow-sm inline-flex items-center gap-2"
                      title="Seçili yılı sil"
                    >
                      <Trash2 size={18} /> Yılı Sil
                    </button>
                    <button
                      onClick={() => {
                        const currentStat = studentStats.find(s => s.id === selectedAcademicYear);
                        if (!currentStat) return;
                        
                        const [y1, y2] = selectedAcademicYear.split('-').map(Number);
                        const nextYearId = `${y1 + 1}-${y2 + 1}`;
                        
                        const teacherCount = personnel.filter(p => p.role === 'Öğretmen').length;
                        const totalPersonnel = personnel.length;
                        const studentCount = currentStat.totalStudents;
                        
                        const confirmMsg = `"${selectedAcademicYear}" → "${nextYearId}" yılına aktarılacaklar:\n\n` +
                          `👨‍🏫 Personel: ${totalPersonnel} kişi (${teacherCount} öğretmen)\n` +
                          `   • Tüm personel bilgileri korunur\n` +
                          `   • Kademe/derece güncellemeleri uygulanır\n` +
                          `   • Performans puanları sıfırlanır (100%)\n\n` +
                          `👨‍🎓 Öğrenci: ${studentCount} öğrenci\n` +
                          `   • Kademeler bir üst sınıfa yükseltilir\n` +
                          `   • Son sınıflar mezunlara aktarılır\n\n` +
                          `Devam etmek istiyor musunuz?`;
                        
                        if (studentStats.find(s => s.id === nextYearId)) {
                          if (!window.confirm(`"${nextYearId} Eğitim-Öğretim Yılı" zaten mevcut.\n\nMevcut verilerin üzerine yazılacak. ${confirmMsg}`)) return;
                        } else {
                          if (!window.confirm(confirmMsg)) return;
                        }
                        
                        // --- PERSONEL AKTARIMI ---
                        const updatedPersonnel = personnel.map(p => {
                          const computed = calculateAutomaticPromotion(p.startDate, p.education);
                          return {
                            ...p,
                            degree: computed.degree,
                            level: computed.level,
                            performance: 100
                          };
                        });
                        setPersonnel(updatedPersonnel);
                        localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
                        
                        // --- ÖĞRENCİ AKTARIMI ---
                        const isLise = (settings.schoolTitle || '').toLowerCase().includes('lise');
                        const mezunKademeler = isLise ? ['Lise-4', 'lise-4', 'Lise 4', '12', '12. Sınıf'] : ['8', '8. Sınıf', '8.Sınıf'];
                        
                        const gradeUpMap: Record<string, string> = isLise
                          ? { 'Lise-1': 'Lise-2', 'lise-1': 'Lise-2', 'Lise 1': 'Lise-2', '9': '10', '9. Sınıf': '10. Sınıf',
                              'Lise-2': 'Lise-3', 'lise-2': 'Lise-3', 'Lise 2': 'Lise-3', '10': '11', '10. Sınıf': '11. Sınıf',
                              'Lise-3': 'Lise-4', 'lise-3': 'Lise-4', 'Lise 3': 'Lise-4', '11': '12', '11. Sınıf': '12. Sınıf' }
                          : { '5': '6', '5. Sınıf': '6. Sınıf', '5.Sınıf': '6.Sınıf',
                              '6': '7', '6. Sınıf': '7. Sınıf', '6.Sınıf': '7.Sınıf',
                              '7': '8', '7. Sınıf': '8. Sınıf', '7.Sınıf': '8.Sınıf' };
                        
                        const newGraduates: GraduateStudent[] = [];
                        const newGradeClassInfo: GradeClassInfo[] = [];
                        let newTotalStudents = 0;
                        let newFemaleStudents = 0;
                        let newMaleStudents = 0;
                        let newTotalClasses = 0;
                        
                        for (const g of currentStat.gradeClassInfo) {
                          const gradeTrimmed = g.grade.trim();
                          
                          if (mezunKademeler.includes(gradeTrimmed)) {
                            for (let i = 0; i < g.femaleCount; i++) {
                              newGraduates.push({
                                id: Date.now() + Math.random() * 10000 + i,
                                name: `${gradeTrimmed} Mezunu (Kız) #${i + 1}`,
                                gender: 'Kız',
                                tookExam: false,
                                passed: false,
                                schoolName: ''
                              });
                            }
                            for (let i = 0; i < g.maleCount; i++) {
                              newGraduates.push({
                                id: Date.now() + Math.random() * 10000 + g.femaleCount + i,
                                name: `${gradeTrimmed} Mezunu (Erkek) #${i + 1}`,
                                gender: 'Erkek',
                                tookExam: false,
                                passed: false,
                                schoolName: ''
                              });
                            }
                          } else if (gradeUpMap[gradeTrimmed]) {
                            const newGrade = gradeUpMap[gradeTrimmed];
                            newGradeClassInfo.push({
                              id: Date.now() + Math.random() * 10000,
                              grade: newGrade,
                              classCount: g.classCount,
                              studentCount: g.studentCount,
                              femaleCount: g.femaleCount,
                              maleCount: g.maleCount
                            });
                            newTotalStudents += g.studentCount;
                            newFemaleStudents += g.femaleCount;
                            newMaleStudents += g.maleCount;
                            newTotalClasses += g.classCount;
                          } else {
                            newGradeClassInfo.push({
                              id: Date.now() + Math.random() * 10000,
                              grade: g.grade,
                              classCount: g.classCount,
                              studentCount: g.studentCount,
                              femaleCount: g.femaleCount,
                              maleCount: g.maleCount
                            });
                            newTotalStudents += g.studentCount;
                            newFemaleStudents += g.femaleCount;
                            newMaleStudents += g.maleCount;
                            newTotalClasses += g.classCount;
                          }
                        }
                        
                        const newStat: StudentStatistics = {
                          id: nextYearId,
                          academicYear: `${nextYearId} Eğitim-Öğretim Yılı`,
                          totalClasses: newTotalClasses,
                          totalStudents: newTotalStudents,
                          femaleStudents: newFemaleStudents,
                          maleStudents: newMaleStudents,
                          gradeClassInfo: newGradeClassInfo.sort((a, b) => a.grade.localeCompare(b.grade, 'tr-TR')),
                          foreignStudents: currentStat.foreignStudents.map(f => ({ ...f, id: Date.now() + Math.random() * 10000 })),
                          graduates: newGraduates,
                          absentees: [],
                          createdDate: new Date().toISOString(),
                          lastModified: new Date().toISOString()
                        };
                        
                        const existingIdx = studentStats.findIndex(s => s.id === nextYearId);
                        let updatedStats: StudentStatistics[];
                        if (existingIdx >= 0) {
                          updatedStats = [...studentStats];
                          updatedStats[existingIdx] = newStat;
                        } else {
                          updatedStats = [...studentStats, newStat];
                        }
                        
                        setStudentStats(updatedStats);
                        setSelectedAcademicYear(nextYearId);
                        
                        const mezunSayisi = newGraduates.length;
                        const aktarilanKademe = newGradeClassInfo.length;
                        addToast('Başarılı', `${nextYearId} yılına aktarıldı: ${totalPersonnel} personel (kademe/derece güncellendi), ${aktarilanKademe} öğrenci kademesi yükseltildi, ${mezunSayisi} öğrenci mezunlara eklendi.`, 'success');
                      }}
                      className="px-4 py-2 bg-amber-600 text-white hover:bg-amber-700 rounded-xl text-sm font-semibold transition-colors shadow-sm inline-flex items-center gap-2"
                      title="Öğretmen ve öğrenci verilerini bir sonraki eğitim-öğretim yılına aktar"
                    >
                      <ArrowRight size={18} /> Sonraki Yıla Aktar
                    </button>
                    </>
                  )}
                  {currentUser.role !== 'user' && (
                  <>
                  <label className="px-4 py-2 bg-emerald-600 text-white hover:bg-emerald-700 rounded-xl text-sm font-semibold transition-colors shadow-sm cursor-pointer inline-flex items-center gap-2">
                    <FileSpreadsheet size={18} /> Excel'den Yükle
                    <input 
                      type="file" 
                      accept=".xlsx, .xls" 
                      className="hidden" 
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (!file) {
                          console.log('❌ Dosya seçilmedi');
                          return;
                        }
                        
                        console.log('📁 Dosya seçildi:', file.name);
                        
                        if (!selectedAcademicYear) {
                          console.log('❌ Eğitim-öğretim yılı seçilmemiş');
                          addToast('Uyarı', 'Lütfen önce bir eğitim-öğretim yılı seçin.', 'warning');
                          e.target.value = '';
                          return;
                        }
                        
                        console.log('✅ Seçili yıl:', selectedAcademicYear);
                        addToast('Bilgi', 'Excel dosyası okunuyor...', 'info');
                        
                        const reader = new FileReader();
                        reader.onerror = (err) => {
                          console.error('❌ Dosya okuma hatası:', err);
                          addToast('Hata', 'Dosya okunamadı.', 'error');
                        };
                        
                        reader.onload = (evt) => {
                          try {
                            console.log('📊 Excel Yükleme Başladı');
                            const bstr = evt.target?.result;
                            if (!bstr) {
                              console.error('❌ Dosya içeriği boş');
                              addToast('Hata', 'Dosya içeriği okunamadı.', 'error');
                              return;
                            }
                            
                            const wb = XLSX.read(bstr, { type: 'array' });
                            console.log('✅ Excel dosyası okundu');
                            console.log('📋 Sheet isimleri:', wb.SheetNames);
                            
                            const normalizeKey = (k: unknown) => {
                              const s = String(k ?? '').trim();
                              return s
                                .toLocaleLowerCase('tr-TR')
                                .normalize('NFKD')
                                .replace(/[\u0300-\u036f]/g, '')
                                .replace(/[^\p{L}\p{N}]+/gu, ' ')
                                .replace(/\s+/g, ' ')
                                .trim();
                            };
                            
                            // Sütun başlığı eşleştirme fonksiyonu
                            const findColumn = (row: any, ...patterns: string[]) => {
                              for (const key of Object.keys(row)) {
                                const normalized = normalizeKey(key);
                                for (const pattern of patterns) {
                                  if (normalized.includes(pattern)) {
                                    return row[key];
                                  }
                                }
                              }
                              return null;
                            };
                            
                            let updated = [...studentStats];
                            const currentStatIndex = updated.findIndex(s => s.id === selectedAcademicYear);
                            if (currentStatIndex === -1) return;
                            
                            // Genel Bilgiler sheet'i - hem 2 sütunlu hem de normal tablo formatını destekle
                            const generalSheet = wb.SheetNames.find(n => 
                              normalizeKey(n).includes('genel') || 
                              normalizeKey(n).includes('ozet') ||
                              normalizeKey(n).includes('bilgi')
                            );
                            
                            if (generalSheet) {
                              console.log('✅ Genel Bilgiler sheet bulundu:', generalSheet);
                              const ws = wb.Sheets[generalSheet];
                              const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
                              console.log('Genel Bilgiler data:', data);
                              
                              // 2 sütunlu format (Başlık, Değer)
                              data.forEach(row => {
                                if (!row[0]) return;
                                const key = normalizeKey(row[0]);
                                const value = Number(row[1]) || 0;
                                
                                if (key.includes('sube') && (key.includes('toplam') || key.includes('sayi'))) {
                                  updated[currentStatIndex].totalClasses = value;
                                } else if (key.includes('ogrenci') && key.includes('toplam')) {
                                  updated[currentStatIndex].totalStudents = value;
                                } else if (key.includes('kiz')) {
                                  updated[currentStatIndex].femaleStudents = value;
                                } else if (key.includes('erkek')) {
                                  updated[currentStatIndex].maleStudents = value;
                                }
                              });
                            }
                            
                            // Kademe Dağılımı sheet'i
                            const gradeSheet = wb.SheetNames.find(n => 
                              normalizeKey(n).includes('kademe') || 
                              normalizeKey(n).includes('sinif') ||
                              normalizeKey(n).includes('dagilim')
                            );
                            
                            if (gradeSheet) {
                              console.log('✅ Kademe Dağılımı sheet bulundu:', gradeSheet);
                              const ws = wb.Sheets[gradeSheet];
                              const data = XLSX.utils.sheet_to_json(ws) as any[];
                              console.log('Kademe Dağılımı data:', data);
                              console.log('İlk satır sütunları:', data[0] ? Object.keys(data[0]) : 'Veri yok');
                              
                              const gradeClassInfo: GradeClassInfo[] = data.map((row, idx) => ({
                                id: Date.now() + idx,
                                grade: String(findColumn(row, 'kademe', 'sinif', 'seviye') || ''),
                                classCount: Number(findColumn(row, 'sube sayi', 'sube', 'sinif sayi') || 0),
                                studentCount: Number(findColumn(row, 'toplam ogrenci', 'toplam', 'ogrenci sayi') || 0),
                                femaleCount: Number(findColumn(row, 'kiz') || 0),
                                maleCount: Number(findColumn(row, 'erkek') || 0)
                              })).filter(g => g.grade);
                              
                              console.log('✅ Kademe bilgileri işlendi:', gradeClassInfo.length, 'kayıt');
                              updated[currentStatIndex].gradeClassInfo = gradeClassInfo;
                            } else {
                              console.log('⚠️ Kademe Dağılımı sheet bulunamadı');
                            }
                            
                            // Yabancı Uyruklu sheet'i
                            const foreignSheet = wb.SheetNames.find(n => 
                              normalizeKey(n).includes('yabanci') || 
                              normalizeKey(n).includes('uyruk') ||
                              normalizeKey(n).includes('ulke')
                            );
                            
                            if (foreignSheet) {
                              const ws = wb.Sheets[foreignSheet];
                              const data = XLSX.utils.sheet_to_json(ws) as any[];
                              
                              const foreignStudents: ForeignStudent[] = data.map((row, idx) => {
                                const country = String(findColumn(row, 'ulke', 'country', 'uyruk') || '');
                                const gender = findColumn(row, 'cinsiyet', 'gender');
                                const count = Number(findColumn(row, 'sayi', 'ogrenci', 'toplam') || 0);
                                
                                return {
                                  id: Date.now() + idx,
                                  country,
                                  gender: (String(gender).toLowerCase().includes('erkek') ? 'Erkek' : 'Kız') as 'Kız' | 'Erkek',
                                  count
                                };
                              }).filter(f => f.country);
                              
                              console.log('✅ Yabancı uyruklu bilgileri işlendi:', foreignStudents.length, 'kayıt');
                              updated[currentStatIndex].foreignStudents = foreignStudents;
                            } else {
                              console.log('⚠️ Yabancı Uyruklu sheet bulunamadı');
                            }
                            
                            // Mezunlar sheet'i
                            const graduateSheet = wb.SheetNames.find(n => 
                              normalizeKey(n).includes('mezun') || 
                              normalizeKey(n).includes('graduate')
                            );
                            
                            if (graduateSheet) {
                              const ws = wb.Sheets[graduateSheet];
                              const data = XLSX.utils.sheet_to_json(ws) as any[];
                              
                              const graduates: GraduateStudent[] = data.map((row, idx) => {
                                const name = String(findColumn(row, 'ogrenci adi', 'ad', 'isim', 'name') || '');
                                const gender = findColumn(row, 'cinsiyet', 'gender');
                                const tookExam = findColumn(row, 'sinav', 'sinava girdi', 'exam');
                                const passed = findColumn(row, 'kazandi', 'kazanan', 'basarili', 'passed');
                                const schoolName = String(findColumn(row, 'okul', 'kazandigi', 'school') || '');
                                
                                return {
                                  id: Date.now() + idx,
                                  name,
                                  gender: (String(gender).toLowerCase().includes('erkek') ? 'Erkek' : 'Kız') as 'Kız' | 'Erkek',
                                  tookExam: String(tookExam).toLowerCase().includes('evet') || String(tookExam).toLowerCase().includes('yes'),
                                  passed: String(passed).toLowerCase().includes('evet') || String(passed).toLowerCase().includes('yes'),
                                  schoolName: schoolName || undefined
                                };
                              }).filter(g => g.name);
                              
                              console.log('✅ Mezun bilgileri işlendi:', graduates.length, 'kayıt');
                              updated[currentStatIndex].graduates = graduates;
                            } else {
                              console.log('⚠️ Mezunlar sheet bulunamadı');
                            }
                            
                            // Sürekli Devamsızlar sheet'i
                            const absenteeSheet = wb.SheetNames.find(n => 
                              normalizeKey(n).includes('devamsiz') || 
                              normalizeKey(n).includes('absent')
                            );
                            
                            if (absenteeSheet) {
                              const ws = wb.Sheets[absenteeSheet];
                              const data = XLSX.utils.sheet_to_json(ws) as any[];
                              
                              const absentees: AbsenteeStudent[] = data.map((row, idx) => {
                                const name = String(findColumn(row, 'ogrenci adi', 'ad', 'isim', 'name') || '');
                                const gender = findColumn(row, 'cinsiyet', 'gender');
                                const grade = String(findColumn(row, 'sinif', 'kademe', 'grade', 'class') || '');
                                const absentDays = Number(findColumn(row, 'devamsizlik', 'gun', 'day', 'absent') || 0);
                                const reason = String(findColumn(row, 'sebep', 'neden', 'reason') || '');
                                
                                return {
                                  id: Date.now() + idx,
                                  name,
                                  gender: (String(gender).toLowerCase().includes('erkek') ? 'Erkek' : 'Kız') as 'Kız' | 'Erkek',
                                  grade,
                                  absentDays,
                                  reason: reason || undefined
                                };
                              }).filter(a => a.name);
                              
                              console.log('✅ Devamsız bilgileri işlendi:', absentees.length, 'kayıt');
                              updated[currentStatIndex].absentees = absentees;
                            } else {
                              console.log('⚠️ Devamsızlar sheet bulunamadı');
                            }
                            
                            updated[currentStatIndex].lastModified = new Date().toISOString();
                            setStudentStats(updated);
                            
                            console.log('✅ Tüm veriler güncellendi:', updated[currentStatIndex]);
                            console.log('📊 Toplam sheet işlendi:', wb.SheetNames.length);
                            addToast('Başarılı', 'Excel verileri başarıyla yüklendi.', 'success');
                            
                          } catch (error) {
                            console.error('Excel okuma hatası:', error);
                            addToast('Hata', 'Excel dosyası okunurken bir hata oluştu.', 'error');
                          }
                        };
                        reader.readAsArrayBuffer(file);
                        e.target.value = '';
                      }}
                    />
                  </label>
                  <button 
                    onClick={() => {
                      setNewYearInput('');
                      setShowNewYearDialog(true);
                    }}
                    className="px-6 py-2 bg-blue-600 text-white hover:bg-blue-700 rounded-xl text-sm font-semibold transition-colors shadow-sm flex items-center gap-2"
                  >
                    <Plus size={18} /> Yeni Yıl Oluştur
                  </button>
                  </>
                  )}
                </div>
              </div>

              {selectedAcademicYear && studentStats.find(s => s.id === selectedAcademicYear) ? (
                <div className="space-y-6">
                  {/* Excel Import Bilgilendirme */}
                  {currentUser.role !== 'user' && (
                  <div className="bg-blue-50 border border-blue-200 rounded-xl p-4">
                    <div className="flex items-start gap-3">
                      <div className="text-blue-600 mt-0.5">
                        <AlertCircle size={20} />
                      </div>
                      <div className="flex-1">
                        <h4 className="font-semibold text-blue-900 mb-1">Excel'den Veri Yükleme</h4>
                        <p className="text-sm text-blue-800 mb-2">
                          Excel dosyanız aşağıdaki sheet isimlerine ve sütun başlıklarına sahip olmalıdır:
                        </p>
                        <ul className="text-xs text-blue-700 space-y-1 ml-4">
                          <li><strong>Genel Bilgiler:</strong> İlk sütun başlık (Toplam Şube Sayısı, Toplam Öğrenci Sayısı, Kız Öğrenci, Erkek Öğrenci), ikinci sütun değer</li>
                          <li><strong>Kademe Dağılımı:</strong> Kademe, Şube Sayısı, Toplam Öğrenci, Kız, Erkek</li>
                          <li><strong>Yabancı Uyruklu:</strong> Ülke, Cinsiyet, Öğrenci Sayısı</li>
                          <li><strong>Mezunlar:</strong> Öğrenci Adı, Cinsiyet, Sınava Girdi (Evet/Hayır), Kazandı (Evet/Hayır), Kazandığı Okul</li>
                          <li><strong>Sürekli Devamsızlar:</strong> Öğrenci Adı, Cinsiyet, Sınıf, Devamsızlık Günü, Sebep</li>
                        </ul>
                      </div>
                    </div>
                  </div>
                  )}

                  {/* Özet Kartlar */}
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                    <div className="bg-gradient-to-br from-blue-500 to-blue-600 text-white p-6 rounded-2xl shadow-lg">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-blue-100 text-sm font-medium">Toplam Şube</span>
                        <Building size={20} className="text-blue-200" />
                      </div>
                      <div className="text-3xl font-bold">{studentStats.find(s => s.id === selectedAcademicYear)?.totalClasses || 0}</div>
                    </div>
                    <div className="bg-gradient-to-br from-emerald-500 to-emerald-600 text-white p-6 rounded-2xl shadow-lg">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-emerald-100 text-sm font-medium">Toplam Öğrenci</span>
                        <GraduationCap size={20} className="text-emerald-200" />
                      </div>
                      <div className="text-3xl font-bold">{studentStats.find(s => s.id === selectedAcademicYear)?.totalStudents || 0}</div>
                    </div>
                    <div className="bg-gradient-to-br from-pink-500 to-pink-600 text-white p-6 rounded-2xl shadow-lg">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-pink-100 text-sm font-medium">Kız Öğrenci</span>
                        <Users size={20} className="text-pink-200" />
                      </div>
                      <div className="text-3xl font-bold">{studentStats.find(s => s.id === selectedAcademicYear)?.femaleStudents || 0}</div>
                    </div>
                    <div className="bg-gradient-to-br from-cyan-500 to-cyan-600 text-white p-6 rounded-2xl shadow-lg">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-cyan-100 text-sm font-medium">Erkek Öğrenci</span>
                        <Users size={20} className="text-cyan-200" />
                      </div>
                      <div className="text-3xl font-bold">{studentStats.find(s => s.id === selectedAcademicYear)?.maleStudents || 0}</div>
                    </div>
                  </div>

                  {/* Veri Giriş Butonları - Sadece admin/manager için */}
                  {currentUser.role !== 'user' && (
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    <button 
                      onClick={() => { setStudentStatsModalType('general'); setShowStudentStatsModal(true); }}
                      className="bg-white p-6 rounded-2xl border-2 border-slate-200 hover:border-blue-400 hover:shadow-lg transition-all text-left group"
                    >
                      <div className="flex items-center gap-3 mb-2">
                        <div className="w-10 h-10 bg-blue-100 rounded-xl flex items-center justify-center group-hover:bg-blue-200 transition-colors">
                          <BarChart2 size={20} className="text-blue-600" />
                        </div>
                        <h3 className="font-semibold text-slate-800">Genel Bilgiler</h3>
                      </div>
                      <p className="text-sm text-slate-500">Şube sayısı, toplam öğrenci, cinsiyet dağılımı</p>
                    </button>

                    <button 
                      onClick={() => { setStudentStatsModalType('gradeClass'); setShowStudentStatsModal(true); }}
                      className="bg-white p-6 rounded-2xl border-2 border-slate-200 hover:border-emerald-400 hover:shadow-lg transition-all text-left group"
                    >
                      <div className="flex items-center gap-3 mb-2">
                        <div className="w-10 h-10 bg-emerald-100 rounded-xl flex items-center justify-center group-hover:bg-emerald-200 transition-colors">
                          <Building size={20} className="text-emerald-600" />
                        </div>
                        <h3 className="font-semibold text-slate-800">Kademe Bazlı Dağılım</h3>
                      </div>
                      <p className="text-sm text-slate-500">Sınıf kademelerine göre şube ve öğrenci sayıları</p>
                    </button>

                    <button 
                      onClick={() => { setStudentStatsModalType('foreign'); setShowStudentStatsModal(true); }}
                      className="bg-white p-6 rounded-2xl border-2 border-slate-200 hover:border-purple-400 hover:shadow-lg transition-all text-left group"
                    >
                      <div className="flex items-center gap-3 mb-2">
                        <div className="w-10 h-10 bg-purple-100 rounded-xl flex items-center justify-center group-hover:bg-purple-200 transition-colors">
                          <MapPin size={20} className="text-purple-600" />
                        </div>
                        <h3 className="font-semibold text-slate-800">Yabancı Uyruklu</h3>
                      </div>
                      <p className="text-sm text-slate-500">Ülke ve cinsiyet bazlı yabancı öğrenci sayıları</p>
                    </button>

                    <button 
                      onClick={() => { setStudentStatsModalType('graduate'); setShowStudentStatsModal(true); }}
                      className="bg-white p-6 rounded-2xl border-2 border-slate-200 hover:border-amber-400 hover:shadow-lg transition-all text-left group"
                    >
                      <div className="flex items-center gap-3 mb-2">
                        <div className="w-10 h-10 bg-amber-100 rounded-xl flex items-center justify-center group-hover:bg-amber-200 transition-colors">
                          <GraduationCap size={20} className="text-amber-600" />
                        </div>
                        <h3 className="font-semibold text-slate-800">Mezun Öğrenciler</h3>
                      </div>
                      <p className="text-sm text-slate-500">Mezuniyet, sınav sonuçları ve kazanılan okullar</p>
                    </button>

                    <button 
                      onClick={() => { setStudentStatsModalType('absentee'); setShowStudentStatsModal(true); }}
                      className="bg-white p-6 rounded-2xl border-2 border-slate-200 hover:border-red-400 hover:shadow-lg transition-all text-left group"
                    >
                      <div className="flex items-center gap-3 mb-2">
                        <div className="w-10 h-10 bg-red-100 rounded-xl flex items-center justify-center group-hover:bg-red-200 transition-colors">
                          <AlertCircle size={20} className="text-red-600" />
                        </div>
                        <h3 className="font-semibold text-slate-800">Sürekli Devamsızlar</h3>
                      </div>
                      <p className="text-sm text-slate-500">Devamsızlık kayıtları ve sebepleri</p>
                    </button>
                  </div>
                  )}

                  {/* Veri Tabloları ve Excel İndirme */}
                  <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-6">
                    <div className="flex justify-between items-center mb-4">
                      <h3 className="text-lg font-semibold text-slate-800">Kayıtlı Veriler</h3>
                      <div className="flex gap-2">
                        <button 
                          onClick={() => {
                            const currentStat = studentStats.find(s => s.id === selectedAcademicYear);
                            if (!currentStat) return;
                            
                            const wb = XLSX.utils.book_new();
                            
                            // Genel Bilgiler
                            const generalData = [
                              ['Eğitim-Öğretim Yılı', currentStat.academicYear],
                              ['Toplam Şube Sayısı', currentStat.totalClasses],
                              ['Toplam Öğrenci Sayısı', currentStat.totalStudents],
                              ['Kız Öğrenci', currentStat.femaleStudents],
                              ['Erkek Öğrenci', currentStat.maleStudents]
                            ];
                            const ws1 = XLSX.utils.aoa_to_sheet(generalData);
                            XLSX.utils.book_append_sheet(wb, ws1, 'Genel Bilgiler');
                            
                            // Kademe Bazlı Dağılım
                            if (currentStat.gradeClassInfo.length > 0) {
                              const ws2 = XLSX.utils.json_to_sheet(currentStat.gradeClassInfo.map(g => ({
                                'Kademe': g.grade,
                                'Şube Sayısı': g.classCount,
                                'Toplam Öğrenci': g.studentCount,
                                'Kız': g.femaleCount,
                                'Erkek': g.maleCount
                              })));
                              XLSX.utils.book_append_sheet(wb, ws2, 'Kademe Dağılımı');
                            }
                            
                            // Yabancı Uyruklu
                            if (currentStat.foreignStudents.length > 0) {
                              const ws3 = XLSX.utils.json_to_sheet(currentStat.foreignStudents.map(f => ({
                                'Ülke': f.country,
                                'Cinsiyet': f.gender,
                                'Öğrenci Sayısı': f.count
                              })));
                              XLSX.utils.book_append_sheet(wb, ws3, 'Yabancı Uyruklu');
                            }
                            
                            // Mezunlar
                            if (currentStat.graduates.length > 0) {
                              const ws4 = XLSX.utils.json_to_sheet(currentStat.graduates.map(g => ({
                                'Öğrenci Adı': g.name,
                                'Cinsiyet': g.gender,
                                'Sınava Girdi': g.tookExam ? 'Evet' : 'Hayır',
                                'Kazandı': g.passed ? 'Evet' : 'Hayır',
                                'Kazandığı Okul': g.schoolName || '-'
                              })));
                              XLSX.utils.book_append_sheet(wb, ws4, 'Mezunlar');
                            }
                            
                            // Sürekli Devamsızlar
                            if (currentStat.absentees.length > 0) {
                              const ws5 = XLSX.utils.json_to_sheet(currentStat.absentees.map(a => ({
                                'Öğrenci Adı': a.name,
                                'Cinsiyet': a.gender,
                                'Sınıf': a.grade,
                                'Devamsızlık Günü': a.absentDays,
                                'Sebep': a.reason || '-'
                              })));
                              XLSX.utils.book_append_sheet(wb, ws5, 'Sürekli Devamsızlar');
                            }
                            
                            XLSX.writeFile(wb, `Ogrenci_Istatistik_${currentStat.id}.xlsx`);
                            addToast('Başarılı', 'Excel dosyası indirildi.', 'success');
                          }}
                          disabled={!selectedAcademicYear}
                          className="px-6 py-2 bg-emerald-600 text-white hover:bg-emerald-700 rounded-xl text-sm font-semibold transition-colors shadow-sm flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                        >
                          <Download size={18} /> Tümünü Excel'e İndir
                        </button>
                      </div>
                    </div>

                    {selectedAcademicYear ? (
                      <div className="space-y-6">
                        {/* Kademe Bazlı Dağılım Tablosu */}
                        {studentStats.find(s => s.id === selectedAcademicYear)?.gradeClassInfo.length! > 0 && (
                          <div>
                            <div className="flex justify-between items-center mb-3">
                              <h4 className="font-semibold text-slate-700">Kademe Bazlı Dağılım</h4>
                              <button 
                                onClick={() => {
                                  const currentStat = studentStats.find(s => s.id === selectedAcademicYear);
                                  if (!currentStat) return;
                                  const ws = XLSX.utils.json_to_sheet(currentStat.gradeClassInfo.map(g => ({
                                    'Kademe': g.grade,
                                    'Şube Sayısı': g.classCount,
                                    'Toplam Öğrenci': g.studentCount,
                                    'Kız': g.femaleCount,
                                    'Erkek': g.maleCount
                                  })));
                                  const wb = XLSX.utils.book_new();
                                  XLSX.utils.book_append_sheet(wb, ws, 'Kademe Dağılımı');
                                  XLSX.writeFile(wb, `Kademe_Dagilimi_${selectedAcademicYear}.xlsx`);
                                }}
                                className="text-xs px-3 py-1.5 bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-700 font-medium transition-colors"
                              >
                                📊 Excel İndir
                              </button>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-sm">
                                <thead className="bg-slate-50">
                                  <tr>
                                    <th className="px-4 py-3 text-left font-semibold text-slate-700">Kademe</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Şube</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Toplam</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Kız</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Erkek</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {studentStats.find(s => s.id === selectedAcademicYear)?.gradeClassInfo.map(g => (
                                    <tr key={g.id} className="border-t border-slate-100 hover:bg-slate-50">
                                      <td className="px-4 py-3 font-medium text-slate-800">{g.grade}</td>
                                      <td className="px-4 py-3 text-center text-slate-700">{g.classCount}</td>
                                      <td className="px-4 py-3 text-center font-semibold text-slate-800">{g.studentCount}</td>
                                      <td className="px-4 py-3 text-center text-pink-600">{g.femaleCount}</td>
                                      <td className="px-4 py-3 text-center text-cyan-600">{g.maleCount}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}

                        {/* Yabancı Uyruklu Öğrenciler */}
                        {studentStats.find(s => s.id === selectedAcademicYear)?.foreignStudents.length! > 0 && (
                          <div>
                            <div className="flex justify-between items-center mb-3">
                              <h4 className="font-semibold text-slate-700">Yabancı Uyruklu Öğrenciler</h4>
                              <button 
                                onClick={() => {
                                  const currentStat = studentStats.find(s => s.id === selectedAcademicYear);
                                  if (!currentStat) return;
                                  const ws = XLSX.utils.json_to_sheet(currentStat.foreignStudents.map(f => ({
                                    'Ülke': f.country,
                                    'Cinsiyet': f.gender,
                                    'Öğrenci Sayısı': f.count
                                  })));
                                  const wb = XLSX.utils.book_new();
                                  XLSX.utils.book_append_sheet(wb, ws, 'Yabancı Uyruklu');
                                  XLSX.writeFile(wb, `Yabanci_Uyruklu_${selectedAcademicYear}.xlsx`);
                                }}
                                className="text-xs px-3 py-1.5 bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-700 font-medium transition-colors"
                              >
                                📊 Excel İndir
                              </button>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-sm">
                                <thead className="bg-slate-50">
                                  <tr>
                                    <th className="px-4 py-3 text-left font-semibold text-slate-700">Ülke</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Cinsiyet</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Öğrenci Sayısı</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {studentStats.find(s => s.id === selectedAcademicYear)?.foreignStudents.map(f => (
                                    <tr key={f.id} className="border-t border-slate-100 hover:bg-slate-50">
                                      <td className="px-4 py-3 font-medium text-slate-800">{f.country}</td>
                                      <td className="px-4 py-3 text-center text-slate-700">{f.gender}</td>
                                      <td className="px-4 py-3 text-center font-semibold text-slate-800">{f.count}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}

                        {/* Mezun Öğrenciler */}
                        {studentStats.find(s => s.id === selectedAcademicYear)?.graduates.length! > 0 && (
                          <div>
                            <div className="flex justify-between items-center mb-3">
                              <h4 className="font-semibold text-slate-700">Mezun Öğrenciler ve Sınav Sonuçları</h4>
                              <button 
                                onClick={() => {
                                  const currentStat = studentStats.find(s => s.id === selectedAcademicYear);
                                  if (!currentStat) return;
                                  const ws = XLSX.utils.json_to_sheet(currentStat.graduates.map(g => ({
                                    'Öğrenci Adı': g.name,
                                    'Cinsiyet': g.gender,
                                    'Sınava Girdi': g.tookExam ? 'Evet' : 'Hayır',
                                    'Kazandı': g.passed ? 'Evet' : 'Hayır',
                                    'Kazandığı Okul': g.schoolName || '-'
                                  })));
                                  const wb = XLSX.utils.book_new();
                                  XLSX.utils.book_append_sheet(wb, ws, 'Mezunlar');
                                  XLSX.writeFile(wb, `Mezunlar_${selectedAcademicYear}.xlsx`);
                                }}
                                className="text-xs px-3 py-1.5 bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-700 font-medium transition-colors"
                              >
                                📊 Excel İndir
                              </button>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-sm">
                                <thead className="bg-slate-50">
                                  <tr>
                                    <th className="px-4 py-3 text-left font-semibold text-slate-700">Öğrenci Adı</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Cinsiyet</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Sınava Girdi</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Kazandı</th>
                                    <th className="px-4 py-3 text-left font-semibold text-slate-700">Kazandığı Okul</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {studentStats.find(s => s.id === selectedAcademicYear)?.graduates.map(g => (
                                    <tr key={g.id} className="border-t border-slate-100 hover:bg-slate-50">
                                      <td className="px-4 py-3 font-medium text-slate-800">{g.name}</td>
                                      <td className="px-4 py-3 text-center text-slate-700">{g.gender}</td>
                                      <td className="px-4 py-3 text-center">{g.tookExam ? '✓' : '✗'}</td>
                                      <td className="px-4 py-3 text-center">{g.passed ? '✓' : '✗'}</td>
                                      <td className="px-4 py-3 text-slate-700">{g.schoolName || '-'}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}

                        {/* Sürekli Devamsızlar */}
                        {studentStats.find(s => s.id === selectedAcademicYear)?.absentees.length! > 0 && (
                          <div>
                            <div className="flex justify-between items-center mb-3">
                              <h4 className="font-semibold text-slate-700">Sürekli Devamsız Öğrenciler</h4>
                              <button 
                                onClick={() => {
                                  const currentStat = studentStats.find(s => s.id === selectedAcademicYear);
                                  if (!currentStat) return;
                                  const ws = XLSX.utils.json_to_sheet(currentStat.absentees.map(a => ({
                                    'Öğrenci Adı': a.name,
                                    'Cinsiyet': a.gender,
                                    'Sınıf': a.grade,
                                    'Devamsızlık Günü': a.absentDays,
                                    'Sebep': a.reason || '-'
                                  })));
                                  const wb = XLSX.utils.book_new();
                                  XLSX.utils.book_append_sheet(wb, ws, 'Devamsızlar');
                                  XLSX.writeFile(wb, `Devamsizlar_${selectedAcademicYear}.xlsx`);
                                }}
                                className="text-xs px-3 py-1.5 bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-700 font-medium transition-colors"
                              >
                                📊 Excel İndir
                              </button>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-sm">
                                <thead className="bg-slate-50">
                                  <tr>
                                    <th className="px-4 py-3 text-left font-semibold text-slate-700">Öğrenci Adı</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Cinsiyet</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Sınıf</th>
                                    <th className="px-4 py-3 text-center font-semibold text-slate-700">Devamsızlık</th>
                                    <th className="px-4 py-3 text-left font-semibold text-slate-700">Sebep</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {studentStats.find(s => s.id === selectedAcademicYear)?.absentees.map(a => (
                                    <tr key={a.id} className="border-t border-slate-100 hover:bg-slate-50">
                                      <td className="px-4 py-3 font-medium text-slate-800">{a.name}</td>
                                      <td className="px-4 py-3 text-center text-slate-700">{a.gender}</td>
                                      <td className="px-4 py-3 text-center text-slate-700">{a.grade}</td>
                                      <td className="px-4 py-3 text-center font-semibold text-red-600">{a.absentDays} gün</td>
                                      <td className="px-4 py-3 text-slate-600 text-xs">{a.reason || '-'}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}
                      </div>
                    ) : (
                      <div className="text-center py-12 text-slate-400">
                        <GraduationCap size={48} className="mx-auto mb-4 opacity-50" />
                        <p>Lütfen bir eğitim-öğretim yılı seçin veya oluşturun</p>
                      </div>
                    )}
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-12 text-center">
                  <GraduationCap size={64} className="mx-auto mb-4 text-slate-300" />
                  <h3 className="text-lg font-semibold text-slate-700 mb-2">Henüz Eğitim-Öğretim Yılı Oluşturulmamış</h3>
                  <p className="text-slate-500 mb-6">Öğrenci istatistiklerini kaydetmek için yeni bir yıl oluşturun</p>
                  <button 
                    onClick={() => {
                      setNewYearInput('');
                      setShowNewYearDialog(true);
                    }}
                    className="px-8 py-3 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-semibold transition-colors shadow-sm inline-flex items-center gap-2"
                  >
                    <Plus size={20} /> Yeni Yıl Oluştur
                  </button>
                </div>
              )}
            </div>
          )}

          {activeTab === 'leave' && (
            <div className="max-w-7xl mx-auto">
              <div className="flex justify-between items-center mb-6">
                <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2"><Calendar size={24} className="text-blue-500" /> {currentUser.role === 'user' ? 'İzin ve Rapor Geçmişim' : 'İzin Yönetimi'}</h2>
              </div>
              <div className={`grid grid-cols-1 ${currentUser.role !== 'user' ? 'lg:grid-cols-3' : ''} gap-6`}>
                {currentUser.role !== 'user' && (
                <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                  <h3 className="text-base font-semibold text-slate-800 mb-5">Yeni İzin Girişi</h3>
                  <div className="flex flex-col gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-slate-500 mb-1.5">Personel Seçin</label>
                      <select value={newLeave.personnelId} onChange={e => setNewLeave({...newLeave, personnelId: Number(e.target.value)})} className="w-full px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <option value="">-- Seçiniz --</option>
                        {personnel.map(p => <option key={p.id} value={p.id}>{p.name} ({p.title})</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-500 mb-1.5">İzin Türü</label>
                      <select value={newLeave.type} onChange={e => setNewLeave({...newLeave, type: e.target.value})} className="w-full px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <option value="Yıllık">Yıllık İzin</option>
                        <option value="Mazeret">Mazeret İzni</option>
                        <option value="Hastalık">Hastalık (Tek Hekim)</option>
                        <option value="Heyet Raporu">Heyet Raporu</option>
                        <option value="Refakat">Refakat İzni</option>
                        <option value="Ücretsiz">Ücretsiz İzin</option>
                      </select>
                    </div>
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-1.5">Başlangıç</label>
                        <input type="date" value={newLeave.startDate} onChange={e => setNewLeave({...newLeave, startDate: e.target.value})} className="w-full px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-1.5">Bitiş</label>
                        <input type="date" value={newLeave.endDate} onChange={e => setNewLeave({...newLeave, endDate: e.target.value})} className="w-full px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                      </div>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-500 mb-1.5">Açıklama</label>
                      <textarea value={newLeave.description} onChange={e => setNewLeave({...newLeave, description: e.target.value})} className="w-full px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none" rows={2}></textarea>
                    </div>
                    <button onClick={() => {
                      if (!newLeave.personnelId || !newLeave.startDate || !newLeave.endDate) {
                        addToast("Hata", "Lütfen personel, başlangıç ve bitiş tarihlerini seçin.", "error");
                        return;
                      }
                      const start = new Date(newLeave.startDate);
                      const end = new Date(newLeave.endDate);
                      const diffTime = Math.abs(end.getTime() - start.getTime());
                      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

                      const updatedPersonnel = personnel.map(p => {
                        if (p.id === newLeave.personnelId) {
                          const newRecord: LeaveRecord = {
                            id: Date.now(),
                            type: newLeave.type,
                            startDate: newLeave.startDate,
                            endDate: newLeave.endDate,
                            duration: diffDays,
                            year: start.getFullYear(),
                            description: newLeave.description
                          };
                          return { ...p, leaveHistory: [...p.leaveHistory, newRecord], leaveTotal: p.leaveTotal + diffDays };
                        }
                        return p;
                      });
                      setPersonnel(updatedPersonnel);
                      localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
                      setNewLeave({ personnelId: '', type: 'Yıllık', startDate: '', endDate: '', description: '' });
                      addToast("Başarılı", "İzin kaydı başarıyla eklendi.", "success");
                    }} className={`w-full py-2.5 ${getPrimaryColorClass(settings.primaryColor, 'bg')} text-white ${getPrimaryColorClass(settings.primaryColor, 'hover-bg')} rounded-xl text-sm font-semibold transition-colors mt-2 shadow-sm`}>İzni Kaydet</button>
                  </div>
                </div>
                )}
                <div className={`${currentUser.role !== 'user' ? 'lg:col-span-2' : ''} bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden flex flex-col`}>
                  <div className="p-6 border-b border-slate-100 flex items-center justify-between">
                    <h3 className="text-base font-semibold text-slate-800">{currentUser.role === 'user' ? 'İzin ve Rapor Geçmişim' : `İzin Kayıtları (${new Date().getFullYear()})`}</h3>
                    {currentUser.role !== 'user' && <span className="text-xs text-slate-400">{personnel.filter(p => (p.leaveHistory || []).some(l => l.year === new Date().getFullYear())).length} personelin izin kaydı var</span>}
                  </div>
                  <div className="flex-1 overflow-y-auto max-h-[600px]">
                    {(currentUser.role === 'user' ? personnel.filter(p => p.id === currentUser.personnelId) : personnel.filter(p => (p.leaveHistory || []).some(l => l.year === new Date().getFullYear()))).length === 0 ? (
                      <div className="py-10 text-center text-slate-400 text-sm">{currentUser.role === 'user' ? 'Henüz izin veya rapor kaydınız bulunmuyor.' : 'Bu yıl izin kaydı bulunan personel yok.'}</div>
                    ) : (
                      <div className="divide-y divide-slate-100">
                        {(currentUser.role === 'user' ? personnel.filter(p => p.id === currentUser.personnelId) : personnel)
                          .filter(p => currentUser.role === 'user' || (p.leaveHistory || []).some(l => l.year === new Date().getFullYear()))
                          .sort((a, b) => {
                            const aLeaves = (a.leaveHistory || []).filter(l => l.year === new Date().getFullYear());
                            const bLeaves = (b.leaveHistory || []).filter(l => l.year === new Date().getFullYear());
                            return bLeaves.reduce((s, l) => s + l.duration, 0) - aLeaves.reduce((s, l) => s + l.duration, 0);
                          })
                          .map(p => {
                            const currentYear = new Date().getFullYear();
                            const yearLeaves = (p.leaveHistory || []).filter(l => l.year === currentYear).sort((a, b) => new Date(b.startDate).getTime() - new Date(a.startDate).getTime());
                            const totalDays = yearLeaves.reduce((s, l) => s + l.duration, 0);
                            const leaveSummary: Record<string, number> = {};
                            yearLeaves.forEach(l => { leaveSummary[l.type] = (leaveSummary[l.type] || 0) + l.duration; });
                            const isOpen = expandedLeaveManageId === p.id;
                            
                            return (
                              <div key={p.id}>
                                <button
                                  onClick={() => { setExpandedLeaveManageId(isOpen ? null : p.id); setEditingLeaveId(null); }}
                                  className={`w-full text-left px-5 py-3.5 flex items-center justify-between hover:bg-slate-50 transition-colors ${isOpen ? 'bg-blue-50/50' : ''}`}
                                >
                                  <div className="flex items-center gap-3 min-w-0">
                                    <ChevronRight size={16} className={`text-slate-400 flex-shrink-0 transition-transform ${isOpen ? 'rotate-90 text-blue-500' : ''}`} />
                                    <div className="min-w-0">
                                      <div className="font-semibold text-slate-800 text-sm truncate">{p.name}</div>
                                      <div className="text-[11px] text-slate-500">{p.title} — {p.branch || '—'}</div>
                                    </div>
                                  </div>
                                  <div className="flex items-center gap-2 flex-shrink-0">
                                    {Object.entries(leaveSummary).map(([type, days]) => (
                                      <span key={type} className={`text-[10px] font-bold px-1.5 py-0.5 rounded-full ${
                                        type === 'Hastalık' ? 'bg-rose-100 text-rose-700' :
                                        type === 'Heyet Raporu' ? 'bg-fuchsia-100 text-fuchsia-700' :
                                        type === 'Refakat' ? 'bg-sky-100 text-sky-700' :
                                        type === 'Yıllık' ? 'bg-emerald-100 text-emerald-700' :
                                        type === 'Mazeret' ? 'bg-amber-100 text-amber-700' :
                                        'bg-slate-100 text-slate-600'
                                      }`}>
                                        {type}: {days}g
                                      </span>
                                    ))}
                                    <span className={`text-xs font-bold px-2 py-0.5 rounded-full border ml-1 ${totalDays > 0 ? 'bg-blue-50 text-blue-700 border-blue-200' : 'bg-slate-50 text-slate-400 border-slate-200'}`}>
                                      {totalDays} gün
                                    </span>
                                  </div>
                                </button>
                                
                                {isOpen && (
                                  <div className="bg-gradient-to-r from-blue-50 to-indigo-50 px-5 pb-4">
                                    {yearLeaves.length === 0 ? (
                                      <p className="text-sm text-slate-400 italic py-3">Bu yıl henüz izin kaydı bulunmuyor.</p>
                                    ) : (
                                      <table className="w-full text-sm">
                                        <thead>
                                          <tr className="border-b border-blue-200/60">
                                            <th className="text-left py-2 px-2 text-[11px] font-semibold text-slate-500">Tür</th>
                                            <th className="text-left py-2 px-2 text-[11px] font-semibold text-slate-500">Başlangıç</th>
                                            <th className="text-left py-2 px-2 text-[11px] font-semibold text-slate-500">Bitiş</th>
                                            <th className="text-center py-2 px-2 text-[11px] font-semibold text-slate-500">Gün</th>
                                            <th className="text-left py-2 px-2 text-[11px] font-semibold text-slate-500">Açıklama</th>
                                            <th className="text-center py-2 px-2 text-[11px] font-semibold text-slate-500">İşlem</th>
                                          </tr>
                                        </thead>
                                        <tbody>
                                          {yearLeaves.map(l => {
                                            const isEditing = editingLeaveId === l.id;
                                            return (
                                              <tr key={l.id} className="border-b border-blue-100/40 hover:bg-white/50">
                                                {isEditing ? (
                                                  <>
                                                    <td className="py-1.5 px-2">
                                                      <select value={editingLeaveData.type} onChange={e => setEditingLeaveData({...editingLeaveData, type: e.target.value})} className="w-full px-1.5 py-1 rounded-lg border border-blue-200 text-xs focus:outline-none focus:ring-1 focus:ring-blue-400 bg-white">
                                                        <option value="Yıllık">Yıllık</option>
                                                        <option value="Mazeret">Mazeret</option>
                                                        <option value="Hastalık">Hastalık (Tek Hekim)</option>
                                                        <option value="Heyet Raporu">Heyet Raporu</option>
                                                        <option value="Refakat">Refakat</option>
                                                        <option value="Ücretsiz">Ücretsiz</option>
                                                      </select>
                                                    </td>
                                                    <td className="py-1.5 px-2">
                                                      <input type="date" value={editingLeaveData.startDate} onChange={e => setEditingLeaveData({...editingLeaveData, startDate: e.target.value})} className="w-full px-1.5 py-1 rounded-lg border border-blue-200 text-xs focus:outline-none focus:ring-1 focus:ring-blue-400 bg-white" />
                                                    </td>
                                                    <td className="py-1.5 px-2">
                                                      <input type="date" value={editingLeaveData.endDate} onChange={e => setEditingLeaveData({...editingLeaveData, endDate: e.target.value})} className="w-full px-1.5 py-1 rounded-lg border border-blue-200 text-xs focus:outline-none focus:ring-1 focus:ring-blue-400 bg-white" />
                                                    </td>
                                                    <td className="py-1.5 px-2 text-center text-xs text-slate-500">
                                                      {editingLeaveData.startDate && editingLeaveData.endDate ? Math.ceil(Math.abs(new Date(editingLeaveData.endDate).getTime() - new Date(editingLeaveData.startDate).getTime()) / (1000 * 60 * 60 * 24)) + 1 : '—'}
                                                    </td>
                                                    <td className="py-1.5 px-2">
                                                      <input type="text" value={editingLeaveData.description} onChange={e => setEditingLeaveData({...editingLeaveData, description: e.target.value})} className="w-full px-1.5 py-1 rounded-lg border border-blue-200 text-xs focus:outline-none focus:ring-1 focus:ring-blue-400 bg-white" placeholder="Açıklama" />
                                                    </td>
                                                    <td className="py-1.5 px-2 text-center">
                                                      <div className="flex items-center justify-center gap-1">
                                                        <button
                                                          onClick={() => {
                                                            if (!editingLeaveData.startDate || !editingLeaveData.endDate) { addToast('Hata', 'Tarih alanları boş bırakılamaz.', 'error'); return; }
                                                            const newDuration = Math.ceil(Math.abs(new Date(editingLeaveData.endDate).getTime() - new Date(editingLeaveData.startDate).getTime()) / (1000 * 60 * 60 * 24)) + 1;
                                                            const updatedPersonnel = personnel.map(pr => {
                                                              if (pr.id === p.id) {
                                                                const updatedHistory = pr.leaveHistory.map(lv => lv.id === l.id ? { ...lv, type: editingLeaveData.type, startDate: editingLeaveData.startDate, endDate: editingLeaveData.endDate, duration: newDuration, year: new Date(editingLeaveData.startDate).getFullYear(), description: editingLeaveData.description } : lv);
                                                                return { ...pr, leaveHistory: updatedHistory, leaveTotal: updatedHistory.reduce((s, lv) => s + lv.duration, 0) };
                                                              }
                                                              return pr;
                                                            });
                                                            setPersonnel(updatedPersonnel);
                                                            localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
                                                            setEditingLeaveId(null);
                                                            addToast('Başarılı', 'İzin kaydı güncellendi.', 'success');
                                                          }}
                                                          className="p-1 bg-emerald-100 text-emerald-700 hover:bg-emerald-200 rounded-md transition-colors"
                                                          title="Kaydet"
                                                        >
                                                          <CheckCircle2 size={14} />
                                                        </button>
                                                        <button onClick={() => setEditingLeaveId(null)} className="p-1 bg-slate-100 text-slate-500 hover:bg-slate-200 rounded-md transition-colors" title="İptal">
                                                          <X size={14} />
                                                        </button>
                                                      </div>
                                                    </td>
                                                  </>
                                                ) : (
                                                  <>
                                                    <td className="py-2 px-2">
                                                      <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-bold ${
                                                        l.type === 'Hastalık' ? 'bg-rose-100 text-rose-700' :
                                                        l.type === 'Heyet Raporu' ? 'bg-fuchsia-100 text-fuchsia-700' :
                                                        l.type === 'Refakat' ? 'bg-sky-100 text-sky-700' :
                                                        l.type === 'Yıllık' ? 'bg-emerald-100 text-emerald-700' :
                                                        l.type === 'Mazeret' ? 'bg-amber-100 text-amber-700' :
                                                        'bg-slate-100 text-slate-700'
                                                      }`}>
                                                        {l.type}
                                                      </span>
                                                    </td>
                                                    <td className="py-2 px-2 text-slate-700 text-xs">{formatDateTR(l.startDate)}</td>
                                                    <td className="py-2 px-2 text-slate-700 text-xs">{formatDateTR(l.endDate)}</td>
                                                    <td className="py-2 px-2 text-center font-bold text-slate-800 text-xs">{l.duration} gün</td>
                                                    <td className="py-2 px-2 text-slate-500 text-[11px]">{l.description || '—'}</td>
                                                    <td className="py-2 px-2 text-center">
                                                      <div className="flex items-center justify-center gap-1">
                                                        <button
                                                          onClick={() => {
                                                            setEditingLeaveId(l.id);
                                                            setEditingLeaveData({ type: l.type, startDate: l.startDate, endDate: l.endDate, description: l.description || '' });
                                                          }}
                                                          className="p-1 text-blue-600 hover:bg-blue-100 rounded-md transition-colors"
                                                          title="Düzenle"
                                                        >
                                                          ✏️
                                                        </button>
                                                        <button
                                                          onClick={() => {
                                                            if (confirm('Bu izin kaydını silmek istediğinizden emin misiniz?')) {
                                                              const updatedPersonnel = personnel.map(pr => {
                                                                if (pr.id === p.id) {
                                                                  const updatedHistory = pr.leaveHistory.filter(lv => lv.id !== l.id);
                                                                  return { ...pr, leaveHistory: updatedHistory, leaveTotal: updatedHistory.reduce((s, lv) => s + lv.duration, 0) };
                                                                }
                                                                return pr;
                                                              });
                                                              setPersonnel(updatedPersonnel);
                                                              localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
                                                              addToast('Başarılı', 'İzin kaydı silindi.', 'success');
                                                            }
                                                          }}
                                                          className="p-1 text-rose-500 hover:bg-rose-100 rounded-md transition-colors"
                                                          title="Sil"
                                                        >
                                                          🗑️
                                                        </button>
                                                      </div>
                                                    </td>
                                                  </>
                                                )}
                                              </tr>
                                            );
                                          })}
                                        </tbody>
                                        <tfoot>
                                          <tr className="border-t-2 border-blue-200">
                                            <td colSpan={3} className="py-2 px-2 text-[11px] font-bold text-slate-600 text-right">Toplam:</td>
                                            <td className="py-2 px-2 text-center font-bold text-blue-700 text-xs">{totalDays} gün</td>
                                            <td colSpan={2}></td>
                                          </tr>
                                        </tfoot>
                                      </table>
                                    )}
                                  </div>
                                )}
                              </div>
                            );
                          })}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'official' && (
            <div className="bg-white min-h-full p-10">
              <div className="no-print mb-8 bg-slate-50 p-6 rounded-2xl border border-slate-200 shadow-sm max-w-[1050px] mx-auto">
                <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">🤝 Sendika İşlemleri</h3>
                <div className="flex flex-wrap gap-4 items-end">
                  <div className="flex-1 min-w-[200px]">
                    <label className="block text-xs font-semibold text-slate-500 mb-1.5">Personel Seçin</label>
                    <select id="union-personnel" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                      <option value="">-- Seçiniz --</option>
                      {personnel.map(p => <option key={p.id} value={p.id}>{p.name} ({p.title})</option>)}
                    </select>
                  </div>
                  <div className="w-40">
                    <label className="block text-xs font-semibold text-slate-500 mb-1.5">İşlem Türü</label>
                    <select id="union-type" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                      <option value="Üye Olma">Üye Olma</option>
                      <option value="Ayrılma">Ayrılma</option>
                    </select>
                  </div>
                  <div className="flex-1 min-w-[200px]">
                    <label className="block text-xs font-semibold text-slate-500 mb-1.5">Sendika Adı</label>
                    <input type="text" id="union-name" placeholder="Örn: Eğitim-Bir-Sen" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                  </div>
                  <button onClick={() => {
                    const pId = (document.getElementById('union-personnel') as HTMLSelectElement).value;
                    const type = (document.getElementById('union-type') as HTMLSelectElement).value as 'Üye Olma' | 'Ayrılma';
                    const name = (document.getElementById('union-name') as HTMLInputElement).value;
                    if (!pId || !name) {
                      addToast("Hata", "Lütfen personel ve sendika adı giriniz.", "error");
                      return;
                    }
                    const newChange = {
                      id: Date.now(),
                      personnelId: Number(pId),
                      type,
                      unionName: name,
                      date: new Date().toISOString().split('T')[0]
                    };
                    const updated = [...unionChanges, newChange];
                    setUnionChanges(updated);
                    localStorage.setItem('pts_union_changes', JSON.stringify(updated));
                    addToast("Başarılı", "Sendika işlemi eklendi.", "success");
                    
                    // Personel kaydını da güncelle
                    if (type === 'Üye Olma') {
                      const updatedPersonnel = personnel.map(p => p.id === Number(pId) ? { ...p, union: name } : p);
                      setPersonnel(updatedPersonnel);
                      localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
                    } else {
                      const updatedPersonnel = personnel.map(p => p.id === Number(pId) ? { ...p, union: '' } : p);
                      setPersonnel(updatedPersonnel);
                      localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
                    }
                    
                    (document.getElementById('union-personnel') as HTMLSelectElement).value = '';
                    (document.getElementById('union-name') as HTMLInputElement).value = '';
                  }} className="px-6 py-2.5 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-semibold transition-colors shadow-sm shadow-blue-600/20 h-[42px]">
                    Ekle
                  </button>
                </div>
                {unionChanges.length > 0 && (
                  <div className="mt-4 flex flex-wrap gap-2">
                    {unionChanges.map(uc => {
                      const p = personnel.find(x => x.id === uc.personnelId);
                      return (
                        <div key={uc.id} className="inline-flex items-center gap-2 bg-white border border-slate-200 px-3 py-1.5 rounded-lg text-xs">
                          <span className="font-semibold">{p?.name}</span>
                          <span className={uc.type === 'Üye Olma' ? 'text-emerald-600' : 'text-rose-600'}>{uc.type}</span>
                          <span className="text-slate-500">({uc.unionName})</span>
                          <button onClick={() => {
                            const updated = unionChanges.filter(x => x.id !== uc.id);
                            setUnionChanges(updated);
                            localStorage.setItem('pts_union_changes', JSON.stringify(updated));
                          }} className="text-slate-400 hover:text-red-500 ml-1"><X size={14} /></button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="no-print mb-8 bg-slate-50 p-6 rounded-2xl border border-slate-200 shadow-sm max-w-[1050px] mx-auto">
                <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">📝 Diğer Değişiklikler (Aile Durum Bildirimi, Dil Tazminatı, Kefalet, Aile Yardım Beyanı, Sakatlık İndirimi)</h3>
                <div className="flex flex-wrap gap-4 items-end">
                  <div className="flex-1 min-w-[200px]">
                    <label className="block text-xs font-semibold text-slate-500 mb-1.5">Personel Seçin</label>
                    <select id="other-personnel" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                      <option value="">-- Seçiniz --</option>
                      {personnel.map(p => <option key={p.id} value={p.id}>{p.name} ({p.title})</option>)}
                    </select>
                  </div>
                  <div className="w-56">
                    <label className="block text-xs font-semibold text-slate-500 mb-1.5">İşlem Türü</label>
                    <select id="other-process-type" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                      <option value="Aile Durum Bildirimi">Aile Durum Bildirimi</option>
                      <option value="Aile Yardım Beyanı">Aile Yardım Beyanı</option>
                      <option value="Dil Tazminatı">Dil Tazminatı</option>
                      <option value="Kefalet">Kefalet</option>
                      <option value="Sakatlık İndirimi">Sakatlık İndirimi</option>
                      <option value="Diğer">Diğer</option>
                    </select>
                  </div>
                  <div className="w-32">
                    <label className="block text-xs font-semibold text-slate-500 mb-1.5">Miktar/Oran</label>
                    <input type="text" id="other-amount" placeholder="Örn: 1 çocuk, %10" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                  </div>
                  <div className="flex-1 min-w-[180px]">
                    <label className="block text-xs font-semibold text-slate-500 mb-1.5">Açıklama</label>
                    <input type="text" id="other-description" placeholder="Ek açıklama (opsiyonel)" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                  </div>
                  <button onClick={() => {
                    const pId = (document.getElementById('other-personnel') as HTMLSelectElement).value;
                    const processType = (document.getElementById('other-process-type') as HTMLSelectElement).value;
                    const amount = (document.getElementById('other-amount') as HTMLInputElement).value;
                    const description = (document.getElementById('other-description') as HTMLInputElement).value;
                    if (!pId) {
                      addToast("Hata", "Lütfen personel seçiniz.", "error");
                      return;
                    }
                    const newChange = {
                      id: Date.now(),
                      personnelId: Number(pId),
                      processType,
                      amount,
                      description,
                      date: new Date().toISOString().split('T')[0]
                    };
                    const updated = [...otherChanges, newChange];
                    setOtherChanges(updated);
                    localStorage.setItem('pts_other_changes', JSON.stringify(updated));
                    addToast("Başarılı", "Değişiklik eklendi. Resmi Çizelgede 'Diğer Değişiklikler' bölümüne yansıtıldı.", "success");

                    (document.getElementById('other-personnel') as HTMLSelectElement).value = '';
                    (document.getElementById('other-amount') as HTMLInputElement).value = '';
                    (document.getElementById('other-description') as HTMLInputElement).value = '';
                  }} className="px-6 py-2.5 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-semibold transition-colors shadow-sm shadow-blue-600/20 h-[42px]">
                    Ekle
                  </button>
                </div>
                {otherChanges.length > 0 && (
                  <div className="mt-4 flex flex-wrap gap-2">
                    {otherChanges.map(oc => {
                      const p = personnel.find(x => x.id === oc.personnelId);
                      return (
                        <div key={oc.id} className="inline-flex items-center gap-2 bg-white border border-slate-200 px-3 py-1.5 rounded-lg text-xs">
                          <span className="font-semibold">{p?.name}</span>
                          <span className="text-blue-600">{oc.processType}</span>
                          {oc.amount && <span className="text-slate-500">({oc.amount})</span>}
                          <button onClick={() => {
                            const updated = otherChanges.filter(x => x.id !== oc.id);
                            setOtherChanges(updated);
                            localStorage.setItem('pts_other_changes', JSON.stringify(updated));
                          }} className="text-slate-400 hover:text-red-500 ml-1"><X size={14} /></button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="no-print mb-8 bg-slate-50 p-6 rounded-2xl border border-slate-200 shadow-sm max-w-[1050px] mx-auto">
                <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">🏢 Görev Durumu Değişiklikleri</h3>
                <div className="flex flex-wrap gap-4 items-end">
                  <div className="flex-1 min-w-[200px]">
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">Personel Seçimi</label>
                    <select id="duty-personnel" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                      <option value="">-- Listeden Personel Seçiniz --</option>
                      {personnel.map(p => <option key={p.id} value={p.id}>{p.name} ({p.title})</option>)}
                    </select>
                  </div>
                  <div className="w-40">
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">İşlem Türü</label>
                    <select id="duty-type" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" onChange={(e) => {
                      const val = e.target.value;
                      const isAyrilma = val === 'Görevden Ayrılma';
                      const isUcretsiz = val === 'Aylıksız İzin';
                      const destEl = document.getElementById('duty-destination');
                      const descEl = document.getElementById('duty-description');
                      if (destEl) destEl.style.display = isAyrilma ? 'block' : 'none';
                      if (descEl) descEl.style.display = (isAyrilma || isUcretsiz) ? 'block' : 'none';
                    }}>
                      <option value="Göreve Başlama">Göreve Başlama</option>
                      <option value="Görevden Ayrılma">Görevden Ayrılma</option>
                      <option value="Aylıksız İzin">Aylıksız İzin</option>
                    </select>
                  </div>
                  <div className="w-40">
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">İşlem Tarihi</label>
                    <input type="date" id="duty-date" defaultValue={new Date().toISOString().split('T')[0]} className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                  </div>
                  <div id="duty-destination" className="flex-1 min-w-[150px]" style={{display: 'none'}}>
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">Gittiği Kurum / Yer</label>
                    <input type="text" id="duty-dest-input" placeholder="Örn: X İlçe Milli Eğitim Müdürlüğü" className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                  </div>
                  <div id="duty-description" className="flex-1 min-w-[150px]" style={{display: 'none'}}>
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">Açıklama / Süre</label>
                    <input type="text" id="duty-desc-input" placeholder="Örn: Doğum sonrası 1 yıl, Naklen Atama vb." className="w-full px-3 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                  </div>
                  <button onClick={() => {
                    const pId = (document.getElementById('duty-personnel') as HTMLSelectElement).value;
                    const type = (document.getElementById('duty-type') as HTMLSelectElement).value as 'Göreve Başlama' | 'Görevden Ayrılma' | 'Aylıksız İzin';
                    const date = (document.getElementById('duty-date') as HTMLInputElement).value;
                    const dest = (document.getElementById('duty-dest-input') as HTMLInputElement).value;
                    const desc = (document.getElementById('duty-desc-input') as HTMLInputElement).value;
                    
                    if (!pId || !date) {
                      addToast("Hata", "Lütfen personel ve tarih giriniz.", "error");
                      return;
                    }
                    const newChange = {
                      id: Date.now(),
                      personnelId: Number(pId),
                      type,
                      date: date.split('-').reverse().join('.'),
                      destination: type === 'Görevden Ayrılma' ? dest : undefined,
                      description: type === 'Görevden Ayrılma' ? desc : (type === 'Aylıksız İzin' ? (desc || 'Aylıksız İzin') : undefined)
                    };
                    const updated = [...dutyChanges, newChange];
                    setDutyChanges(updated);
                    localStorage.setItem('pts_duty_changes', JSON.stringify(updated));
                    addToast("Başarılı", "Görev durumu değişikliği eklendi.", "success");
                    
                    (document.getElementById('duty-personnel') as HTMLSelectElement).value = '';
                    (document.getElementById('duty-dest-input') as HTMLInputElement).value = '';
                    (document.getElementById('duty-desc-input') as HTMLInputElement).value = '';
                  }} className="px-6 py-2.5 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-semibold transition-colors shadow-sm shadow-blue-600/20 h-[42px]">
                    Ekle
                  </button>
                </div>
                {dutyChanges.length > 0 && (
                  <div className="mt-4 flex flex-wrap gap-2">
                    {dutyChanges.map(dc => {
                      const p = personnel.find(x => x.id === dc.personnelId);
                      return (
                        <div key={dc.id} className="inline-flex items-center gap-2 bg-white border border-slate-200 px-3 py-1.5 rounded-lg text-xs">
                          <span className="font-semibold">{p?.name}</span>
                          <span className={
                            dc.type === 'Göreve Başlama' ? 'text-emerald-600' :
                            dc.type === 'Aylıksız İzin' ? 'text-amber-600' :
                            'text-rose-600'
                          }>{dc.type}</span>
                          <span className="text-slate-500">({formatDateTR(dc.date)})</span>
                          <button onClick={() => {
                            const updated = dutyChanges.filter(x => x.id !== dc.id);
                            setDutyChanges(updated);
                            localStorage.setItem('pts_duty_changes', JSON.stringify(updated));
                          }} className="text-slate-400 hover:text-red-500 ml-1"><X size={14} /></button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="no-print text-right mb-6 max-w-[1050px] mx-auto">
                <div className="flex justify-between items-center mb-4 gap-4">
                  <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-2">
                    <span>📄 Resmi Formlara Manuel Satır Ekle</span>
                  </h3>
                  <div className="flex flex-wrap gap-2 items-end text-[11px]">
                    <select
                      id="promo-personnel"
                      className="px-2 py-1.5 rounded-lg bg-white border border-slate-300"
                      onChange={(e) => {
                        const val = (e.target as HTMLSelectElement).value;
                        const saymanEl = document.getElementById('promo-sayman') as HTMLInputElement | null;
                        const nameEl = document.getElementById('promo-name') as HTMLInputElement | null;
                        const tcEl = document.getElementById('promo-tc') as HTMLInputElement | null;
                        const degEl = document.getElementById('promo-degree') as HTMLInputElement | null;
                        const levEl = document.getElementById('promo-level') as HTMLInputElement | null;
                        const newDegEl = document.getElementById('promo-new-degree') as HTMLInputElement | null;
                        const newLevEl = document.getElementById('promo-new-level') as HTMLInputElement | null;

                        const clear = () => {
                          if (saymanEl) saymanEl.value = '';
                          if (nameEl) nameEl.value = '';
                          if (tcEl) tcEl.value = '';
                          if (degEl) degEl.value = '';
                          if (levEl) levEl.value = '';
                          if (newDegEl) newDegEl.value = '';
                          if (newLevEl) newLevEl.value = '';
                        };

                        if (!val) return clear();
                        const p = personnel.find(x => x.id === Number(val));
                        if (!p) return clear();

                        if (saymanEl) saymanEl.value = String(p.personnelNo ?? '');
                        if (nameEl) nameEl.value = String(p.name ?? '');
                        if (tcEl) tcEl.value = String(p.tc ?? '');
                        if (degEl) degEl.value = String(p.degree ?? '');
                        if (levEl) levEl.value = String(p.level ?? '');

                        // Yeni derece/kademe otomatik hesap
                        let nD = p.degree ?? 0;
                        let nL = (p.level ?? 0) + 1;
                        // Derece 1 için kademe 4'e kadar çıkabilir
                        if (nD === 1) {
                          if (nL > 4) { nL = 4; }
                        } else {
                          if (nL > 3) { nL = 1; nD = Math.max(1, nD - 1); }
                        }
                        if (newDegEl) newDegEl.value = String(nD || '');
                        if (newLevEl) newLevEl.value = String(nL || '');
                      }}
                    >
                      <option value="">Personel Seç (opsiyonel)</option>
                      {personnel.map(p => (
                        <option key={p.id} value={p.id}>
                          {p.name} ({p.title})
                        </option>
                      ))}
                    </select>
                    <input id="promo-sayman" placeholder="Saymanlık No" className="px-2 py-1.5 rounded-lg bg-white border border-slate-300 w-24" />
                    <input id="promo-name" placeholder="Adı Soyadı" className="px-2 py-1.5 rounded-lg bg-white border border-slate-300 w-40" />
                    <input id="promo-tc" placeholder="T.C. Kimlik" className="px-2 py-1.5 rounded-lg bg-white border border-slate-300 w-28" />
                    <input id="promo-degree" type="number" placeholder="Derc." className="px-2 py-1.5 rounded-lg bg-white border border-slate-300 w-16" />
                    <input id="promo-level" type="number" placeholder="Kad." className="px-2 py-1.5 rounded-lg bg-white border border-slate-300 w-16" />
                    <input id="promo-new-degree" type="number" placeholder="Yeni Derc." className="px-2 py-1.5 rounded-lg bg-white border border-slate-300 w-20" />
                    <input id="promo-new-level" type="number" placeholder="Yeni Kad." className="px-2 py-1.5 rounded-lg bg-white border border-slate-300 w-20" />
                    <input
                      id="promo-date"
                      type="date"
                      defaultValue={new Date().toISOString().split('T')[0]}
                      className="px-2 py-1.5 rounded-lg bg-white border border-slate-300"
                    />
                    <input id="promo-desc" placeholder="Açıklama" className="px-2 py-1.5 rounded-lg bg-white border border-slate-300 w-40" />
                    <button
                      onClick={() => {
                        const selId = (document.getElementById('promo-personnel') as HTMLSelectElement | null)?.value;
                        const saymanEl = document.getElementById('promo-sayman') as HTMLInputElement | null;
                        const nameEl = document.getElementById('promo-name') as HTMLInputElement | null;
                        const tcEl = document.getElementById('promo-tc') as HTMLInputElement | null;
                        const degEl = document.getElementById('promo-degree') as HTMLInputElement | null;
                        const levEl = document.getElementById('promo-level') as HTMLInputElement | null;
                        const newDegEl = document.getElementById('promo-new-degree') as HTMLInputElement | null;
                        const newLevEl = document.getElementById('promo-new-level') as HTMLInputElement | null;
                        const dateEl = document.getElementById('promo-date') as HTMLInputElement | null;
                        const descEl = document.getElementById('promo-desc') as HTMLInputElement | null;

                        const selectedPerson = selId ? personnel.find(p => p.id === Number(selId)) : undefined;
                        const personnelNo = (saymanEl?.value || '').trim() || (selectedPerson?.personnelNo ?? '');
                        const name = (nameEl?.value || '').trim() || (selectedPerson?.name ?? '');
                        const tc = (tcEl?.value || '').trim() || (selectedPerson?.tc ?? '');
                        const oldDegree = Number(degEl?.value || selectedPerson?.degree || 0);
                        const oldLevel = Number(levEl?.value || selectedPerson?.level || 0);
                        const newDegree = Number(newDegEl?.value || oldDegree || 0);
                        const newLevel = Number(newLevEl?.value || oldLevel || 0);
                        const rawDate = (dateEl?.value || '').trim();
                        const description = (descEl?.value || '').trim() || 'kademe';

                        if (!name || !rawDate) {
                          addToast('Hata', 'En azından Adı Soyadı ve Geçerlilik Tarihi girilmelidir.', 'error');
                          return;
                        }

                        const displayDate = rawDate ? rawDate.split('-').reverse().join('.') : '';
                        const newRow = {
                          id: Date.now(),
                          personnelNo,
                          name,
                          tc,
                          education: selectedPerson?.education,
                          title: selectedPerson?.title,
                          branch: selectedPerson?.branch,
                          oldDegree,
                          oldLevel,
                          newDegree,
                          newLevel,
                          date: displayDate,
                          description,
                        };
                        const updated = [...manualPromotions, newRow];
                        setManualPromotions(updated);
                        localStorage.setItem('pts_manual_promotions', JSON.stringify(updated));
                        addToast('Başarılı', 'Terfi satırı eklendi.', 'success');

                        if (saymanEl) saymanEl.value = '';
                        if (nameEl) nameEl.value = '';
                        if (tcEl) tcEl.value = '';
                        if (degEl) degEl.value = '';
                        if (levEl) levEl.value = '';
                        if (newDegEl) newDegEl.value = '';
                        if (newLevEl) newLevEl.value = '';
                        if (descEl) descEl.value = '';
                      }}
                      className="px-4 py-1.5 bg-blue-600 text-white rounded-lg text-xs font-semibold hover:bg-blue-700 transition-colors"
                    >
                      Satır Ekle
                    </button>
                  </div>
                </div>
                <div className="no-print flex justify-between items-center mb-4">
                  <div className="inline-flex bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                    <button
                      onClick={() => setOfficialForm('salaryChange')}
                      className={`px-4 py-2 text-xs font-semibold transition-colors ${officialForm === 'salaryChange' ? 'bg-slate-900 text-white' : 'bg-white text-slate-700 hover:bg-slate-50'}`}
                    >
                      Maaş Değişikliği Formu
                    </button>
                    <button
                      onClick={() => setOfficialForm('stepPromotion')}
                      className={`px-4 py-2 text-xs font-semibold transition-colors ${officialForm === 'stepPromotion' ? 'bg-slate-900 text-white' : 'bg-white text-slate-700 hover:bg-slate-50'}`}
                    >
                      Kademe Terfi Formu (EK-1)
                    </button>
                  </div>
                  <div className="flex gap-3">
                    <button onClick={() => setShowPreviewModal(true)} className="px-6 py-3 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-bold inline-flex items-center gap-2 transition-colors shadow-sm">
                      👁️ Önizleme
                    </button>
                    <button 
                      onClick={async () => {
                        if (pdfLoading) return;
                        const orientation: 'portrait' | 'landscape' = officialForm === 'stepPromotion' ? 'landscape' : 'portrait';
                        const filename = officialForm === 'stepPromotion' ? 'Kademe_Terfi_Formu.pdf' : 'Maas_Degisikligi_Formu.pdf';

                        setPdfLoading(true);
                        document.body.classList.add('printing');
                        const cleanup = applyPrintOrientation(orientation);
                        // CSS güncellemesi için kısa bekleme
                        await new Promise(resolve => setTimeout(resolve, 150));

                        try {
                          if (window.printer) {
                            const res = await window.printer.savePDF({ orientation, filename });
                            if (res.success) {
                              addToast('PDF Kaydedildi', `Dosya: ${res.path}`, 'success');
                            } else if (!res.canceled) {
                              addToast('Hata', res.error || 'PDF oluşturulamadı.', 'error');
                            }
                          } else {
                            // Tarayıcı/dev ortamı: window.print fallback'i
                            window.print();
                          }
                        } catch (error: any) {
                          console.error('PDF oluşturma hatası:', error);
                          addToast('Hata', error?.message || 'PDF oluşturulurken bir hata oluştu.', 'error');
                        } finally {
                          cleanup();
                          document.body.classList.remove('printing');
                          setPdfLoading(false);
                        }
                      }} 
                      disabled={pdfLoading}
                      className="px-6 py-3 bg-emerald-600 text-white hover:bg-emerald-700 rounded-xl font-bold inline-flex items-center gap-2 transition-colors shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      {pdfLoading ? '⏳ PDF Hazırlanıyor...' : '📥 PDF İndir'}
                    </button>
                    <button 
                      onClick={() => setShowPreviewModal(true)} 
                      className="px-8 py-3 bg-slate-800 text-white hover:bg-slate-900 rounded-xl font-bold inline-flex items-center gap-2 transition-colors shadow-sm"
                      title="Yazdırmadan önce önizleme açılır"
                    >
                      🖨 Yazdır
                    </button>
                  </div>
                </div>
              </div>

              {officialForm === 'salaryChange' && (
                <div className="w-full overflow-x-auto">
                  <div className="border-2 border-black p-6 text-black font-sans min-w-[800px] max-w-[1050px] mx-auto text-[10px] leading-tight print-portrait">
                   <div className="text-center font-bold text-[11px] mb-3 uppercase">{settings.provinceTitle || 'MAMAK MÜDÜRLÜĞÜNE'}</div>
                   <div className="mb-3">
                      <div className="text-[9px]"><span className="font-bold">Kurumun Adı:</span> {settings.schoolTitle}</div>
                      <div className="text-[9px] mt-1"><span className="font-bold">Ayı Yılı:</span> {MONTHS[new Date().getMonth()]} {new Date().getFullYear().toString().slice(-2)}</div>
                   </div>
                   <div className="bg-yellow-300 border border-black text-center font-bold p-1 text-[10px] uppercase">Terfiler</div>
                   <table className="w-full border-collapse border border-black text-[9px]">
                      <thead>
                        <tr className="bg-white">
                          <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                          <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                          <th colSpan={3} className="border border-black p-1 font-bold text-center">Yeni</th>
                          <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Yeni Özel Hizmet Tazm.</th>
                          <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Geçerlilik Tarihi</th>
                          <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Açıklama</th>
                          <th rowSpan={2} className="border border-black p-1 no-print bg-rose-50">İşlem</th>
                        </tr>
                        <tr className="bg-white">
                          <th className="border border-black p-1 font-bold text-center">Derece</th>
                          <th className="border border-black p-1 font-bold text-center">Kademe</th>
                          <th className="border border-black p-1 font-bold text-center">Ek Göst.</th>
                        </tr>
                      </thead>
                      <tbody>
                         {periodPromotions.length + manualPromotions.length === 0 ? (
                           <tr>
                             <td colSpan={9} className="border border-black p-4 text-center text-slate-500">
                               - Maaş Döneminde (15-14) Terfi Kaydı Bulunmamaktadır -
                             </td>
                           </tr>
                         ) : (
                           <>
                             {periodPromotions.map((item, idx) => {
                              const dateKey = `promo-${item.person.id}`;
                              const displayDate = promotionDateOverrides[dateKey] || item.date;
                              return (
                              <tr key={`auto-${idx}`}>
                                <td className="border border-black p-1 text-center">{item.person.personnelNo}</td>
                                <td className="border border-black p-1">{item.person.name}</td>
                                <td className="border border-black p-1 text-center">{item.nextDegree}</td>
                                <td className="border border-black p-1 text-center">{item.nextLevel}</td>
                                <td className="border border-black p-1"></td>
                                <td className="border border-black p-1"></td>
                                <td className="border border-black p-1 text-center">
                                  <span 
                                    contentEditable
                                    suppressContentEditableWarning
                                    onBlur={(e) => {
                                      const newDate = e.currentTarget.textContent?.trim() || item.date;
                                      const updated = {...promotionDateOverrides, [dateKey]: newDate};
                                      setPromotionDateOverrides(updated);
                                      localStorage.setItem('pts_promotion_date_overrides', JSON.stringify(updated));
                                    }}
                                    className="cursor-text hover:bg-yellow-100 px-1 rounded"
                                  >
                                    {displayDate}
                                  </span>
                                </td>
                                <td className="border border-black p-1">kademe</td>
                                <td className="border border-black p-1 no-print text-center">
                                  <div className="flex gap-1 justify-center">
                                    <button 
                                      onClick={() => {
                                        // Terfi tarihini personel kaydına kaydet
                                        const updatedPersonnel = personnel.map(p => 
                                          p.id === item.person.id 
                                            ? {...p, lastPromotionDate: displayDate, degree: item.nextDegree, level: item.nextLevel} 
                                            : p
                                        );
                                        setPersonnel(updatedPersonnel);
                                        localStorage.setItem('pts_personnel', JSON.stringify(updatedPersonnel));
                                        addToast('Başarılı', 'Terfi kaydedildi ve personel bilgileri güncellendi.', 'success');
                                      }}
                                      className="text-green-600 hover:text-green-800 font-bold text-xs"
                                      title="Terfiyi Kaydet ve Uygula"
                                    >
                                      ✓
                                    </button>
                                    <button 
                                      onClick={() => {
                                        const updated = [...excludedAutoPromotions, item.person.id];
                                        setExcludedAutoPromotions(updated);
                                        localStorage.setItem('pts_excluded_auto_promotions', JSON.stringify(updated));
                                        addToast('Başarılı', 'Otomatik terfi çizelgeden kaldırıldı.', 'success');
                                      }}
                                      className="text-rose-600 hover:text-rose-800 font-bold text-xs"
                                      title="Çizelgeden Kaldır"
                                    >
                                      ✕
                                    </button>
                                  </div>
                                </td>
                              </tr>
                              );
                            })}
                             {manualPromotions.map(row => (
                               <tr key={`manual-${row.id}`}>
                                 <td className="border border-black p-1 text-center">{row.personnelNo}</td>
                                 <td className="border border-black p-1">{row.name}</td>
                                 <td className="border border-black p-1 text-center">{row.newDegree}</td>
                                 <td className="border border-black p-1 text-center">{row.newLevel}</td>
                                 <td className="border border-black p-1"></td>
                                 <td className="border border-black p-1"></td>
                                 <td className="border border-black p-1 text-center">{row.date}</td>
                                 <td className="border border-black p-1">{row.description}</td>
                                 <td className="border border-black p-1 no-print text-center">
                                   <button 
                                     onClick={() => {
                                       const updated = manualPromotions.filter(x => x.id !== row.id);
                                       setManualPromotions(updated);
                                       localStorage.setItem('pts_manual_promotions', JSON.stringify(updated));
                                       addToast('Başarılı', 'Manuel terfi kaydı silindi.', 'success');
                                     }}
                                     className="text-rose-600 hover:text-rose-800 font-bold text-xs"
                                     title="Sil"
                                   >
                                     ✕
                                   </button>
                                 </td>
                               </tr>
                             ))}
                             {Array.from({ length: Math.max(0, 5 - (periodPromotions.length + manualPromotions.length)) }).map((_, i) => (
                               <tr key={`empty-${i}`}>
                                 <td colSpan={9} className="border border-black p-2"></td>
                               </tr>
                             ))}
                           </>
                         )}
                      </tbody>
                   </table>
                 <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">Göreve Başlama</div>
                 <table className="w-full border-collapse border border-black text-[9px]">
                    <thead>
                       <tr className="bg-white">
                          <th className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Derece</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Kademe</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Özel Hiz. Taz.</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Yan Ödeme</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Ek Ödeme Oranı</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">İban No</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Başlama Tarihi</th>
                       </tr>
                    </thead>
                    <tbody>
                      {periodDutyStarters.length === 0 ? (
                        Array.from({length: 3}).map((_, i) => (<tr key={i}><td colSpan={9} className="border border-black p-2"></td></tr>))
                      ) : (
                        periodDutyStarters.map((item, idx) => (
                          <tr key={idx}>
                            <td className="border border-black p-1 text-center">{item.person.personnelNo}</td>
                            <td className="border border-black p-1">{item.person.name}</td>
                            <td className="border border-black p-1 text-center">{item.degree}</td>
                            <td className="border border-black p-1 text-center">{item.level}</td>
                            <td className="border border-black p-1"></td>
                            <td className="border border-black p-1"></td>
                            <td className="border border-black p-1"></td>
                            <td className="border border-black p-1 text-[8px]">{item.person.iban}</td>
                            <td className="border border-black p-1 text-center">{item.date}</td>
                          </tr>
                        ))
                      )}
                      {Array.from({length: Math.max(0, 3 - periodDutyStarters.length)}).map((_, i) => (
                        <tr key={`empty-${i}`}><td colSpan={9} className="border border-black p-2"></td></tr>
                      ))}
                    </tbody>
                 </table>
                 <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">Görevden Ayrılma (Naklen Atama, Emekli, Aylıksız İzin, İhraç, Açığa Alınma)</div>
                 <table className="w-full border-collapse border border-black text-[9px]">
                    <thead>
                       <tr className="bg-white">
                          <th className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Ayrıldığı Tarih</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Gittiği Yer</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Açıklama</th>
                       </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const ayrilmaList = dutyChanges.filter(dc => dc.type === 'Görevden Ayrılma' || dc.type === 'Aylıksız İzin');
                        if (ayrilmaList.length === 0) {
                          return Array.from({length: 3}).map((_, i) => (<tr key={i}><td colSpan={5} className="border border-black p-2"></td></tr>));
                        }
                        return (
                          <>
                            {ayrilmaList.map(dc => {
                              const p = personnel.find(x => x.id === dc.personnelId);
                              return (
                                <tr key={dc.id}>
                                  <td className="border border-black p-1 text-center">{p?.personnelNo}</td>
                                  <td className="border border-black p-1">{p?.name}</td>
                                  <td className="border border-black p-1 text-center">{formatDateTR(dc.date)}</td>
                                  <td className="border border-black p-1">{dc.destination || (dc.type === 'Aylıksız İzin' ? '—' : '')}</td>
                                  <td className="border border-black p-1">{dc.description || (dc.type === 'Aylıksız İzin' ? 'Aylıksız İzin' : '')}</td>
                                </tr>
                              );
                            })}
                            {Array.from({length: Math.max(0, 3 - ayrilmaList.length)}).map((_, i) => (
                              <tr key={`empty-${i}`}><td colSpan={5} className="border border-black p-2"></td></tr>
                            ))}
                          </>
                        );
                      })()}
                    </tbody>
                 </table>
                <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">İcra, Nafaka, Kefalet Giriş Aidatı, Para Cezası, vb. Kesintisi</div>
                <table className="w-full border-collapse border border-black text-[9px]">
                   <thead>
                      <tr className="bg-white">
                         <th className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                         <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                         <th className="border border-black p-1 font-bold text-center align-middle">Türü</th>
                         <th className="border border-black p-1 font-bold text-center align-middle">Dosya No</th>
                         <th className="border border-black p-1 font-bold text-center align-middle">Tutar</th>
                         <th className="border border-black p-1 font-bold text-center align-middle">Açıklama</th>
                      </tr>
                   </thead>
                    <tbody>
                       {reportExceeders.map((item, idx) => (
                         <tr key={`report-${idx}`}>
                            <td className="border border-black p-1 text-center">{item.person.personnelNo}</td>
                            <td className="border border-black p-1">{item.person.name}</td>
                            <td className="border border-black p-1 text-center">Rapor</td>
                            <td className="border border-black p-1"></td>
                            <td className="border border-black p-1 text-center">{item.days} gün</td>
                            <td className="border border-black p-1">Kesinti</td>
                         </tr>
                       ))}
                       {Array.from({length: Math.max(0, 3 - reportExceeders.length)}).map((_, i) => (
                         <tr key={i}><td colSpan={6} className="border border-black p-2"></td></tr>
                       ))}
                    </tbody>
                 </table>
                 <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">Diğer Değişiklikler (Aile Durum Bildirimi, Dil Tazminatı, Kefalet, Aile Yardım Beyanı, Sakatlık İndirimi)</div>
                 <table className="w-full border-collapse border border-black text-[9px]">
                    <thead>
                       <tr className="bg-white">
                          <th className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Yapılacak İşlemin</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Miktar/Oran</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Açıklama</th>
                       </tr>
                    </thead>
                    <tbody>
                       {otherChanges.map(oc => {
                         const p = personnel.find(x => x.id === oc.personnelId);
                         return (
                           <tr key={oc.id}>
                             <td className="border border-black p-1 text-center">{p?.personnelNo}</td>
                             <td className="border border-black p-1">{p?.name}</td>
                             <td className="border border-black p-1 text-center">{oc.processType}</td>
                             <td className="border border-black p-1 text-center">{oc.amount || '-'}</td>
                             <td className="border border-black p-1">{oc.description || '-'}</td>
                           </tr>
                         );
                       })}
                       {Array.from({length: Math.max(0, 3 - otherChanges.length)}).map((_, i) => (
                         <tr key={`empty-${i}`}><td colSpan={5} className="border border-black p-2"></td></tr>
                       ))}
                    </tbody>
                 </table>
                 <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">Sendika Değişiklikleri</div>
                 <table className="w-full border-collapse border border-black text-[9px]">
                    <thead>
                       <tr className="bg-white">
                          <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">İşlem Türü</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Sendika Adı</th>
                          <th className="border border-black p-1 font-bold text-center align-middle">Tarih</th>
                       </tr>
                    </thead>
                    <tbody>
                      {unionChanges.length === 0 ? (
                        <tr><td colSpan={4} className="border border-black p-3 text-center text-slate-500">Değişiklik Yok</td></tr>
                      ) : unionChanges.map(uc => {
                        const p = personnel.find(x => x.id === uc.personnelId);
                        return (
                          <tr key={uc.id}>
                            <td className="border border-black p-1.5">{p?.name}</td>
                            <td className="border border-black p-1.5 text-center">{uc.type}</td>
                            <td className="border border-black p-1.5 text-center">{uc.unionName}</td>
                            <td className="border border-black p-1.5 text-center">{formatDateTR(uc.date)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                 </table>
                 <div className="mt-12 flex justify-end pr-16">
                    <div className="text-center">
                       <div>{new Date().toLocaleDateString('tr-TR')}</div>
                       <div className="mt-12 font-bold text-[15px]">{settings.principalName}</div>
                       <div className="text-xs">Okul Müdürü</div>
                    </div>
                 </div>
                </div>
              </div>
              )}

              {officialForm === 'stepPromotion' && (
                <div className="w-full overflow-x-auto">
                  <div className="border-2 border-black p-4 text-black font-sans min-w-[1200px] max-w-[1400px] mx-auto text-[9px] leading-tight print-landscape">
                  <div className="text-center font-bold text-[11px] mb-2 uppercase">Milli Eğitim Bakanlığı Personeli (Branş) Öğretmenlerine Ait Kademe Terfi Onayı</div>
                  <div className="flex justify-between items-center font-bold mb-2 text-[9px]">
                    <div>İL : ANKARA</div>
                    <div className="text-right">EK-1 FORM</div>
                  </div>
                  <div className="mb-2 text-[9px]"><span className="font-bold">KURUM :</span> {settings.schoolTitle}</div>
                  <table className="w-full border-collapse border border-black text-[8px]">
                    <thead>
                      <tr className="bg-white">
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Sıra No</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">T.C.Kimlik No</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Görev Yaptığı Okul veya Kurum</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Adı ve Soyadı</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Sınıfı</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Mezuniyet</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Unvanı</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Branşı</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Kadro Derecesi</th>
                        <th colSpan={3} className="border border-black p-0.5 font-bold text-center">ESKİ DURUMU</th>
                        <th colSpan={3} className="border border-black p-0.5 font-bold text-center">YENİ DURUMU</th>
                        <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">AÇIKLAMA</th>
                      </tr>
                      <tr className="bg-white">
                        <th className="border border-black p-0.5 font-bold text-center">Maaş Derecesi</th>
                        <th className="border border-black p-0.5 font-bold text-center">Kademesi</th>
                        <th className="border border-black p-0.5 font-bold text-center">Bu Kademeyi Aldığı Tarih</th>
                        <th className="border border-black p-0.5 font-bold text-center">Maaş Derecesi</th>
                        <th className="border border-black p-0.5 font-bold text-center">Kademesi</th>
                        <th className="border border-black p-0.5 font-bold text-center">Bu Kademeyi Aldığı Tarih</th>
                      </tr>
                    </thead>
                    <tbody>
                      {periodPromotions.length + manualPromotions.length === 0 ? (
                        <tr><td className="border border-black p-2 text-center text-slate-500" colSpan={16}>-</td></tr>
                      ) : (
                        <>
                          {periodPromotions.map((item, idx) => {
                            const dateKey = `promo-${item.person.id}`;
                            const oldDateKey = `old-promo-${item.person.id}`;
                            const displayDate = promotionDateOverrides[dateKey] || item.date;
                            const oldDisplayDate = oldPromotionDateOverrides[oldDateKey] || item.person.lastPromotionDate || '';
                            return (
                            <tr key={`ek1-auto-${idx}`}>
                              <td className="border border-black p-0.5 text-center text-[7px]">{idx + 1}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.person.tc}</td>
                              <td className="border border-black p-0.5 text-[7px]">{settings.schoolTitle}</td>
                              <td className="border border-black p-0.5 text-[7px]">{item.person.name}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]"></td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.person.education}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.person.title}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.person.branch}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.person.degree}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.person.degree}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.person.level}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">
                                <span 
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => {
                                    const newDate = e.currentTarget.textContent?.trim() || '';
                                    const updated = {...oldPromotionDateOverrides, [oldDateKey]: newDate};
                                    setOldPromotionDateOverrides(updated);
                                    localStorage.setItem('pts_old_promotion_date_overrides', JSON.stringify(updated));
                                  }}
                                  className="cursor-text hover:bg-yellow-100 px-1 rounded inline-block min-w-[60px]"
                                >
                                  {oldDisplayDate}
                                </span>
                              </td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.nextDegree}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{item.nextLevel}</td>
                              <td className="border border-black p-0.5 text-center text-[7px]">{displayDate}</td>
                              <td className="border border-black p-0.5 text-[7px]">Kademe</td>
                            </tr>
                            );
                          })}
                          {manualPromotions.map((row, i) => {
                            return (
                              <tr key={`ek1-manual-${row.id}`}>
                                <td className="border border-black p-0.5 text-center text-[7px]">{periodPromotions.length + i + 1}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.tc || ''}</td>
                                <td className="border border-black p-0.5 text-[7px]">{settings.schoolTitle}</td>
                                <td className="border border-black p-0.5 text-[7px]">{row.name}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]"></td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.education || ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.title || ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.branch || ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.oldDegree ?? ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.oldDegree ?? ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.oldLevel ?? ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">
                                  <span 
                                    contentEditable
                                    suppressContentEditableWarning
                                    onBlur={(e) => {
                                      const newDate = e.currentTarget.textContent?.trim() || '';
                                      const updated = manualPromotions.map(mp => 
                                        mp.id === row.id ? {...mp, oldPromotionDate: newDate} : mp
                                      );
                                      setManualPromotions(updated);
                                      localStorage.setItem('pts_manual_promotions', JSON.stringify(updated));
                                    }}
                                    className="cursor-text hover:bg-yellow-100 px-1 rounded inline-block min-w-[60px]"
                                  >
                                    {row.oldPromotionDate || ''}
                                  </span>
                                </td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.newDegree}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.newLevel}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.date}</td>
                                <td className="border border-black p-0.5 text-[7px]">{row.description}</td>
                              </tr>
                            );
                          })}
                        </>
                      )}
                    </tbody>
                  </table>
                  <div className="mt-3 text-[8px]">
                    <div className="mb-2">Yukarıda durumu belirtilen {periodPromotions.length + manualPromotions.length === 1 ? 'bir (1)' : `${periodPromotions.length + manualPromotions.length} (${['', 'bir', 'iki', 'üç', 'dört', 'beş', 'altı', 'yedi', 'sekiz', 'dokuz', 'on'][periodPromotions.length + manualPromotions.length] || periodPromotions.length + manualPromotions.length})`} öğretmenin 29/06/1984 tarih ve 18446 sayılı Resmi Gazetede yayınlanan 241 Sayılı Kanun Hükmündeki kararname gereğince kademe/derece terfisini tasviplerinize arz ederim.</div>
                    <div className="flex justify-between items-end mt-6">
                      <div></div>
                      <div className="text-right">
                        <div className="mb-8">{settings.principalName || 'Adı-Soyadı'}</div>
                        <div>Okul Müdürü</div>
                      </div>
                    </div>
                  </div>
                </div>
                </div>
              )}
            </div>
          )}

          {activeTab === 'org' && (
            <div className="max-w-7xl mx-auto">
               <div className="flex justify-between items-center mb-6">
                  <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2"><Network size={24} className="text-blue-500" /> Organizasyon Şeması</h2>
               </div>
               <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-10 overflow-auto min-h-[600px] flex flex-col items-center">
                 <div className="flex flex-col items-center min-w-[800px]">
                   {/* Principal */}
                   <div className="bg-slate-900 text-white p-4 rounded-xl shadow-md w-64 text-center border-b-4 border-blue-500 relative z-10">
                     <div className="font-bold text-lg">{settings.principalName}</div>
                     <div className="text-xs text-slate-400 mt-1 uppercase tracking-wider">Okul Müdürü</div>
                   </div>
                   
                   <div className="w-px h-8 bg-slate-300"></div>
                   
                   {/* Head Vice Principals */}
                   {settings.headVicePrincipals.some(n => n.trim() !== "") && (
                     <>
                       <div className="flex gap-10 justify-center relative">
                         {settings.headVicePrincipals.filter(n => n.trim() !== "").map((name, idx) => (
                           <div key={idx} className="relative flex flex-col items-center">
                             <div className="w-px h-5 bg-slate-300"></div>
                             <div className="bg-amber-500 text-white p-3 rounded-xl shadow-sm w-48 text-center border border-amber-600 relative z-10">
                               <div className="font-bold text-sm">{name}</div>
                               <div className="text-[10px] text-amber-100 mt-1 uppercase tracking-wider">Müdür Başyardımcısı</div>
                             </div>
                           </div>
                         ))}
                       </div>
                       <div className="w-px h-8 bg-slate-300"></div>
                     </>
                   )}

                   {/* Vice Principals */}
                   <div className="w-full max-w-3xl border-t-2 border-slate-300 flex justify-center gap-6 relative pt-6">
                     <div className="absolute top-0 left-1/2 -translate-x-1/2 w-px h-6 bg-slate-300"></div>
                     {settings.vicePrincipals.filter(n => n.trim() !== "").map((name, idx) => (
                        <div key={idx} className="bg-blue-50 border border-blue-200 p-3 rounded-xl shadow-sm w-48 text-center relative z-10">
                          <div className="font-bold text-slate-800 text-sm">{name}</div>
                          <div className="text-[10px] text-blue-600 mt-1 uppercase tracking-wider">Müdür Yardımcısı</div>
                        </div>
                     ))}
                   </div>
                   
                   {settings.vicePrincipals.filter(n => n.trim() !== "").length === 0 && settings.headVicePrincipals.filter(n => n.trim() !== "").length === 0 && (
                     <div className="text-slate-400 italic mt-5 text-sm">Henüz idari kadro ayarlanmamış. (Ayarlar sekmesinden güncelleyebilirsiniz)</div>
                   )}

                   <div className="w-px h-10 bg-slate-300 mt-6"></div>
                   
                   {/* Alt Kadrolar (Personel Listesinden) */}
                   <div className="flex justify-center gap-6 mt-5 flex-wrap max-w-5xl">
                     <div className="bg-slate-50 border border-slate-200 p-6 rounded-2xl w-56 text-center shadow-sm">
                       <div className="text-blue-500 flex justify-center mb-3"><Briefcase size={28} /></div>
                       <div className="text-3xl font-bold text-slate-800">{personnel.filter(p => p.role === 'Öğretmen').length}</div>
                       <div className="text-xs text-slate-500 font-bold uppercase tracking-wider mt-1">Öğretmen Kadrosu</div>
                     </div>
                     <div className="bg-slate-50 border border-slate-200 p-6 rounded-2xl w-56 text-center shadow-sm">
                       <div className="text-indigo-500 flex justify-center mb-3"><Users size={28} /></div>
                       <div className="text-3xl font-bold text-slate-800">{personnel.filter(p => ['Memur', 'VHKİ', 'Teknisyen'].includes(p.role) || ['Memur', 'VHKİ', 'Teknisyen'].includes(p.title)).length}</div>
                       <div className="text-xs text-slate-500 font-bold uppercase tracking-wider mt-1">İdari Personel</div>
                     </div>
                     <div className="bg-slate-50 border border-slate-200 p-6 rounded-2xl w-56 text-center shadow-sm">
                       <div className="text-emerald-500 flex justify-center mb-3"><Building size={28} /></div>
                       <div className="text-3xl font-bold text-slate-800">{personnel.filter(p => ['Hizmetli', 'Sürekli İşçi'].includes(p.role)).length}</div>
                       <div className="text-xs text-slate-500 font-bold uppercase tracking-wider mt-1">Yardımcı Hizmetler</div>
                     </div>
                   </div>

                  {/* Öğretmen Branşları (Branşa göre dallanmış) */}
                  {personnel.filter(p => p.role === 'Öğretmen').length > 0 && (
                     <>
                       <div className="w-px h-10 bg-slate-300 mt-6"></div>
                      <div className="w-full max-w-6xl border-t-2 border-slate-300 flex justify-center gap-4 relative pt-6 flex-wrap">
                         <div className="absolute top-0 left-1/2 -translate-x-1/2 w-px h-6 bg-slate-300"></div>
                        {Object.entries(
                          personnel
                            .filter(p => p.role === 'Öğretmen')
                            .reduce((acc, curr) => {
                              const branch = (curr.branch || '').trim() || 'Branş Belirtilmemiş';
                              (acc[branch] ||= []).push(curr);
                              return acc;
                            }, {} as Record<string, Personnel[]>)
                        )
                          .sort(([a], [b]) => a.localeCompare(b, 'tr-TR'))
                          .map(([branch, people]) => (
                            <div
                              key={branch}
                              className="bg-white border border-slate-200 p-4 rounded-2xl shadow-sm w-[260px] text-left relative z-10 hover:border-blue-300 transition-colors"
                            >
                              <div className="flex items-start justify-between gap-3">
                                <div className="font-bold text-slate-800 text-sm leading-tight">{branch}</div>
                                <div className="text-[11px] text-blue-700 font-extrabold bg-blue-50 border border-blue-200 px-2 py-0.5 rounded-full whitespace-nowrap">
                                  {people.length} kişi
                                </div>
                              </div>
                              <div className="mt-3 max-h-[180px] overflow-auto pr-1">
                                <ul className="space-y-1">
                                  {people
                                    .slice()
                                    .sort((x, y) => x.name.localeCompare(y.name, 'tr-TR'))
                                    .map(p => (
                                      <li key={p.id} className="text-[11px] text-slate-700 flex items-center justify-between gap-2">
                                        <span className="truncate">{p.name}</span>
                                        <span className="text-[10px] text-slate-400 whitespace-nowrap">{p.title}</span>
                                      </li>
                                    ))}
                                </ul>
                              </div>
                            </div>
                          ))}
                       </div>
                     </>
                   )}
                 </div>
               </div>
            </div>
          )}

          {activeTab === 'schedules' && (
            <div className="max-w-7xl mx-auto">
               <div className="flex justify-between items-center mb-6">
                  <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2"><Clock size={24} className="text-blue-500" /> Ders Programı ve Nöbet Çizelgeleri</h2>
               </div>
               
               <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                 <div className="lg:col-span-1 bg-white rounded-2xl border border-slate-100 shadow-sm p-6 h-fit">
                   <h3 className="text-lg font-bold text-slate-800 mb-4">Yeni Çizelge Yükle</h3>
                   <div className="space-y-4">
                     <div>
                       <label className="block text-xs font-semibold text-slate-500 mb-1.5">Personel Seçin</label>
                       <select id="schedule-personnel" className="w-full px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                         <option value="">-- Tüm Personel (Genel Çizelge) --</option>
                         {personnel.map(p => <option key={p.id} value={p.id}>{p.name} ({p.title})</option>)}
                       </select>
                     </div>
                     <div>
                       <label className="block text-xs font-semibold text-slate-500 mb-1.5">Çizelge Türü</label>
                       <select id="schedule-type" className="w-full px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                         <option value="Ders Programı">Ders Programı</option>
                         <option value="Nöbet Çizelgesi">Nöbet Çizelgesi</option>
                       </select>
                     </div>
                     <div>
                       <label className="block text-xs font-semibold text-slate-500 mb-1.5">Dönem / Açıklama</label>
                       <input type="text" id="schedule-term" placeholder="Örn: 2023-2024 Güz Dönemi" className="w-full px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                     </div>
                     <div>
                       <label className="block text-xs font-semibold text-slate-500 mb-1.5">Dosya Yükle (PDF, Görsel, Excel)</label>
                       <input type="file" id="schedule-file" className="w-full px-3 py-2 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-xs file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" />
                       <p className="text-[10px] text-slate-400 mt-1">Not: Tarayıcı belleği sınırlı olduğundan küçük boyutlu dosyalar yükleyin.</p>
                     </div>
                     <button onClick={() => {
                       const pId = (document.getElementById('schedule-personnel') as HTMLSelectElement).value;
                       const type = (document.getElementById('schedule-type') as HTMLSelectElement).value as 'Ders Programı' | 'Nöbet Çizelgesi';
                       const term = (document.getElementById('schedule-term') as HTMLInputElement).value;
                       const fileInput = document.getElementById('schedule-file') as HTMLInputElement;
                       const file = fileInput.files?.[0];

                       if (!term || !file) {
                         addToast("Hata", "Lütfen dönem/açıklama girin ve bir dosya seçin.", "error");
                         return;
                       }

                       const reader = new FileReader();
                       reader.onload = (e) => {
                         const fileData = e.target?.result as string;
                         const newSchedule: ScheduleRecord = {
                           id: Date.now(),
                           personnelId: pId ? Number(pId) : 0,
                           type,
                           term,
                           fileName: file.name,
                           fileData,
                           uploadDate: new Date().toLocaleDateString('tr-TR')
                         };
                         
                         const updated = [newSchedule, ...schedules];
                         setSchedules(updated);
                         localStorage.setItem('pts_schedules', JSON.stringify(updated));
                         addToast("Başarılı", "Çizelge başarıyla yüklendi.", "success");
                         
                         (document.getElementById('schedule-term') as HTMLInputElement).value = '';
                         fileInput.value = '';
                       };
                       reader.readAsDataURL(file);
                     }} className="w-full py-2.5 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-semibold transition-colors shadow-sm shadow-blue-600/20 mt-2">
                       Yükle ve Kaydet
                     </button>
                   </div>
                 </div>
                 
                 <div className="lg:col-span-2 bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden flex flex-col h-[600px]">
                   <div className="p-4 border-b border-slate-100 bg-slate-50 flex justify-between items-center">
                     <h3 className="font-bold text-slate-800">Yüklü Çizelgeler</h3>
                     <div className="flex gap-2">
                       <span className="px-3 py-1 bg-blue-100 text-blue-700 rounded-lg text-xs font-semibold">{schedules.filter(s => s.type === 'Ders Programı').length} Ders Programı</span>
                       <span className="px-3 py-1 bg-amber-100 text-amber-700 rounded-lg text-xs font-semibold">{schedules.filter(s => s.type === 'Nöbet Çizelgesi').length} Nöbet Çizelgesi</span>
                     </div>
                   </div>
                   <div className="flex-1 overflow-auto p-4">
                     {schedules.length === 0 ? (
                       <div className="h-full flex flex-col items-center justify-center text-slate-400">
                         <Clock size={48} className="mb-4 opacity-20" />
                         <p>Henüz yüklenmiş bir çizelge bulunmuyor.</p>
                       </div>
                     ) : (
                       <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                         {schedules.map(schedule => {
                           const person = personnel.find(p => p.id === schedule.personnelId);
                           return (
                             <div key={schedule.id} className="border border-slate-200 rounded-xl p-4 hover:border-blue-300 transition-colors bg-white shadow-sm flex flex-col">
                               <div className="flex justify-between items-start mb-3">
                                 <div className={`px-2.5 py-1 rounded-md text-[10px] font-bold uppercase tracking-wider ${schedule.type === 'Ders Programı' ? 'bg-blue-50 text-blue-600' : 'bg-amber-50 text-amber-600'}`}>
                                   {schedule.type}
                                 </div>
                                 <button onClick={() => {
                                   if (window.confirm('Bu çizelgeyi silmek istediğinize emin misiniz?')) {
                                     const updated = schedules.filter(s => s.id !== schedule.id);
                                     setSchedules(updated);
                                     localStorage.setItem('pts_schedules', JSON.stringify(updated));
                                     addToast("Silindi", "Çizelge silindi.", "info");
                                   }
                                 }} className="text-slate-400 hover:text-red-500 transition-colors"><Trash2 size={16} /></button>
                               </div>
                               <h4 className="font-bold text-slate-800 text-sm mb-1">{schedule.term}</h4>
                               <p className="text-xs text-slate-500 mb-3 flex items-center gap-1">
                                 <Users size={12} /> {person ? person.name : 'Genel Çizelge (Tüm Personel)'}
                               </p>
                               <div className="mt-auto pt-3 border-t border-slate-100 flex justify-between items-center">
                                 <span className="text-[10px] text-slate-400">{schedule.uploadDate}</span>
                                 <div className="flex gap-2">
                                   <button 
                                     onClick={() => {
                                       setPreviewSchedule(schedule);
                                       setShowSchedulePreview(true);
                                     }} 
                                     className="text-xs font-semibold text-emerald-600 hover:text-emerald-800 flex items-center gap-1"
                                   >
                                     👁️ Önizle
                                   </button>
                                   <button onClick={() => downloadFile(schedule.fileData, schedule.fileName)} className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1">
                                     <Download size={14} /> İndir
                                   </button>
                                 </div>
                               </div>
                             </div>
                           );
                         })}
                       </div>
                     )}
                   </div>
                 </div>
               </div>
            </div>
          )}

          {activeTab === 'settings' && (
            <div className="max-w-4xl mx-auto">
               <div className="flex justify-between items-center mb-6">
                  <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2"><Settings size={24} className="text-blue-500" /> Sistem ve Resmi Başlık Ayarları</h2>
               </div>
               <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-8">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                     <div>
                       <h3 className="text-base font-semibold text-slate-800 mb-4 border-b border-slate-100 pb-2">Kurum Bilgileri</h3>
                       <div className="space-y-4">
                         <div>
                           <label className="block text-xs font-semibold text-slate-500 mb-1.5">Okul/Kurum Adı (Resmi)</label>
                           <input type="text" value={settings.schoolTitle} onChange={e => setSettings({...settings, schoolTitle: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                         </div>
                         <div>
                           <label className="block text-xs font-semibold text-slate-500 mb-1.5">Kurum Müdürü Adı Soyadı</label>
                           <input type="text" value={settings.principalName} onChange={e => setSettings({...settings, principalName: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                         </div>
                         <div>
                           <label className="block text-xs font-semibold text-slate-500 mb-1.5">Resmi Form Üst Başlığı (Müdürlük)</label>
                           <input type="text" value={settings.provinceTitle} onChange={e => setSettings({...settings, provinceTitle: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                         </div>
                       </div>
                     </div>
                     <div>
                       <h3 className="text-base font-semibold text-slate-800 mb-4 border-b border-slate-100 pb-2">İdari Kadro</h3>
                       <div className="space-y-4">
                         <div>
                           <label className="block text-xs font-semibold text-slate-500 mb-1.5 flex justify-between items-center">
                             Müdür Başyardımcıları (Maks 2)
                           </label>
                           {settings.headVicePrincipals.map((vp, idx) => (
                             <div key={idx} className="flex gap-2 mb-2">
                               <input type="text" value={vp} onChange={e => {
                                 const newVps = [...settings.headVicePrincipals];
                                 newVps[idx] = e.target.value;
                                 setSettings({...settings, headVicePrincipals: newVps});
                               }} className="flex-1 px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder={`${idx + 1}. Müdür Başyardımcısı`} />
                             </div>
                           ))}
                         </div>
                         <div>
                           <label className="block text-xs font-semibold text-slate-500 mb-1.5 flex justify-between items-center">
                             Müdür Yardımcıları
                             <button onClick={() => setSettings({...settings, vicePrincipals: [...settings.vicePrincipals, ""]})} className="text-blue-600 hover:text-blue-700 flex items-center gap-1"><Plus size={14} /> Yeni Ekle</button>
                           </label>
                           {settings.vicePrincipals.map((vp, idx) => (
                             <div key={idx} className="flex gap-2 mb-2">
                               <input type="text" value={vp} onChange={e => {
                                 const newVps = [...settings.vicePrincipals];
                                 newVps[idx] = e.target.value;
                                 setSettings({...settings, vicePrincipals: newVps});
                               }} className="flex-1 px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder={`${idx + 1}. Müdür Yardımcısı`} />
                               <button onClick={() => {
                                 const newVps = settings.vicePrincipals.filter((_, i) => i !== idx);
                                 setSettings({...settings, vicePrincipals: newVps});
                               }} className="p-2.5 text-slate-400 hover:text-red-500 transition-colors"><Trash2 size={16} /></button>
                             </div>
                           ))}
                         </div>
                       </div>
                     </div>
                  </div>
                  
                  <div className="mt-8 pt-8 border-t border-slate-100">
                    <h3 className="text-base font-semibold text-slate-800 mb-4 border-b border-slate-100 pb-2 flex items-center gap-2">
                      <Palette size={18} className="text-indigo-500" /> Görünüm ve Kişiselleştirme
                    </h3>
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                      <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-2">Ana Renk Teması</label>
                        <div className="flex gap-3">
                          {[
                            { id: 'blue', color: 'bg-blue-600' },
                            { id: 'indigo', color: 'bg-indigo-600' },
                            { id: 'emerald', color: 'bg-emerald-600' },
                            { id: 'rose', color: 'bg-rose-600' },
                            { id: 'amber', color: 'bg-amber-600' }
                          ].map(c => (
                            <button 
                              key={c.id} 
                              onClick={() => setSettings({...settings, primaryColor: c.id})}
                              className={`w-8 h-8 rounded-full ${c.color} flex items-center justify-center transition-transform hover:scale-110 ${settings.primaryColor === c.id ? 'ring-2 ring-offset-2 ring-slate-800' : ''}`}
                            >
                              {settings.primaryColor === c.id && <CheckCircle2 size={16} className="text-white" />}
                            </button>
                          ))}
                        </div>
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-2">Ana Sayfa Düzeni</label>
                        <select 
                          value={settings.dashboardLayout || 'grid'} 
                          onChange={e => setSettings({...settings, dashboardLayout: e.target.value as 'grid' | 'list'})}
                          className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                        >
                          <option value="grid">Grid (Kartlar)</option>
                          <option value="list">Liste (Kompakt)</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-2">Takvim Stili</label>
                        <select 
                          value={settings.calendarStyle || 'modern'} 
                          onChange={e => setSettings({...settings, calendarStyle: e.target.value as 'modern' | 'classic'})}
                          className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                        >
                          <option value="modern">Modern (Geniş)</option>
                          <option value="classic">Klasik (Kompakt)</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-2">Yazdırma Yönü</label>
                        <select 
                          value={settings.printOrientation || 'landscape'} 
                          onChange={e => setSettings({...settings, printOrientation: e.target.value as 'portrait' | 'landscape'})}
                          className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                        >
                          <option value="landscape">Yatay (Landscape)</option>
                          <option value="portrait">Dikey (Portrait)</option>
                        </select>
                      </div>
                    </div>
                  </div>

                  {currentUser.role === 'admin' && (
                    <div className="mt-8 pt-8 border-t border-slate-100">
                      <h3 className="text-base font-semibold text-slate-800 mb-4 border-b border-slate-100 pb-2">
                        Kullanıcılar ve Şifreler
                      </h3>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                        <div className="bg-slate-50 border border-slate-200 rounded-2xl p-6">
                          <div className="font-semibold text-slate-800 mb-3">Admin Şifre Değiştir</div>
                          <div className="space-y-3">
                            <input id="admin-new-pass" type="password" placeholder="Yeni admin şifresi" className="w-full px-4 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                            <button onClick={() => {
                              const el = document.getElementById('admin-new-pass') as HTMLInputElement | null;
                              const pass = (el?.value || '').trim();
                              if (!pass) return addToast('Hata', 'Yeni şifre boş olamaz.', 'error');
                              const users = (settings.users && Array.isArray(settings.users) ? [...settings.users] : [...(DEFAULT_SETTINGS.users || [])]) as AppUser[];
                              const idx = users.findIndex(u => u.username === 'admin');
                              if (idx >= 0) users[idx] = { ...users[idx], password: pass, role: 'admin' };
                              else users.unshift({ username: 'admin', password: pass, role: 'admin' });
                              const updated = { ...settings, users, password: pass, username: 'admin' };
                              setSettings(updated);
                              localStorage.setItem('pts_settings', JSON.stringify(updated));
                              if (el) el.value = '';
                              addToast('Başarılı', 'Admin şifresi güncellendi.', 'success');
                            }} className="w-full px-4 py-2.5 bg-slate-900 text-white hover:bg-slate-800 rounded-xl text-sm font-semibold transition-colors">
                              Şifreyi Güncelle
                            </button>
                          </div>
                        </div>

                        <div className="bg-slate-50 border border-slate-200 rounded-2xl p-6">
                          <div className="font-semibold text-slate-800 mb-3">Öğretmen Hesabı Oluştur / Güncelle</div>
                          <div className="space-y-3">
                            <select id="teacher-personnel" className="w-full px-4 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                              <option value="">Personel seçin</option>
                              {personnel.filter(p => p.role === 'Öğretmen').map(p => (
                                <option key={p.id} value={p.id}>{p.name} ({p.branch})</option>
                              ))}
                            </select>
                            <input id="teacher-username" placeholder="Kullanıcı adı (örn: tckn veya sicil)" className="w-full px-4 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                            <input id="teacher-password" type="password" placeholder="Şifre" className="w-full px-4 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                            <button onClick={() => {
                              const pId = Number((document.getElementById('teacher-personnel') as HTMLSelectElement | null)?.value || 0);
                              const username = ((document.getElementById('teacher-username') as HTMLInputElement | null)?.value || '').trim();
                              const password = ((document.getElementById('teacher-password') as HTMLInputElement | null)?.value || '').trim();
                              if (!pId || !username || !password) return addToast('Hata', 'Personel, kullanıcı adı ve şifre zorunludur.', 'error');

                              const users = (settings.users && Array.isArray(settings.users) ? [...settings.users] : [...(DEFAULT_SETTINGS.users || [])]) as AppUser[];
                              // aynı username varsa güncelle
                              const idx = users.findIndex(u => u.username === username);
                              const record: AppUser = { username, password, role: 'user', personnelId: pId };
                              if (idx >= 0) users[idx] = { ...users[idx], ...record };
                              else users.push(record);
                              const updated = { ...settings, users };
                              setSettings(updated);
                              localStorage.setItem('pts_settings', JSON.stringify(updated));
                              (document.getElementById('teacher-username') as HTMLInputElement | null)?.blur();
                              (document.getElementById('teacher-password') as HTMLInputElement | null)?.blur();
                              (document.getElementById('teacher-username') as HTMLInputElement | null)!.value = '';
                              (document.getElementById('teacher-password') as HTMLInputElement | null)!.value = '';
                              addToast('Başarılı', 'Öğretmen hesabı kaydedildi.', 'success');
                            }} className="w-full px-4 py-2.5 bg-blue-600 text-white hover:bg-blue-700 rounded-xl text-sm font-semibold transition-colors shadow-sm shadow-blue-600/20">
                              Hesabı Kaydet
                            </button>
                          </div>
                          <div className="mt-4 text-[11px] text-slate-500">
                            Not: Öğretmen rolü sadece kendi personel bilgilerini, duyuruları ve ders/nöbet çizelgelerini görebilir.
                          </div>
                        </div>
                      </div>

                      <div className="mt-6">
                        <div className="text-sm font-semibold text-slate-800 mb-3">Mevcut Öğretmen Hesapları</div>
                        <div className="space-y-2">
                          {(settings.users || []).filter(u => u.role === 'user').length === 0 ? (
                            <div className="text-xs text-slate-500">Henüz öğretmen hesabı yok.</div>
                          ) : (
                            (settings.users || []).filter(u => u.role === 'user').map(u => {
                              const p = personnel.find(x => x.id === u.personnelId);
                              return (
                                <div key={u.username} className="flex items-center justify-between bg-white border border-slate-200 rounded-xl px-4 py-2.5">
                                  <div className="text-xs">
                                    <div className="font-semibold text-slate-800">{u.username}</div>
                                    <div className="text-slate-500">{p ? `${p.name} (${p.branch})` : `PersonelId: ${u.personnelId || '-'}`}</div>
                                  </div>
                                  <button onClick={() => {
                                    const users = (settings.users || []).filter(x => x.username !== u.username);
                                    const updated = { ...settings, users };
                                    setSettings(updated);
                                    localStorage.setItem('pts_settings', JSON.stringify(updated));
                                    addToast('Silindi', 'Öğretmen hesabı silindi.', 'info');
                                  }} className="px-3 py-1.5 bg-white border border-red-200 text-red-600 hover:bg-red-50 rounded-lg text-xs font-semibold transition-colors">
                                    Sil
                                  </button>
                                </div>
                              );
                            })
                          )}
                        </div>
                      </div>
                    </div>
                  )}

                  <div className="mt-8 pt-6 border-t border-slate-100 flex justify-end">
                     <button onClick={() => {
                       localStorage.setItem('pts_settings', JSON.stringify(settings));
                       addToast("Başarılı", "Ayarlar kaydedildi.", "success");
                     }} className="px-6 py-2.5 bg-blue-600 text-white hover:bg-blue-700 rounded-xl text-sm font-semibold transition-colors shadow-sm shadow-blue-600/20">Ayarları Kaydet</button>
                  </div>
               </div>

               {/* Eğitim-Öğretim Yılı Arşivi */}
               <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-8 mt-6">
                  <h3 className="text-lg font-bold text-slate-800 mb-2 flex items-center gap-2">
                    🗂️ Eğitim-Öğretim Yılı Arşivi
                  </h3>
                  <p className="text-sm text-slate-500 mb-6">
                    Her eğitim-öğretim yılı sonunda (Haziran-Ağustos) tüm veriler (personel, izinler, görev değişiklikleri, sendika, terfiler, evraklar, duyurular, ders/nöbet çizelgeleri vb.) otomatik olarak arşivlenir. Aşağıdan manuel arşiv alabilir, eski arşivleri görüntüleyebilir, JSON olarak indirebilir veya geri yükleyebilirsiniz.
                  </p>

                  <div className="flex items-center justify-between p-4 bg-slate-50 rounded-xl mb-4">
                    <div>
                      <div className="text-xs text-slate-500">Mevcut Eğitim-Öğretim Yılı</div>
                      <div className="text-lg font-bold text-blue-600">{getEduYear()}</div>
                    </div>
                    <div className="text-right">
                      <div className="text-xs text-slate-500">Toplam Arşiv</div>
                      <div className="text-lg font-bold text-slate-700">{archives.length}</div>
                    </div>
                  </div>

                  {currentUser.role === 'admin' && (
                    <div className="flex flex-wrap gap-3 mb-4">
                      <button
                        onClick={() => {
                          const eduYear = getEduYear();
                          if (window.confirm(`${eduYear} eğitim-öğretim yılı için tüm verilerin (A'dan Z'ye) bir kopyası arşivlenecek. Mevcut veriler silinmeyecek.\n\nDevam edilsin mi?`)) {
                            archiveCurrentEduYear(false, 'Manuel arşiv');
                            addToast('Başarılı', `${eduYear} eğitim-öğretim yılı arşivlendi.`, 'success');
                          }
                        }}
                        className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold transition-colors shadow-sm"
                      >
                        📦 Şimdi Arşivle ({getEduYear()})
                      </button>

                      <button
                        onClick={() => {
                          const eduYear = getEduYear();
                          if (!window.confirm(`DİKKAT: Bu işlem ${eduYear} eğitim-öğretim yılını arşivledikten sonra TÜM dönemsel verileri (izinler, görev değişiklikleri, sendika, terfiler, evraklar, duyurular, bildirimler) SIFIRLAYACAK.\n\nKalıcı veriler korunacak: Personel listesi, kullanıcı hesapları, ayarlar, ders/nöbet çizelgeleri.\n\nDevam etmek istediğinize emin misiniz?`)) return;
                          if (!window.confirm('Son kez sormak için: Yeni eğitim-öğretim yılına temiz başlanacak. Onaylıyor musunuz?')) return;
                          archiveCurrentEduYear(false, 'Yeni yıla geçiş arşivi');
                          clearPeriodicData();
                          addToast('Başarılı', `${eduYear} arşivlendi ve dönemsel veriler sıfırlandı. Yeni eğitim-öğretim yılına temiz başlanabilir.`, 'success');
                        }}
                        className="flex items-center gap-2 px-4 py-2.5 bg-amber-600 hover:bg-amber-700 text-white rounded-xl text-sm font-semibold transition-colors shadow-sm"
                      >
                        🔄 Yeni Yıla Temiz Başla (Arşivle + Sıfırla)
                      </button>
                    </div>
                  )}

                  <div className="border-t border-slate-100 pt-4">
                    <div className="text-sm font-semibold text-slate-700 mb-3">Arşivlenmiş Yıllar</div>
                    {archives.length === 0 ? (
                      <div className="text-xs text-slate-400 italic py-6 text-center bg-slate-50 rounded-xl">
                        Henüz arşivlenmiş eğitim-öğretim yılı bulunmuyor.
                      </div>
                    ) : (
                      <div className="space-y-2 max-h-96 overflow-y-auto">
                        {archives.map(arc => {
                          const dt = new Date(arc.archivedAt);
                          const dataKeys = Object.keys(arc.data);
                          const sizeKB = (JSON.stringify(arc.data).length / 1024).toFixed(1);
                          return (
                            <div key={arc.id} className="flex items-center justify-between bg-slate-50 border border-slate-200 rounded-xl px-4 py-3">
                              <div className="flex-1 min-w-0">
                                <div className="flex items-center gap-2 flex-wrap">
                                  <span className="font-bold text-slate-800">{arc.eduYear}</span>
                                  {arc.autoArchived ? (
                                    <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-blue-100 text-blue-700">Otomatik</span>
                                  ) : (
                                    <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700">Manuel</span>
                                  )}
                                  {arc.note && <span className="text-[10px] text-slate-500">— {arc.note}</span>}
                                </div>
                                <div className="text-[11px] text-slate-500 mt-0.5">
                                  {dt.toLocaleString('tr-TR')} • {dataKeys.length} veri kümesi • {sizeKB} KB
                                </div>
                              </div>
                              <div className="flex items-center gap-2 flex-shrink-0">
                                <button
                                  onClick={() => {
                                    const blob = new Blob([JSON.stringify(arc, null, 2)], { type: 'application/json' });
                                    const url = URL.createObjectURL(blob);
                                    const a = document.createElement('a');
                                    a.href = url;
                                    a.download = `arsiv_${arc.eduYear}_${dt.toISOString().slice(0, 10)}.json`;
                                    document.body.appendChild(a);
                                    a.click();
                                    document.body.removeChild(a);
                                    URL.revokeObjectURL(url);
                                  }}
                                  className="px-3 py-1.5 bg-white border border-slate-200 text-slate-700 hover:bg-slate-50 rounded-lg text-xs font-semibold transition-colors"
                                  title="JSON olarak indir"
                                >
                                  ⬇ İndir
                                </button>
                                {currentUser.role === 'admin' && (
                                  <>
                                    <button
                                      onClick={() => {
                                        if (!window.confirm(`${arc.eduYear} eğitim-öğretim yılı arşivi GERİ YÜKLENECEK.\n\nUYARI: Mevcut tüm veriler (personel, izinler, ayarlar, vb.) arşivdeki verilerle DEĞİŞTİRİLECEK. Kayıp önlemek için önce mevcut yılı ayrı bir arşiv olarak almanız önerilir.\n\nDevam edilsin mi?`)) return;
                                        Object.entries(arc.data).forEach(([k, v]) => {
                                          localStorage.setItem(k, v);
                                        });
                                        addToast('Geri Yüklendi', `${arc.eduYear} arşivi geri yüklendi. Sayfa yenileniyor...`, 'success');
                                        setTimeout(() => window.location.reload(), 1500);
                                      }}
                                      className="px-3 py-1.5 bg-amber-50 border border-amber-200 text-amber-700 hover:bg-amber-100 rounded-lg text-xs font-semibold transition-colors"
                                      title="Bu arşivi geri yükle"
                                    >
                                      ↻ Geri Yükle
                                    </button>
                                    <button
                                      onClick={() => {
                                        if (!window.confirm(`${arc.eduYear} (${dt.toLocaleDateString('tr-TR')}) arşivini silmek istediğinize emin misiniz? Bu işlem geri alınamaz.`)) return;
                                        const updated = archives.filter(a => a.id !== arc.id);
                                        setArchives(updated);
                                        localStorage.setItem('pts_archives', JSON.stringify(updated));
                                        addToast('Silindi', 'Arşiv silindi.', 'info');
                                      }}
                                      className="p-1.5 bg-white border border-red-200 text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                                      title="Arşivi sil"
                                    >
                                      <Trash2 size={14} />
                                    </button>
                                  </>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>

                  <div className="mt-4 text-[11px] text-slate-500 space-y-1">
                    <p>• Eğitim-öğretim yılı: Eylül başında başlar, ertesi yıl Ağustos sonunda biter (örn: 2025-2026).</p>
                    <p>• Otomatik arşiv: 15 Haziran sonrası uygulama açıldığında, ilgili yıl için arşiv yoksa kendiliğinden alınır.</p>
                    <p>• "İndir" butonu ile arşivinizi yedek dosya olarak bilgisayarınıza kaydedebilirsiniz.</p>
                  </div>
               </div>

               {/* Yedekleme & Geri Yükleme Bölümü */}
               {currentUser.role === 'admin' && (
               <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-8 mt-6">
                  <h3 className="text-lg font-bold text-slate-800 mb-2 flex items-center gap-2">
                    💾 Veri Yedekleme ve Geri Yükleme
                  </h3>
                  <p className="text-sm text-slate-500 mb-6">
                    Veriler otomatik olarak diske yedeklenir. Güncelleme, yeniden kurulum veya
                    beklenmedik veri kaybında otomatik olarak geri yüklenir.
                  </p>

                  <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 mb-6">
                    <div className="flex items-start gap-3">
                      <CheckCircle2 size={20} className="text-emerald-600 flex-shrink-0 mt-0.5" />
                      <div className="text-sm text-emerald-800">
                        <div className="font-semibold mb-1">Otomatik Yedekleme Aktif</div>
                        <ul className="text-xs space-y-1 text-emerald-700">
                          <li>• Her veri değişikliğinden sonra otomatik yedek alınır.</li>
                          <li>• Güncelleme sonrası veriler korunur; localStorage silinmiş olsa bile yedekten otomatik geri yüklenir.</li>
                          <li>• Son 7 günün günlük yedekleri ayrıca saklanır.</li>
                          <li>• Yedekler güvenli kullanıcı klasöründe tutulur (AppData/Roaming/Okul Takip Sistemi/backups).</li>
                        </ul>
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
                    <button
                      onClick={async () => {
                        if (!window.backup) {
                          addToast('Bilgi', 'Bu özellik yalnızca masaüstü uygulamada kullanılabilir.', 'info');
                          return;
                        }
                        try {
                          const payload = collectBackupPayload();
                          const res = await window.backup.saveAs(payload);
                          if (res.success) {
                            addToast('Yedek Alındı', `Yedek dosyası kaydedildi: ${res.path}`, 'success');
                          } else if (!res.canceled) {
                            addToast('Hata', res.error || 'Yedek kaydedilemedi.', 'error');
                          }
                        } catch (e: any) {
                          addToast('Hata', e?.message || 'Yedek alma sırasında hata oluştu.', 'error');
                        }
                      }}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold transition-colors shadow-sm shadow-blue-600/20"
                    >
                      <Download size={16} />
                      Yedek Dosyası Olarak Kaydet
                    </button>

                    <button
                      onClick={async () => {
                        if (!window.backup) {
                          addToast('Bilgi', 'Bu özellik yalnızca masaüstü uygulamada kullanılabilir.', 'info');
                          return;
                        }
                        try {
                          const payload = collectBackupPayload();
                          const res = await window.backup.exportToDocuments(payload);
                          if (res.success) {
                            addToast('Yedek Alındı', `Belgeler klasörüne kaydedildi.`, 'success');
                          } else {
                            addToast('Hata', res.error || 'Yedek alınamadı.', 'error');
                          }
                        } catch (e: any) {
                          addToast('Hata', e?.message || 'Yedek alma sırasında hata oluştu.', 'error');
                        }
                      }}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-xl text-sm font-semibold transition-colors"
                    >
                      📁 Belgeler Klasörüne Hızlı Yedek
                    </button>

                    <button
                      onClick={async () => {
                        if (!window.backup) {
                          addToast('Bilgi', 'Bu özellik yalnızca masaüstü uygulamada kullanılabilir.', 'info');
                          return;
                        }
                        if (!window.confirm('Yedek dosyası seçeceksiniz.\n\nUYARI: Mevcut tüm veriler (personel, izinler, ayarlar, vb.) seçilen dosyadaki verilerle DEĞİŞTİRİLECEK. Devam etmeden önce mevcut verilerin yedeğini almanız önerilir.\n\nDevam edilsin mi?')) return;
                        try {
                          const res = await window.backup.import();
                          if (res.canceled) return;
                          if (!res.success || !res.data) {
                            addToast('Hata', res.error || 'Yedek dosyası okunamadı.', 'error');
                            return;
                          }
                          const apply = applyBackupPayload(res.data);
                          if (!apply.ok) {
                            addToast('Hata', 'Yedek dosyası geçerli bir biçimde değil veya boş.', 'error');
                            return;
                          }
                          addToast('Geri Yüklendi', `${apply.count} veri kümesi geri yüklendi. Sayfa yenileniyor...`, 'success');
                          setTimeout(() => window.location.reload(), 1200);
                        } catch (e: any) {
                          addToast('Hata', e?.message || 'Geri yükleme sırasında hata oluştu.', 'error');
                        }
                      }}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-amber-600 hover:bg-amber-700 text-white rounded-xl text-sm font-semibold transition-colors shadow-sm shadow-amber-600/20"
                    >
                      ↻ Yedek Dosyasından Geri Yükle
                    </button>

                    <button
                      onClick={async () => {
                        if (!window.backup) {
                          addToast('Bilgi', 'Bu özellik yalnızca masaüstü uygulamada kullanılabilir.', 'info');
                          return;
                        }
                        try {
                          await window.backup.openFolder();
                        } catch (e: any) {
                          addToast('Hata', e?.message || 'Klasör açılamadı.', 'error');
                        }
                      }}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-xl text-sm font-semibold transition-colors"
                    >
                      📂 Yedek Klasörünü Aç
                    </button>
                  </div>

                  <div className="mt-4 p-4 bg-amber-50 border border-amber-200 rounded-xl">
                    <div className="flex items-start gap-2">
                      <AlertCircle size={16} className="text-amber-600 flex-shrink-0 mt-0.5" />
                      <div className="text-xs text-amber-800 space-y-1">
                        <p className="font-semibold">Önemli Bilgiler</p>
                        <p>• Güncelleme sonrası verileriniz kaybolmaz. Yedekler kullanıcı verisi klasöründe saklanır ve güncelleme bunlara dokunmaz.</p>
                        <p>• Yine de önemli zamanlarda manuel olarak <strong>"Belgeler Klasörüne Hızlı Yedek"</strong> almanızı öneririz. Bu dosyayı USB veya bulut depolamaya kopyalayabilirsiniz.</p>
                        <p>• Bilgisayar değiştiriyorsanız: Önce eski bilgisayarda yedek dosyası alın, yeni bilgisayarda "Yedek Dosyasından Geri Yükle" butonunu kullanın.</p>
                      </div>
                    </div>
                  </div>

                  <div className="mt-3 p-4 bg-sky-50 border border-sky-200 rounded-xl">
                    <div className="flex items-start gap-2">
                      <span className="text-base leading-none mt-0.5">☁️</span>
                      <div className="text-xs text-sky-900 space-y-1">
                        <p className="font-semibold">İpucu: Otomatik Bulut Yedeği</p>
                        <p>
                          Belgeler klasörünüz <strong>OneDrive</strong>, <strong>Google Drive</strong> veya
                          <strong> Yandex Disk</strong> ile senkronize ediliyorsa, <strong>"Belgeler Klasörüne Hızlı Yedek"</strong>
                          butonuyla aldığınız yedekler otomatik olarak bulutunuza da kopyalanır.
                          Böylece bilgisayar arızası, hırsızlık veya format durumunda verileriniz güvende olur.
                        </p>
                        <p className="text-[11px] text-sky-700 mt-1">
                          Kontrol: OneDrive simgesine sağ tıklayın → Ayarlar → "Yedekle" sekmesinde
                          <strong> "Belgeler"</strong> seçeneğinin işaretli olduğundan emin olun.
                          Yedek dosyalarının üzerindeki bulut/onay simgesi senkronizasyonun çalıştığını gösterir.
                        </p>
                      </div>
                    </div>
                  </div>
               </div>
               )}

               {/* Güncelleme Bölümü */}
               <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-8 mt-6">
                  <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                    <RefreshCw size={20} className="text-emerald-500" /> Uygulama Güncellemeleri
                  </h3>
                  
                  <div className="flex items-center justify-between p-4 bg-slate-50 rounded-xl mb-4">
                    <div>
                      <div className="text-sm font-semibold text-slate-700">Mevcut Sürüm</div>
                      <div className="text-2xl font-bold text-blue-600">v{appVersion}</div>
                    </div>
                    <div className="text-right">
                      <div className="text-xs text-slate-500">Otomatik güncelleme</div>
                      <div className="text-sm font-medium text-emerald-600">✓ Aktif</div>
                    </div>
                  </div>

                  {updateStatus && (
                    <div className={`p-4 rounded-xl mb-4 ${
                      updateStatus.status === 'available' || updateStatus.status === 'downloading' ? 'bg-amber-50 border border-amber-200' :
                      updateStatus.status === 'downloaded' ? 'bg-emerald-50 border border-emerald-200' :
                      updateStatus.status === 'error' ? 'bg-rose-50 border border-rose-200' :
                      updateStatus.status === 'not-available' ? 'bg-slate-50 border border-slate-200' :
                      'bg-blue-50 border border-blue-200'
                    }`}>
                      <div className="flex items-center gap-3">
                        {updateStatus.status === 'checking' && (
                          <RefreshCw size={18} className="text-blue-500 animate-spin" />
                        )}
                        {updateStatus.status === 'downloading' && (
                          <div className="w-5 h-5 border-2 border-amber-500 border-t-transparent rounded-full animate-spin" />
                        )}
                        {updateStatus.status === 'downloaded' && (
                          <CheckCircle2 size={18} className="text-emerald-500" />
                        )}
                        {updateStatus.status === 'error' && (
                          <AlertCircle size={18} className="text-rose-500" />
                        )}
                        {updateStatus.status === 'not-available' && (
                          <CheckCircle2 size={18} className="text-slate-500" />
                        )}
                        {updateStatus.status === 'available' && (
                          <AlertCircle size={18} className="text-amber-500" />
                        )}
                        <div className="flex-1">
                          <div className="text-sm font-medium text-slate-700">{updateStatus.message}</div>
                          {updateStatus.status === 'downloading' && updateStatus.percent !== undefined && (
                            <div className="mt-2">
                              <div className="w-full h-2 bg-slate-200 rounded-full overflow-hidden">
                                <div 
                                  className="h-full bg-amber-500 transition-all duration-300" 
                                  style={{ width: `${updateStatus.percent}%` }}
                                />
                              </div>
                              <div className="text-xs text-slate-500 mt-1">{updateStatus.percent}% tamamlandı</div>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  )}

                  <div className="flex gap-3">
                    <button
                      onClick={() => {
                        if (window.updater) {
                          setIsCheckingUpdate(true);
                          window.updater.checkForUpdates();
                        } else {
                          addToast('Bilgi', 'Güncelleme kontrolü sadece paketlenmiş uygulamada çalışır.', 'info');
                        }
                      }}
                      disabled={isCheckingUpdate}
                      className="flex items-center gap-2 px-4 py-2.5 bg-slate-100 hover:bg-slate-200 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl text-sm font-semibold text-slate-700 transition-colors"
                    >
                      <RefreshCw size={16} className={isCheckingUpdate ? 'animate-spin' : ''} />
                      Güncelleme Kontrolü
                    </button>
                    
                    {updateStatus?.status === 'downloaded' && (
                      <button
                        onClick={() => {
                          if (window.updater) {
                            window.updater.installUpdate();
                          }
                        }}
                        className="flex items-center gap-2 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 rounded-xl text-sm font-semibold text-white transition-colors shadow-sm"
                      >
                        <Download size={16} />
                        Şimdi Yükle ve Yeniden Başlat
                      </button>
                    )}
                  </div>

                  <div className="mt-4 text-xs text-slate-500">
                    <p>• Güncelleme arka planda otomatik olarak indirilir</p>
                    <p>• İndirilen güncelleme, uygulama yeniden başlatıldığında yüklenir</p>
                    <p>• Her 30 dakikada bir otomatik güncelleme kontrolü yapılır</p>
                  </div>
               </div>
            </div>
          )}

        </div>
      </div>

      {/* --- MODALS & NOTIFICATIONS --- */}
      
      {/* Bildirim Paneli */}
      {showNotificationPanel && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-end z-[10001]" onClick={() => setShowNotificationPanel(false)}>
          <div 
            className="w-[400px] bg-white h-full shadow-2xl animate-in slide-in-from-right duration-300"
            onClick={e => e.stopPropagation()}
          >
            <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-gradient-to-r from-blue-600 to-blue-700">
              <div className="flex items-center gap-3">
                <Bell size={22} className="text-white" />
                <h3 className="text-lg font-bold text-white">Bildirimler</h3>
                {notifications.filter(n => !n.read).length > 0 && (
                  <span className="px-2 py-0.5 bg-red-500 text-white text-xs font-bold rounded-full">
                    {notifications.filter(n => !n.read).length} yeni
                  </span>
                )}
              </div>
              <button 
                onClick={() => setShowNotificationPanel(false)} 
                className="text-white/70 hover:text-white transition-colors p-1 hover:bg-white/10 rounded-lg"
              >
                <X size={20} />
              </button>
            </div>
            
            <div className="p-4 border-b border-slate-100 flex gap-2">
              <button 
                onClick={() => {
                  const updated = notifications.map(n => ({...n, read: true}));
                  setNotifications(updated);
                  localStorage.setItem('pts_notifs', JSON.stringify(updated));
                }}
                className="flex-1 px-3 py-2 text-xs font-semibold bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-700 transition-colors"
              >
                Tümünü Okundu İşaretle
              </button>
              <button 
                onClick={() => {
                  setNotifications([]);
                  localStorage.setItem('pts_notifs', JSON.stringify([]));
                }}
                className="px-3 py-2 text-xs font-semibold bg-rose-50 hover:bg-rose-100 rounded-lg text-rose-600 transition-colors"
              >
                Temizle
              </button>
            </div>
            
            <div className="overflow-y-auto h-[calc(100vh-180px)]">
              {notifications.length === 0 ? (
                <div className="flex flex-col items-center justify-center h-64 text-slate-400">
                  <Bell size={48} className="mb-4 opacity-30" />
                  <p className="text-sm">Henüz bildirim yok</p>
                </div>
              ) : (
                <div className="divide-y divide-slate-100">
                  {notifications.map(n => (
                    <div 
                      key={n.id} 
                      className={`p-4 hover:bg-slate-50 transition-colors cursor-pointer ${!n.read ? 'bg-blue-50/50' : ''}`}
                      onClick={() => {
                        const updated = notifications.map(notif => 
                          notif.id === n.id ? {...notif, read: true} : notif
                        );
                        setNotifications(updated);
                        localStorage.setItem('pts_notifs', JSON.stringify(updated));
                      }}
                    >
                      <div className="flex items-start gap-3">
                        <div className={`w-10 h-10 rounded-xl flex items-center justify-center flex-shrink-0 ${
                          n.type === 'success' ? 'bg-emerald-100 text-emerald-600' :
                          n.type === 'error' ? 'bg-rose-100 text-rose-600' :
                          n.type === 'warning' ? 'bg-amber-100 text-amber-600' :
                          'bg-blue-100 text-blue-600'
                        }`}>
                          {n.type === 'success' ? <CheckCircle2 size={20} /> :
                           n.type === 'error' ? <X size={20} /> :
                           n.type === 'warning' ? <AlertCircle size={20} /> :
                           <Bell size={20} />}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2">
                            <h4 className="text-sm font-semibold text-slate-800 truncate">{n.title}</h4>
                            {!n.read && (
                              <span className="w-2 h-2 bg-blue-500 rounded-full flex-shrink-0"></span>
                            )}
                          </div>
                          <p className="text-xs text-slate-600 mt-1 line-clamp-2">{n.message}</p>
                          <p className="text-[10px] text-slate-400 mt-2">{n.timestamp}</p>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      <div className="fixed bottom-6 right-6 z-[10000] flex flex-col gap-3">
        {toasts.map(t => (
          <div key={t.id} className={`flex items-center gap-4 p-4 rounded-xl shadow-lg min-w-[300px] text-white animate-in slide-in-from-right-8 ${
            t.type === 'success' ? 'bg-emerald-600' : t.type === 'error' ? 'bg-rose-600' : 'bg-blue-600'
          }`}>
            <div className="text-xl">{t.type === 'success' ? '✓' : t.type === 'error' ? '✕' : 'ℹ'}</div>
            <div className="flex-1">
              <div className="font-bold text-sm">{t.title}</div>
              <div className="text-xs opacity-90 mt-0.5">{t.message}</div>
            </div>
            <button onClick={() => removeToast(t.id)} className="text-white/70 hover:text-white transition-colors">✕</button>
          </div>
        ))}
      </div>

      {showAnnounceModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-center items-center z-[10000]">
          <div className="bg-white p-8 rounded-2xl w-[500px] shadow-2xl">
            <h3 className="text-xl font-bold text-slate-800 mb-6">Yeni Duyuru Yayınla</h3>
            <div className="flex gap-2 mb-6 bg-slate-100 p-1.5 rounded-xl">
              {['text', 'link', 'file'].map(type => (
                <button key={type} onClick={() => setNewAnnounce({...newAnnounce, type: type as any})} className={`flex-1 py-2 rounded-lg text-sm font-medium transition-all ${newAnnounce.type === type ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}>
                  {type === 'text' ? '📝 Metin' : type === 'link' ? '🔗 Link' : '📁 Dosya'}
                </button>
              ))}
            </div>
            <input placeholder="Duyuru Başlığı" value={newAnnounce.title || ''} onChange={e => setNewAnnounce({...newAnnounce, title: e.target.value})} className="w-full px-4 py-3 mb-4 rounded-xl border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
            {newAnnounce.type === 'text' && (
              <textarea placeholder="Duyuru İçeriği" value={newAnnounce.content || ''} onChange={e => setNewAnnounce({...newAnnounce, content: e.target.value})} className="w-full px-4 py-3 mb-6 rounded-xl border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 h-32 resize-none" />
            )}
            {newAnnounce.type === 'link' && (
              <input placeholder="URL (https://...)" value={newAnnounce.url || ''} onChange={e => setNewAnnounce({...newAnnounce, url: e.target.value})} className="w-full px-4 py-3 mb-6 rounded-xl border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
            )}
            {newAnnounce.type === 'file' && (
              <div className="mb-6">
                <input type="file" onChange={handleFileUpload} className="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer" />
                {newAnnounce.fileName && <p className="text-xs text-emerald-600 mt-2 font-medium">✓ Seçili: {newAnnounce.fileName}</p>}
              </div>
            )}
            <div className="flex gap-3">
               <button onClick={() => setShowAnnounceModal(false)} className="flex-1 py-3 bg-slate-100 text-slate-700 hover:bg-slate-200 rounded-xl font-semibold transition-colors">Vazgeç</button>
               <button onClick={handleAddAnnouncement} className="flex-1 py-3 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-semibold transition-colors shadow-sm shadow-blue-600/20">Hemen Yayınla</button>
            </div>
          </div>
        </div>
      )}

      {showAddModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-center items-center z-[10000]">
          <div className="bg-white p-8 rounded-2xl w-[850px] max-h-[95vh] overflow-y-auto shadow-2xl">
            <h3 className="text-xl font-bold text-slate-800 mb-6 border-b border-slate-100 pb-4">{editingPersonId ? '✏️ Personel Bilgilerini Güncelle' : '🆕 Detaylı Personel Kaydı'}</h3>
            
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              
              {/* Bölüm 1: Kimlik Bilgileri */}
              <div className="col-span-1 md:col-span-3 border-l-4 border-blue-500 pl-3 font-bold text-blue-600 mt-2">🛡️ Kimlik ve Kurum Bilgileri</div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Ad Soyad *</label>
                <input placeholder="Ad Soyad" value={newPerson.name || ''} onChange={e => setNewPerson({...newPerson, name: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">TC Kimlik No *</label>
                <input placeholder="11 Haneli TC" maxLength={11} value={newPerson.tc || ''} onChange={e => setNewPerson({...newPerson, tc: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Personel No / Sicil</label>
                <input placeholder="Örn: 12345678" value={newPerson.personnelNo || ''} onChange={e => setNewPerson({...newPerson, personnelNo: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>

              {/* Bölüm 2: Görev ve Branş */}
              <div className="col-span-1 md:col-span-3 border-l-4 border-amber-500 pl-3 font-bold text-amber-600 mt-4">💼 Görev ve Branş Detayları</div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Ana Görevi</label>
                <select value={newPerson.role} onChange={e => setNewPerson({...newPerson, role: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                  {ROLES.map(r => <option key={r} value={r}>{r}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Görev Ünvanı</label>
                <select value={newPerson.title} onChange={e => setNewPerson({...newPerson, title: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                  {TITLES.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              {!NON_ACADEMIC_ROLES.includes(newPerson.role || '') && (
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Branş / Alan (Seçin veya Yazın)</label>
                <input 
                  list="branch-list" 
                  placeholder="Örn: Matematik" 
                  value={newPerson.branch || ''} 
                  onChange={e => setNewPerson({...newPerson, branch: e.target.value})} 
                  className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" 
                />
                <datalist id="branch-list">
                  {BRANCHES.map(b => <option key={b} value={b} />)}
                </datalist>
              </div>
              )}
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">İstihdam Türü</label>
                <select value={newPerson.employmentType} onChange={e => setNewPerson({...newPerson, employmentType: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                  {EMPLOYMENT_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              {!NON_ACADEMIC_ROLES.includes(newPerson.role || '') && (
              <>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Derece (Otomatik/Manuel)</label>
                <input type="number" min="1" max="15" value={newPerson.degree} onChange={e => setNewPerson({...newPerson, degree: Number(e.target.value)})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Kademe (Otomatik/Manuel) {newPerson.degree === 1 ? '(1/4 son)' : ''}</label>
                <input type="number" min="1" max={newPerson.degree === 1 ? 4 : 3} value={newPerson.level} onChange={e => setNewPerson({...newPerson, level: Number(e.target.value)})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              </>
              )}

              {/* Bölüm 3: Tarih ve Üyelik */}
              <div className="col-span-1 md:col-span-3 border-l-4 border-emerald-500 pl-3 font-bold text-emerald-600 mt-4">📅 Atama ve Sendika Bilgileri</div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Göreve Başlama Tarihi *</label>
                <input 
                  type="date" 
                  value={newPerson.startDate || ''} 
                  onChange={e => setNewPerson({...newPerson, startDate: e.target.value})} 
                  className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" 
                />
                <div className="text-[10px] text-emerald-600 mt-1 font-medium">ℹ️ Değiştirdiğinizde Derece/Kademe otomatik hesaplanır.</div>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Sendika</label>
                <input placeholder="Örn: Eğitim-Bir-Sen" value={newPerson.union || ''} onChange={e => setNewPerson({...newPerson, union: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Öğrenim Durumu</label>
                <select value={newPerson.education} onChange={e => setNewPerson({...newPerson, education: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                  {EDUCATIONS.map(ed => <option key={ed} value={ed}>{ed}</option>)}
                </select>
                <div className="text-[10px] text-blue-500 mt-1 font-medium">ℹ️ Başlangıç D/K öğrenim durumuna göre belirlenir.</div>
              </div>

              {/* Bölüm 4: İletişim ve Banka */}
              <div className="col-span-1 md:col-span-3 border-l-4 border-indigo-500 pl-3 font-bold text-indigo-600 mt-4">📞 İletişim ve Banka</div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Medeni Durum</label>
                <select value={newPerson.maritalStatus} onChange={e => setNewPerson({...newPerson, maritalStatus: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                  {MARITAL_STATUSES.map(m => <option key={m} value={m}>{m}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Çocuk Sayısı</label>
                <input type="number" min="0" value={newPerson.childrenCount} onChange={e => setNewPerson({...newPerson, childrenCount: Number(e.target.value)})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">Telefon Numarası</label>
                <input placeholder="05XX XXX XX XX" value={newPerson.phone || ''} onChange={e => setNewPerson({...newPerson, phone: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div className="col-span-1 md:col-span-2">
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">IBAN Numarası</label>
                <input placeholder="TRXX XXXX XXXX XXXX XXXX XXXX XX" value={newPerson.iban || ''} onChange={e => setNewPerson({...newPerson, iban: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">E-Posta</label>
                <input type="email" placeholder="ornek@meb.k12.tr" value={newPerson.email || ''} onChange={e => setNewPerson({...newPerson, email: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div className="col-span-1 md:col-span-3">
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">İkametgah Adresi</label>
                <textarea placeholder="Mahalle, Sokak, No, Daire, İlçe/İl" value={newPerson.address || ''} onChange={e => setNewPerson({...newPerson, address: e.target.value})} className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 h-20 resize-none" />
              </div>
            </div>

            <div className="flex gap-4 mt-8 pt-6 border-t border-slate-100">
               <button onClick={() => { setShowAddModal(false); setEditingPersonId(null); }} className="flex-1 py-3 bg-slate-100 text-slate-700 hover:bg-slate-200 rounded-xl font-semibold transition-colors">Vazgeç</button>
               <button onClick={handleSavePersonnel} className={`flex-[2] py-3 text-white rounded-xl font-semibold transition-colors shadow-sm ${editingPersonId ? 'bg-blue-600 hover:bg-blue-700 shadow-blue-600/20' : 'bg-emerald-600 hover:bg-emerald-700 shadow-emerald-600/20'}`}>
                {editingPersonId ? '💾 Değişiklikleri Güncelle' : '💾 Personel Kaydını Tamamla'}
               </button>
            </div>
          </div>
        </div>
      )}

      {showDetailModal && selectedPersonForDetail && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-center items-center z-[10000]">
          <div className="bg-white p-8 rounded-2xl w-[900px] max-h-[95vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center border-b border-slate-100 pb-4 mb-6">
              <h3 className="text-xl font-bold text-slate-800 m-0">👤 Personel Detaylı Bilgi Kartı</h3>
              <button onClick={() => setShowDetailModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors text-2xl">✕</button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
              {/* Sol Sütun: Profil Özeti */}
              <div className="bg-slate-50 p-6 rounded-2xl text-center border border-slate-100 h-fit">
                <div className="w-24 h-24 bg-blue-500 text-white rounded-full flex items-center justify-center text-4xl mx-auto mb-4 shadow-md shadow-blue-500/20 font-bold">
                  {selectedPersonForDetail.name.charAt(0)}
                </div>
                <h2 className="text-lg font-bold text-slate-800 mb-1">{selectedPersonForDetail.name}</h2>
                <div className="text-blue-600 font-semibold mb-4 text-sm">{selectedPersonForDetail.title}</div>
                <div className="text-xs text-slate-500 mb-1">Sicil: {selectedPersonForDetail.personnelNo}</div>
                <div className="text-xs text-slate-500">TC: {selectedPersonForDetail.tc}</div>
                
                <div className="mt-6 pt-5 border-t border-slate-200">
                   <div className="flex justify-between mb-3 text-sm">
                      <span className="text-slate-500">Derece/Kademe:</span>
                      <span className="font-bold text-slate-800">{selectedPersonForDetail.degree}/{selectedPersonForDetail.level}</span>
                   </div>
                   <div className="flex justify-between mb-3 text-sm">
                      <span className="text-slate-500">Branş:</span>
                      <span className="font-bold text-slate-800">{selectedPersonForDetail.branch}</span>
                   </div>
                   <div className="flex justify-between text-sm">
                      <span className="text-slate-500">Başlama:</span>
                      <span className="font-bold text-slate-800">{formatDateTR(selectedPersonForDetail.startDate)}</span>
                   </div>
                </div>
              </div>

              {/* Sağ Sütun: Detaylı Bilgiler */}
              <div className="md:col-span-2 space-y-8">
                <div>
                  <h4 className="border-b border-slate-100 pb-2 text-slate-800 font-semibold flex items-center gap-2 mb-4">📞 İletişim Bilgileri</h4>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold mb-1">Telefon</div>
                      <div className="text-sm font-medium text-slate-800">{selectedPersonForDetail.phone || '-'}</div>
                    </div>
                    <div>
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold mb-1">E-Posta</div>
                      <div className="text-sm font-medium text-slate-800">{selectedPersonForDetail.email || '-'}</div>
                    </div>
                    <div className="col-span-2">
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold mb-1">Adres</div>
                      <div className="text-sm font-medium text-slate-800">{selectedPersonForDetail.address || '-'}</div>
                    </div>
                  </div>
                </div>

                <div>
                  <h4 className="border-b border-slate-100 pb-2 text-slate-800 font-semibold flex items-center gap-2 mb-4">🏦 Özlük ve Finans</h4>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold mb-1">İstihdam Türü</div>
                      <div className="text-sm font-medium text-slate-800">{selectedPersonForDetail.employmentType}</div>
                    </div>
                    <div>
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold mb-1">Sendika</div>
                      <div className="text-sm font-medium text-slate-800">{selectedPersonForDetail.union || 'Yok'}</div>
                    </div>
                    <div className="col-span-2">
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold mb-1">IBAN</div>
                      <div className="text-sm font-medium text-slate-800 font-mono">{selectedPersonForDetail.iban || '-'}</div>
                    </div>
                    <div>
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold mb-1">Medeni Durum</div>
                      <div className="text-sm font-medium text-slate-800">{selectedPersonForDetail.maritalStatus} ({selectedPersonForDetail.childrenCount} Çocuk)</div>
                    </div>
                    <div>
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold mb-1">Eğitim</div>
                      <div className="text-sm font-medium text-slate-800">{selectedPersonForDetail.education}</div>
                    </div>
                  </div>
                </div>

                <div>
                  <h4 className="border-b border-slate-100 pb-2 text-slate-800 font-semibold flex items-center gap-2 mb-4">📜 İzin Geçmişi (Son Kayıtlar)</h4>
                  <div>
                    {selectedPersonForDetail.leaveHistory.length === 0 ? (
                      <div className="text-sm text-slate-400 italic">Henüz izin kaydı bulunmuyor.</div>
                    ) : (
                      <table className="w-full text-left border-collapse text-sm">
                        <thead>
                          <tr className="text-slate-500 border-b border-slate-100">
                            <th className="py-2 font-semibold">Tür</th>
                            <th className="py-2 font-semibold">Başlangıç</th>
                            <th className="py-2 font-semibold">Bitiş</th>
                            <th className="py-2 font-semibold text-center">Gün</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                          {selectedPersonForDetail.leaveHistory.slice(-5).reverse().map(l => (
                            <tr key={l.id}>
                              <td className="py-2.5 text-slate-800">{l.type}</td>
                              <td className="py-2.5 text-slate-600">{formatDateTR(l.startDate)}</td>
                              <td className="py-2.5 text-slate-600">{formatDateTR(l.endDate)}</td>
                              <td className="py-2.5 text-center font-bold text-slate-800">{l.duration}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                </div>
              </div>
            </div>

            <div className="mt-8 pt-6 border-t border-slate-100 flex justify-end">
              <button onClick={() => setShowDetailModal(false)} className="px-8 py-2.5 bg-slate-800 text-white hover:bg-slate-900 rounded-xl text-sm font-semibold transition-colors shadow-sm">Kapat</button>
            </div>
          </div>
        </div>
      )}

      {showPreviewModal && (
        <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm flex justify-center items-center z-[10000] p-4">
          <div className="bg-white rounded-2xl w-full max-w-[1200px] max-h-[95vh] overflow-hidden shadow-2xl flex flex-col">
            <div className="no-print p-6 border-b border-slate-200 flex justify-between items-center bg-slate-50">
              <h2 className="text-xl font-bold text-slate-800">📋 Resmi Çizelge Önizleme</h2>
              <div className="flex gap-3">
                <button 
                  onClick={async () => {
                    if (pdfLoading) return;
                    const orientation: 'portrait' | 'landscape' = officialForm === 'stepPromotion' ? 'landscape' : 'portrait';
                    const filename = officialForm === 'stepPromotion' ? 'Kademe_Terfi_Formu.pdf' : 'Maas_Degisikligi_Formu.pdf';

                    setPdfLoading(true);
                    document.body.classList.add('printing');
                    const cleanup = applyPrintOrientation(orientation);
                    await new Promise(resolve => setTimeout(resolve, 150));

                    try {
                      if (window.printer) {
                        const res = await window.printer.savePDF({ orientation, filename });
                        if (res.success) {
                          addToast('PDF Kaydedildi', `Dosya: ${res.path}`, 'success');
                        } else if (!res.canceled) {
                          addToast('Hata', res.error || 'PDF oluşturulamadı.', 'error');
                        }
                      } else {
                        window.print();
                      }
                    } catch (error: any) {
                      console.error('PDF oluşturma hatası:', error);
                      addToast('Hata', error?.message || 'PDF oluşturulurken bir hata oluştu.', 'error');
                    } finally {
                      cleanup();
                      document.body.classList.remove('printing');
                      setPdfLoading(false);
                    }
                  }} 
                  disabled={pdfLoading}
                  className="px-6 py-2.5 bg-emerald-600 text-white hover:bg-emerald-700 rounded-xl font-semibold inline-flex items-center gap-2 transition-colors shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {pdfLoading ? '⏳ PDF Hazırlanıyor...' : '📥 PDF İndir'}
                </button>
                <button 
                  onClick={async () => {
                    const orientation: 'portrait' | 'landscape' = officialForm === 'stepPromotion' ? 'landscape' : 'portrait';
                    // Modal kapanmadan önce yazdırma yapılmalı; modal görünüyor olsa da
                    // .printing class'ı sayesinde no-print elemanlar gizlenir.
                    document.body.classList.add('printing');
                    const cleanup = applyPrintOrientation(orientation);
                    await new Promise(resolve => setTimeout(resolve, 150));
                    try {
                      if (window.printer) {
                        await window.printer.print({ orientation });
                      } else {
                        window.print();
                      }
                    } catch (e: any) {
                      console.error('Yazdırma hatası:', e);
                    } finally {
                      cleanup();
                      document.body.classList.remove('printing');
                    }
                  }} 
                  className="px-6 py-2.5 bg-slate-800 text-white hover:bg-slate-900 rounded-xl font-semibold inline-flex items-center gap-2 transition-colors shadow-sm"
                >
                  🖨 Yazdır
                </button>
                <button 
                  onClick={() => setShowPreviewModal(false)} 
                  className="px-6 py-2.5 bg-slate-200 text-slate-700 hover:bg-slate-300 rounded-xl font-semibold transition-colors"
                >
                  Kapat
                </button>
              </div>
            </div>
            
            <div className="flex-1 overflow-y-auto overflow-x-auto p-8 bg-slate-100">
              <div className="bg-white p-8 rounded-lg shadow-lg">
                {officialForm === 'salaryChange' && (
                  <div className="w-full overflow-x-auto">
                    <div className="border-2 border-black p-6 text-black font-sans min-w-[800px] text-[10px] leading-tight print-portrait">
                     <div className="text-center font-bold text-[11px] mb-3 uppercase">{settings.provinceTitle || 'MAMAK MÜDÜRLÜĞÜNE'}</div>
                     <div className="mb-3">
                        <div className="text-[9px]"><span className="font-bold">Kurumun Adı:</span> {settings.schoolTitle}</div>
                        <div className="text-[9px] mt-1"><span className="font-bold">Ayı Yılı:</span> {MONTHS[new Date().getMonth()]} {new Date().getFullYear().toString().slice(-2)}</div>
                     </div>
                     <div className="bg-yellow-300 border border-black text-center font-bold p-1 text-[10px] uppercase">Terfiler</div>
                     <table className="w-full border-collapse border border-black text-[9px]">
                        <thead>
                          <tr className="bg-white">
                            <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                            <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                            <th colSpan={3} className="border border-black p-1 font-bold text-center">Yeni</th>
                            <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Yeni Özel Hizmet Tazm.</th>
                            <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Geçerlilik Tarihi</th>
                            <th rowSpan={2} className="border border-black p-1 font-bold text-center align-middle">Açıklama</th>
                          </tr>
                          <tr className="bg-white">
                            <th className="border border-black p-1 font-bold text-center">Derece</th>
                            <th className="border border-black p-1 font-bold text-center">Kademe</th>
                            <th className="border border-black p-1 font-bold text-center">Ek Göst.</th>
                          </tr>
                        </thead>
                        <tbody>
                           {periodPromotions.length + manualPromotions.length === 0 ? (
                             <tr>
                               <td colSpan={8} className="border border-black p-4 text-center text-slate-500">
                                 - Maaş Döneminde (15-14) Terfi Kaydı Bulunmamaktadır -
                               </td>
                             </tr>
                           ) : (
                             <>
                               {periodPromotions.map((item, idx) => {
                                 const dateKey = `promo-${item.person.id}`;
                                 const displayDate = promotionDateOverrides[dateKey] || item.date;
                                 return (
                                 <tr key={`auto-${idx}`}>
                                   <td className="border border-black p-1 text-center">{item.person.personnelNo}</td>
                                   <td className="border border-black p-1">{item.person.name}</td>
                                   <td className="border border-black p-1 text-center">{item.nextDegree}</td>
                                   <td className="border border-black p-1 text-center">{item.nextLevel}</td>
                                   <td className="border border-black p-1"></td>
                                   <td className="border border-black p-1"></td>
                                   <td className="border border-black p-1 text-center">{displayDate}</td>
                                   <td className="border border-black p-1">kademe</td>
                                 </tr>
                                 );
                               })}
                               {manualPromotions.map(row => (
                                 <tr key={`manual-${row.id}`}>
                                   <td className="border border-black p-1 text-center">{row.personnelNo}</td>
                                   <td className="border border-black p-1">{row.name}</td>
                                   <td className="border border-black p-1 text-center">{row.newDegree}</td>
                                   <td className="border border-black p-1 text-center">{row.newLevel}</td>
                                   <td className="border border-black p-1"></td>
                                   <td className="border border-black p-1"></td>
                                   <td className="border border-black p-1 text-center">{row.date}</td>
                                   <td className="border border-black p-1">{row.description}</td>
                                 </tr>
                               ))}
                               {Array.from({ length: Math.max(0, 5 - (periodPromotions.length + manualPromotions.length)) }).map((_, i) => (
                                 <tr key={`empty-${i}`}>
                                   <td colSpan={8} className="border border-black p-2"></td>
                                 </tr>
                               ))}
                             </>
                           )}
                        </tbody>
                     </table>
                   <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">Göreve Başlama</div>
                   <table className="w-full border-collapse border border-black text-[9px]">
                      <thead>
                         <tr className="bg-white">
                            <th className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Derece</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Kademe</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Özel Hiz. Taz.</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Yan Ödeme</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Ek Ödeme Oranı</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">İban No</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Başlama Tarihi</th>
                         </tr>
                      </thead>
                      <tbody>
                        {periodDutyStarters.length === 0 ? (
                          Array.from({length: 3}).map((_, i) => (<tr key={i}><td colSpan={9} className="border border-black p-2"></td></tr>))
                        ) : (
                          <>
                            {periodDutyStarters.map((item, idx) => (
                              <tr key={idx}>
                                <td className="border border-black p-1 text-center">{item.person.personnelNo}</td>
                                <td className="border border-black p-1">{item.person.name}</td>
                                <td className="border border-black p-1 text-center">{item.degree}</td>
                                <td className="border border-black p-1 text-center">{item.level}</td>
                                <td className="border border-black p-1"></td>
                                <td className="border border-black p-1"></td>
                                <td className="border border-black p-1"></td>
                                <td className="border border-black p-1 text-[8px]">{item.person.iban}</td>
                                <td className="border border-black p-1 text-center">{item.date}</td>
                              </tr>
                            ))}
                            {Array.from({length: Math.max(0, 3 - periodDutyStarters.length)}).map((_, i) => (
                              <tr key={`empty-${i}`}><td colSpan={9} className="border border-black p-2"></td></tr>
                            ))}
                          </>
                        )}
                      </tbody>
                   </table>
                   <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">Görevden Ayrılma (Naklen Atama, Emekli, Aylıksız İzin, İhraç, Açığa Alınma)</div>
                   <table className="w-full border-collapse border border-black text-[9px]">
                      <thead>
                         <tr className="bg-white">
                            <th className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Ayrıldığı Tarih</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Gittiği Yer</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Açıklama</th>
                         </tr>
                      </thead>
                      <tbody>
                         {(() => {
                           const ayrilmaList = dutyChanges.filter(dc => dc.type === 'Görevden Ayrılma' || dc.type === 'Aylıksız İzin');
                           if (ayrilmaList.length === 0) {
                             return Array.from({length: 3}).map((_, i) => (<tr key={i}><td colSpan={5} className="border border-black p-2"></td></tr>));
                           }
                           return (
                             <>
                               {ayrilmaList.map(dc => {
                                 const p = personnel.find(x => x.id === dc.personnelId);
                                 return (
                                   <tr key={dc.id}>
                                     <td className="border border-black p-1 text-center">{p?.personnelNo}</td>
                                     <td className="border border-black p-1">{p?.name}</td>
                                     <td className="border border-black p-1 text-center">{formatDateTR(dc.date)}</td>
                                     <td className="border border-black p-1">{dc.destination || (dc.type === 'Aylıksız İzin' ? '—' : '')}</td>
                                     <td className="border border-black p-1">{dc.description || (dc.type === 'Aylıksız İzin' ? 'Aylıksız İzin' : '')}</td>
                                   </tr>
                                 );
                               })}
                               {Array.from({length: Math.max(0, 3 - ayrilmaList.length)}).map((_, i) => (
                                 <tr key={`empty-${i}`}><td colSpan={5} className="border border-black p-2"></td></tr>
                               ))}
                             </>
                           );
                         })()}
                      </tbody>
                   </table>
                  <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">İcra, Nafaka, Kefalet Giriş Aidatı, Para Cezası, vb. Kesintisi</div>
                  <table className="w-full border-collapse border border-black text-[9px]">
                     <thead>
                        <tr className="bg-white">
                           <th className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                           <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                           <th className="border border-black p-1 font-bold text-center align-middle">Türü</th>
                           <th className="border border-black p-1 font-bold text-center align-middle">Dosya No</th>
                           <th className="border border-black p-1 font-bold text-center align-middle">Tutar</th>
                           <th className="border border-black p-1 font-bold text-center align-middle">Açıklama</th>
                        </tr>
                     </thead>
                      <tbody>
                         {reportExceeders.map((item, idx) => (
                           <tr key={`report-${idx}`}>
                              <td className="border border-black p-1 text-center">{item.person.personnelNo}</td>
                              <td className="border border-black p-1">{item.person.name}</td>
                              <td className="border border-black p-1 text-center">Rapor</td>
                              <td className="border border-black p-1"></td>
                              <td className="border border-black p-1 text-center">{item.days} gün</td>
                              <td className="border border-black p-1">Kesinti</td>
                           </tr>
                         ))}
                         {Array.from({length: Math.max(0, 3 - reportExceeders.length)}).map((_, i) => (
                           <tr key={i}><td colSpan={6} className="border border-black p-2"></td></tr>
                         ))}
                      </tbody>
                   </table>
                   <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">Diğer Değişiklikler (Aile Durum Bildirimi, Dil Tazminatı, Kefalet, Aile Yardım Beyanı, Sakatlık İndirimi)</div>
                   <table className="w-full border-collapse border border-black text-[9px]">
                      <thead>
                         <tr className="bg-white">
                            <th className="border border-black p-1 font-bold text-center align-middle">Saymanlık No</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Yapılacak İşlemin</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Miktar/Oran</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Açıklama</th>
                         </tr>
                      </thead>
                      <tbody>
                         {otherChanges.map(oc => {
                           const p = personnel.find(x => x.id === oc.personnelId);
                           return (
                             <tr key={oc.id}>
                               <td className="border border-black p-1 text-center">{p?.personnelNo}</td>
                               <td className="border border-black p-1">{p?.name}</td>
                               <td className="border border-black p-1 text-center">{oc.processType}</td>
                               <td className="border border-black p-1 text-center">{oc.amount || '-'}</td>
                               <td className="border border-black p-1">{oc.description || '-'}</td>
                             </tr>
                           );
                         })}
                         {Array.from({length: Math.max(0, 3 - otherChanges.length)}).map((_, i) => (
                           <tr key={`empty-${i}`}><td colSpan={5} className="border border-black p-2"></td></tr>
                         ))}
                      </tbody>
                   </table>
                   <div className="bg-yellow-300 border border-black text-center font-bold p-1 mt-2 text-[10px] uppercase">Sendika Değişiklikleri</div>
                   <table className="w-full border-collapse border border-black text-[9px]">
                      <thead>
                         <tr className="bg-white">
                            <th className="border border-black p-1 font-bold text-center align-middle">Adı Soyadı</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">İşlem Türü</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Sendika Adı</th>
                            <th className="border border-black p-1 font-bold text-center align-middle">Tarih</th>
                         </tr>
                      </thead>
                      <tbody>
                         {unionChanges.length === 0 ? 
                           <tr><td colSpan={4} className="border border-black p-2 text-center text-slate-500">-</td></tr> : 
                           unionChanges.map((uc, idx) => {
                             const p = personnel.find(x => x.id === uc.personnelId);
                             return (
                               <tr key={idx}>
                                  <td className="border border-black p-1">{p?.name}</td>
                                  <td className="border border-black p-1 text-center">{uc.type}</td>
                                  <td className="border border-black p-1">{uc.unionName}</td>
                                  <td className="border border-black p-1 text-center">{formatDateTR(uc.date)}</td>
                               </tr>
                             );
                           })}
                         {Array.from({length: Math.max(0, 3 - unionChanges.length)}).map((_, i) => (
                           <tr key={i}><td colSpan={4} className="border border-black p-2"></td></tr>
                         ))}
                      </tbody>
                   </table>
                   <div className="mt-4 text-[9px]">
                      <div className="flex justify-between items-end">
                        <div className="text-left">
                          <div className="mb-8">Adı-Soyadı</div>
                          <div>Okul Müdürü</div>
                        </div>
                        <div className="text-right">
                          <div className="mb-8">Adı-Soyadı</div>
                          <div>Sayman</div>
                        </div>
                      </div>
                   </div>
                    </div>
                  </div>
                )}

                {officialForm === 'stepPromotion' && (
                  <div className="w-full overflow-x-auto">
                    <div className="border-2 border-black p-4 text-black font-sans min-w-[1200px] text-[9px] leading-tight print-landscape">
                    <div className="text-center font-bold text-[11px] mb-2 uppercase">Milli Eğitim Bakanlığı Personeli (Branş) Öğretmenlerine Ait Kademe Terfi Onayı</div>
                    <div className="flex justify-between items-center font-bold mb-2 text-[9px]">
                      <div>İL : ANKARA</div>
                      <div className="text-right">EK-1 FORM</div>
                    </div>
                    <div className="mb-2 text-[9px]"><span className="font-bold">KURUM :</span> {settings.schoolTitle}</div>
                    <table className="w-full border-collapse border border-black text-[8px]">
                      <thead>
                        <tr className="bg-white">
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Sıra No</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">T.C.Kimlik No</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Görev Yaptığı Okul veya Kurum</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Adı ve Soyadı</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Sınıfı</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Mezuniyet</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Unvanı</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Branşı</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">Kadro Derecesi</th>
                          <th colSpan={3} className="border border-black p-0.5 font-bold text-center">ESKİ DURUMU</th>
                          <th colSpan={3} className="border border-black p-0.5 font-bold text-center">YENİ DURUMU</th>
                          <th rowSpan={2} className="border border-black p-0.5 font-bold text-center align-middle">AÇIKLAMA</th>
                        </tr>
                        <tr className="bg-white">
                          <th className="border border-black p-0.5 font-bold text-center">Maaş Derecesi</th>
                          <th className="border border-black p-0.5 font-bold text-center">Kademesi</th>
                          <th className="border border-black p-0.5 font-bold text-center">Bu Kademeyi Aldığı Tarih</th>
                          <th className="border border-black p-0.5 font-bold text-center">Maaş Derecesi</th>
                          <th className="border border-black p-0.5 font-bold text-center">Kademesi</th>
                          <th className="border border-black p-0.5 font-bold text-center">Bu Kademeyi Aldığı Tarih</th>
                        </tr>
                      </thead>
                      <tbody>
                        {periodPromotions.length + manualPromotions.length === 0 ? (
                          <tr><td colSpan={16} className="border border-black p-2 text-center text-slate-500">-</td></tr>
                        ) : (
                          <>
                            {periodPromotions.map((item, idx) => {
                              const dateKey = `promo-${item.person.id}`;
                              const oldDateKey = `old-promo-${item.person.id}`;
                              const displayDate = promotionDateOverrides[dateKey] || item.date;
                              const oldDisplayDate = oldPromotionDateOverrides[oldDateKey] || item.person.lastPromotionDate || '';
                              return (
                              <tr key={`ek1-auto-${idx}`}>
                                <td className="border border-black p-0.5 text-center text-[7px]">{idx + 1}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.person.tc}</td>
                                <td className="border border-black p-0.5 text-[7px]">{settings.schoolTitle}</td>
                                <td className="border border-black p-0.5 text-[7px]">{item.person.name}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]"></td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.person.education}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.person.title}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.person.branch}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.person.degree}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.person.degree}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.person.level}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{oldDisplayDate}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.nextDegree}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{item.nextLevel}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{displayDate}</td>
                                <td className="border border-black p-0.5 text-[7px]">Kademe</td>
                              </tr>
                              );
                            })}
                            {manualPromotions.map((row, idx) => (
                              <tr key={row.id}>
                                <td className="border border-black p-0.5 text-center text-[7px]">{periodPromotions.length + idx + 1}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.tc || ''}</td>
                                <td className="border border-black p-0.5 text-[7px]">{settings.schoolTitle}</td>
                                <td className="border border-black p-0.5 text-[7px]">{row.name}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]"></td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.education || ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.title || ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.branch || ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.oldDegree ?? ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.oldDegree ?? ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.oldLevel ?? ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.oldPromotionDate || ''}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.newDegree}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.newLevel}</td>
                                <td className="border border-black p-0.5 text-center text-[7px]">{row.date}</td>
                                <td className="border border-black p-0.5 text-[7px]">{row.description}</td>
                              </tr>
                            ))}
                          </>
                        )}
                      </tbody>
                    </table>
                    <div className="mt-3 text-[8px]">
                      <div className="mb-2">Yukarıda durumu belirtilen {periodPromotions.length + manualPromotions.length === 1 ? 'bir (1)' : `${periodPromotions.length + manualPromotions.length} (${['', 'bir', 'iki', 'üç', 'dört', 'beş', 'altı', 'yedi', 'sekiz', 'dokuz', 'on'][periodPromotions.length + manualPromotions.length] || periodPromotions.length + manualPromotions.length})`} öğretmenin 29/06/1984 tarih ve 18446 sayılı Resmi Gazetede yayınlanan 241 Sayılı Kanun Hükmündeki kararname gereğince kademe/derece terfisini tasviplerinize arz ederim.</div>
                      <div className="flex justify-between items-end mt-6">
                        <div></div>
                        <div className="text-right">
                          <div className="mb-8">{settings.principalName || 'Adı-Soyadı'}</div>
                          <div>Okul Müdürü</div>
                        </div>
                      </div>
                    </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {showSchedulePreview && previewSchedule && (
        <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm flex justify-center items-center z-[10000] p-4">
          <div className="bg-white rounded-2xl w-full max-w-[95vw] max-h-[95vh] overflow-hidden shadow-2xl flex flex-col">
            <div className="p-6 border-b border-slate-200 flex justify-between items-center bg-slate-50">
              <div>
                <h2 className="text-xl font-bold text-slate-800">📋 {previewSchedule.type} Önizleme</h2>
                <p className="text-sm text-slate-600 mt-1">{previewSchedule.term}</p>
              </div>
              <div className="flex gap-3">
                <button 
                  onClick={() => downloadFile(previewSchedule.fileData, previewSchedule.fileName)} 
                  className="px-6 py-2.5 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-semibold inline-flex items-center gap-2 transition-colors shadow-sm"
                >
                  <Download size={16} /> İndir
                </button>
                <button 
                  onClick={() => {
                    setShowSchedulePreview(false);
                    setPreviewSchedule(null);
                  }} 
                  className="px-6 py-2.5 bg-slate-200 text-slate-700 hover:bg-slate-300 rounded-xl font-semibold transition-colors"
                >
                  Kapat
                </button>
              </div>
            </div>
            
            <div className="flex-1 overflow-auto p-8 bg-slate-100">
              <div className="bg-white rounded-lg shadow-lg max-w-full mx-auto">
                {previewSchedule.fileName.toLowerCase().endsWith('.pdf') ? (
                  <iframe
                    src={previewSchedule.fileData}
                    className="w-full h-[75vh] rounded-lg"
                    title="PDF Önizleme"
                  />
                ) : previewSchedule.fileName.match(/\.(jpg|jpeg|png|gif|webp)$/i) ? (
                  <div className="p-4">
                    <img 
                      src={previewSchedule.fileData} 
                      alt={previewSchedule.fileName}
                      className="max-w-full h-auto mx-auto rounded-lg shadow-md"
                      style={{ maxHeight: '75vh' }}
                    />
                  </div>
                ) : previewSchedule.fileName.match(/\.(xlsx?|csv)$/i) ? (
                  <div className="p-8 text-center">
                    <div className="inline-flex items-center justify-center w-24 h-24 bg-emerald-100 rounded-full mb-4">
                      <FileSpreadsheet size={48} className="text-emerald-600" />
                    </div>
                    <h3 className="text-lg font-bold text-slate-800 mb-2">Excel/CSV Dosyası</h3>
                    <p className="text-slate-600 mb-6">
                      Excel ve CSV dosyaları tarayıcıda önizlenemiyor.<br />
                      Dosyayı indirip bilgisayarınızda açabilirsiniz.
                    </p>
                    <button 
                      onClick={() => downloadFile(previewSchedule.fileData, previewSchedule.fileName)} 
                      className="px-8 py-3 bg-emerald-600 text-white hover:bg-emerald-700 rounded-xl font-semibold inline-flex items-center gap-2 transition-colors shadow-sm"
                    >
                      <Download size={18} /> Dosyayı İndir
                    </button>
                  </div>
                ) : (
                  <div className="p-8 text-center">
                    <div className="inline-flex items-center justify-center w-24 h-24 bg-slate-100 rounded-full mb-4">
                      <FileText size={48} className="text-slate-400" />
                    </div>
                    <h3 className="text-lg font-bold text-slate-800 mb-2">Dosya Önizlenemiyor</h3>
                    <p className="text-slate-600 mb-6">
                      Bu dosya türü ({previewSchedule.fileName.split('.').pop()?.toUpperCase()}) tarayıcıda önizlenemiyor.<br />
                      Dosyayı indirip bilgisayarınızda açabilirsiniz.
                    </p>
                    <button 
                      onClick={() => downloadFile(previewSchedule.fileData, previewSchedule.fileName)} 
                      className="px-8 py-3 bg-blue-600 text-white hover:bg-blue-700 rounded-xl font-semibold inline-flex items-center gap-2 transition-colors shadow-sm"
                    >
                      <Download size={18} /> Dosyayı İndir
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Yeni Eğitim-Öğretim Yılı Oluşturma Dialog */}
      {showNewYearDialog && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-center items-center z-[10001] p-4">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl">
            <div className="p-6 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
              <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <GraduationCap size={20} className="text-blue-500" />
                Yeni Eğitim-Öğretim Yılı
              </h2>
              <button onClick={() => setShowNewYearDialog(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                <X size={22} />
              </button>
            </div>
            <div className="p-6 space-y-4">
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-2">
                  Eğitim-Öğretim Yılı
                  <span className="ml-2 text-xs font-normal text-slate-400">Örn: 2025-2026</span>
                </label>
                <input
                  type="text"
                  value={newYearInput}
                  onChange={(e) => setNewYearInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      const val = newYearInput.trim();
                      if (!val) return;
                      if (!/^\d{4}-\d{4}$/.test(val)) {
                        addToast('Hata', 'Lütfen "YYYY-YYYY" formatında girin. Örn: 2025-2026', 'error');
                        return;
                      }
                      const [y1, y2] = val.split('-').map(Number);
                      if (y2 !== y1 + 1) {
                        addToast('Hata', 'İkinci yıl, birinci yıldan 1 fazla olmalıdır. Örn: 2025-2026', 'error');
                        return;
                      }
                      if (studentStats.find(s => s.id === val)) {
                        addToast('Uyarı', 'Bu eğitim-öğretim yılı zaten mevcut.', 'warning');
                        return;
                      }
                      const newStat: StudentStatistics = {
                        id: val,
                        academicYear: `${val} Eğitim-Öğretim Yılı`,
                        totalClasses: 0, totalStudents: 0, femaleStudents: 0, maleStudents: 0,
                        gradeClassInfo: [], foreignStudents: [], graduates: [], absentees: [],
                        createdDate: new Date().toISOString(), lastModified: new Date().toISOString()
                      };
                      setStudentStats([...studentStats, newStat]);
                      setSelectedAcademicYear(val);
                      setShowNewYearDialog(false);
                      addToast('Başarılı', `${val} Eğitim-Öğretim Yılı oluşturuldu.`, 'success');
                    }
                  }}
                  placeholder="2025-2026"
                  className="w-full px-4 py-2.5 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-lg font-semibold text-center tracking-wider"
                  autoFocus
                />
                <p className="mt-2 text-xs text-slate-400">İstediğiniz eğitim-öğretim yılını YYYY-YYYY formatında girebilirsiniz.</p>
              </div>

              {/* Hızlı seçim butonları */}
              <div>
                <p className="text-xs font-semibold text-slate-500 mb-2">Hızlı Seçim</p>
                <div className="flex flex-wrap gap-2">
                  {[-1, 0, 1, 2].map(offset => {
                    const y = new Date().getFullYear() + offset;
                    const label = `${y}-${y + 1}`;
                    const exists = !!studentStats.find(s => s.id === label);
                    return (
                      <button
                        key={label}
                        onClick={() => setNewYearInput(label)}
                        disabled={exists}
                        className={`px-3 py-1.5 rounded-lg text-xs font-semibold border transition-colors
                          ${newYearInput === label ? 'bg-blue-600 text-white border-blue-600' : 'bg-slate-50 text-slate-700 border-slate-200 hover:border-blue-400 hover:bg-blue-50'}
                          ${exists ? 'opacity-40 cursor-not-allowed' : ''}`}
                      >
                        {label} {exists ? '✓' : ''}
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
            <div className="p-4 border-t border-slate-100 flex gap-3 justify-end bg-slate-50 rounded-b-2xl">
              <button
                onClick={() => setShowNewYearDialog(false)}
                className="px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100 rounded-xl transition-colors"
              >
                İptal
              </button>
              <button
                onClick={() => {
                  const val = newYearInput.trim();
                  if (!val) return;
                  if (!/^\d{4}-\d{4}$/.test(val)) {
                    addToast('Hata', 'Lütfen "YYYY-YYYY" formatında girin. Örn: 2025-2026', 'error');
                    return;
                  }
                  const [y1, y2] = val.split('-').map(Number);
                  if (y2 !== y1 + 1) {
                    addToast('Hata', 'İkinci yıl, birinci yıldan 1 fazla olmalıdır. Örn: 2025-2026', 'error');
                    return;
                  }
                  if (studentStats.find(s => s.id === val)) {
                    addToast('Uyarı', 'Bu eğitim-öğretim yılı zaten mevcut.', 'warning');
                    return;
                  }
                  const newStat: StudentStatistics = {
                    id: val,
                    academicYear: `${val} Eğitim-Öğretim Yılı`,
                    totalClasses: 0, totalStudents: 0, femaleStudents: 0, maleStudents: 0,
                    gradeClassInfo: [], foreignStudents: [], graduates: [], absentees: [],
                    createdDate: new Date().toISOString(), lastModified: new Date().toISOString()
                  };
                  setStudentStats([...studentStats, newStat]);
                  setSelectedAcademicYear(val);
                  setShowNewYearDialog(false);
                  addToast('Başarılı', `${val} Eğitim-Öğretim Yılı oluşturuldu.`, 'success');
                }}
                className="px-6 py-2 bg-blue-600 text-white hover:bg-blue-700 rounded-xl text-sm font-semibold transition-colors shadow-sm flex items-center gap-2"
              >
                <Plus size={16} /> Oluştur
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Öğrenci İstatistik Modal */}
      {showStudentStatsModal && selectedAcademicYear && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-center items-center z-[10000] p-4">
          <div className="bg-white rounded-2xl w-full max-w-[800px] max-h-[90vh] overflow-hidden shadow-2xl flex flex-col">
            <div className="p-6 border-b border-slate-200 flex justify-between items-center bg-slate-50">
              <h2 className="text-xl font-bold text-slate-800">
                {studentStatsModalType === 'general' && '📊 Genel Bilgiler'}
                {studentStatsModalType === 'gradeClass' && '🏫 Kademe Bazlı Dağılım'}
                {studentStatsModalType === 'foreign' && '🌍 Yabancı Uyruklu Öğrenciler'}
                {studentStatsModalType === 'graduate' && '🎓 Mezun Öğrenciler'}
                {studentStatsModalType === 'absentee' && '⚠️ Sürekli Devamsızlar'}
              </h2>
              <button onClick={() => setShowStudentStatsModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                <X size={24} />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-6">
              {/* Genel Bilgiler Formu */}
              {studentStatsModalType === 'general' && (
                <div className="space-y-4">
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-sm font-semibold text-slate-700 mb-2">Toplam Şube Sayısı</label>
                      <input 
                        type="number" 
                        value={studentStats.find(s => s.id === selectedAcademicYear)?.totalClasses || 0}
                        onChange={(e) => {
                          const updated = studentStats.map(s => 
                            s.id === selectedAcademicYear 
                              ? { ...s, totalClasses: Number(e.target.value), lastModified: new Date().toISOString() }
                              : s
                          );
                          setStudentStats(updated);
                        }}
                        className="w-full px-4 py-2.5 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                        min="0"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-semibold text-slate-700 mb-2">
                        Toplam Öğrenci Sayısı
                        <span className="ml-2 text-xs font-normal text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded-full">Kız + Erkek otomatik</span>
                      </label>
                      <input 
                        type="number" 
                        value={studentStats.find(s => s.id === selectedAcademicYear)?.totalStudents || 0}
                        readOnly
                        className="w-full px-4 py-2.5 border border-emerald-200 rounded-xl bg-emerald-50 text-emerald-800 font-semibold cursor-not-allowed"
                        min="0"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-semibold text-slate-700 mb-2">Kız Öğrenci Sayısı</label>
                      <input 
                        type="number" 
                        value={studentStats.find(s => s.id === selectedAcademicYear)?.femaleStudents || 0}
                        onChange={(e) => {
                          const female = Number(e.target.value);
                          const updated = studentStats.map(s => 
                            s.id === selectedAcademicYear 
                              ? { ...s, femaleStudents: female, totalStudents: female + s.maleStudents, lastModified: new Date().toISOString() }
                              : s
                          );
                          setStudentStats(updated);
                        }}
                        className="w-full px-4 py-2.5 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-pink-500"
                        min="0"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-semibold text-slate-700 mb-2">Erkek Öğrenci Sayısı</label>
                      <input 
                        type="number" 
                        value={studentStats.find(s => s.id === selectedAcademicYear)?.maleStudents || 0}
                        onChange={(e) => {
                          const male = Number(e.target.value);
                          const updated = studentStats.map(s => 
                            s.id === selectedAcademicYear 
                              ? { ...s, maleStudents: male, totalStudents: s.femaleStudents + male, lastModified: new Date().toISOString() }
                              : s
                          );
                          setStudentStats(updated);
                        }}
                        className="w-full px-4 py-2.5 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-cyan-500"
                        min="0"
                      />
                    </div>
                  </div>
                  <div className="mt-4 p-4 bg-amber-50 rounded-xl border border-amber-200">
                    <p className="text-sm text-amber-800">
                      <strong>💡 Otomatik Hesaplama:</strong> Kız ve Erkek öğrenci sayılarını girdiğinizde <strong>Toplam Öğrenci</strong> otomatik hesaplanır. Kademe Dağılımı kartından veri girilmişse tüm toplamlar da otomatik güncellenir.
                    </p>
                  </div>
                </div>
              )}

              {/* Kademe Bazlı Dağılım Formu */}
              {studentStatsModalType === 'gradeClass' && (
                <div className="space-y-4">
                  {/* Excel Yükleme Alanı */}
                  <div className="p-4 bg-emerald-50 border-2 border-dashed border-emerald-300 rounded-xl">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <FileSpreadsheet size={24} className="text-emerald-600" />
                        <div>
                          <h4 className="font-semibold text-emerald-800">Excel'den Yükle</h4>
                          <p className="text-xs text-emerald-600">Kademe, Şube Sayısı, Toplam Öğrenci, Kız, Erkek sütunları</p>
                        </div>
                      </div>
                      <label className="cursor-pointer px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-sm font-semibold transition-colors flex items-center gap-2">
                        <Download size={16} className="rotate-180" />
                        Dosya Seç
                        <input 
                          type="file" 
                          accept=".xlsx,.xls"
                          className="hidden"
                          onChange={(e) => {
                            const file = e.target.files?.[0];
                            if (!file) return;
                            const reader = new FileReader();
                            reader.onload = (ev) => {
                              try {
                                const data = new Uint8Array(ev.target?.result as ArrayBuffer);
                                const wb = XLSX.read(data, { type: 'array' });
                                const ws = wb.Sheets[wb.SheetNames[0]];
                                const json = XLSX.utils.sheet_to_json(ws) as any[];
                                
                                const findCol = (row: any, ...keys: string[]) => {
                                  for (const key of Object.keys(row)) {
                                    const k = key.toLowerCase().replace(/\s+/g, '');
                                    for (const search of keys) {
                                      if (k.includes(search.toLowerCase())) return row[key];
                                    }
                                  }
                                  return null;
                                };
                                
                                const gradeClassInfo: GradeClassInfo[] = json.map((row, idx) => ({
                                  id: Date.now() + idx,
                                  grade: String(findCol(row, 'kademe', 'sınıf', 'sinif', 'grade') || ''),
                                  classCount: Number(findCol(row, 'şube', 'sube', 'class') || 0),
                                  studentCount: Number(findCol(row, 'toplam', 'total', 'öğrenci') || 0),
                                  femaleCount: Number(findCol(row, 'kız', 'kiz', 'female') || 0),
                                  maleCount: Number(findCol(row, 'erkek', 'male') || 0)
                                })).filter(g => g.grade);
                                
                                if (gradeClassInfo.length === 0) {
                                  addToast('Hata', 'Excel dosyasından kademe verisi okunamadı.', 'error');
                                  return;
                                }
                                
                                const updated = studentStats.map(s => {
                                  if (s.id === selectedAcademicYear) {
                                    const totalClasses = gradeClassInfo.reduce((sum, g) => sum + g.classCount, 0);
                                    const totalStudents = gradeClassInfo.reduce((sum, g) => sum + g.studentCount, 0);
                                    const femaleStudents = gradeClassInfo.reduce((sum, g) => sum + g.femaleCount, 0);
                                    const maleStudents = gradeClassInfo.reduce((sum, g) => sum + g.maleCount, 0);
                                    return { ...s, gradeClassInfo, totalClasses, totalStudents, femaleStudents, maleStudents, lastModified: new Date().toISOString() };
                                  }
                                  return s;
                                });
                                setStudentStats(updated);
                                addToast('Başarılı', `${gradeClassInfo.length} kademe verisi yüklendi.`, 'success');
                              } catch (err) {
                                addToast('Hata', 'Excel dosyası okunamadı.', 'error');
                              }
                            };
                            reader.readAsArrayBuffer(file);
                            e.target.value = '';
                          }}
                        />
                      </label>
                    </div>
                  </div>

                  <div className="grid grid-cols-5 gap-3 p-3 bg-slate-50 rounded-xl font-semibold text-sm text-slate-700">
                    <div>Kademe</div>
                    <div className="text-center">Şube</div>
                    <div className="text-center">
                      Toplam
                      <span className="block text-xs font-normal text-emerald-600">otomatik</span>
                    </div>
                    <div className="text-center">Kız</div>
                    <div className="text-center">Erkek</div>
                  </div>
                  <div className="p-3 bg-amber-50 rounded-xl border border-amber-200 text-xs text-amber-800">
                    💡 <strong>Kız</strong> ve <strong>Erkek</strong> sayılarını girince her satırın <strong>Toplamı</strong> otomatik hesaplanır. Aynı zamanda <strong>Genel Bilgiler</strong> kartındaki toplam şube, toplam öğrenci, kız ve erkek sayıları da otomatik güncellenir.
                  </div>
                  
                  {studentStats.find(s => s.id === selectedAcademicYear)?.gradeClassInfo.map((grade, idx) => (
                    <div key={grade.id} className="grid grid-cols-5 gap-3 items-center">
                      <input 
                        type="text" 
                        value={grade.grade}
                        onChange={(e) => {
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const updatedGrades = [...s.gradeClassInfo];
                              updatedGrades[idx] = { ...updatedGrades[idx], grade: e.target.value };
                              return { ...s, gradeClassInfo: updatedGrades, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        placeholder="5. Sınıf"
                        className="px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                      />
                      <input 
                        type="number" 
                        value={grade.classCount}
                        onChange={(e) => {
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const updatedGrades = [...s.gradeClassInfo];
                              updatedGrades[idx] = { ...updatedGrades[idx], classCount: Number(e.target.value) };
                              const totalClasses = updatedGrades.reduce((sum, g) => sum + g.classCount, 0);
                              return { ...s, gradeClassInfo: updatedGrades, totalClasses, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        className="px-3 py-2 border border-slate-200 rounded-lg text-sm text-center focus:outline-none focus:ring-2 focus:ring-blue-500"
                        min="0"
                      />
                      <input 
                        type="number" 
                        value={grade.studentCount}
                        readOnly
                        className="px-3 py-2 border border-emerald-200 rounded-lg text-sm text-center bg-emerald-50 text-emerald-800 font-semibold cursor-not-allowed"
                        min="0"
                        title="Kız + Erkek otomatik hesaplanır"
                      />
                      <input 
                        type="number" 
                        value={grade.femaleCount}
                        onChange={(e) => {
                          const female = Number(e.target.value);
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const updatedGrades = [...s.gradeClassInfo];
                              const male = updatedGrades[idx].maleCount;
                              updatedGrades[idx] = { ...updatedGrades[idx], femaleCount: female, studentCount: female + male };
                              const totalClasses = updatedGrades.reduce((sum, g) => sum + g.classCount, 0);
                              const totalStudents = updatedGrades.reduce((sum, g) => sum + g.studentCount, 0);
                              const femaleStudents = updatedGrades.reduce((sum, g) => sum + g.femaleCount, 0);
                              const maleStudents = updatedGrades.reduce((sum, g) => sum + g.maleCount, 0);
                              return { ...s, gradeClassInfo: updatedGrades, totalClasses, totalStudents, femaleStudents, maleStudents, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        className="px-3 py-2 border border-slate-200 rounded-lg text-sm text-center focus:outline-none focus:ring-2 focus:ring-pink-500"
                        min="0"
                      />
                      <div className="flex gap-2">
                        <input 
                          type="number" 
                          value={grade.maleCount}
                          onChange={(e) => {
                            const male = Number(e.target.value);
                            const updated = studentStats.map(s => {
                              if (s.id === selectedAcademicYear) {
                                const updatedGrades = [...s.gradeClassInfo];
                                const female = updatedGrades[idx].femaleCount;
                                updatedGrades[idx] = { ...updatedGrades[idx], maleCount: male, studentCount: female + male };
                                const totalClasses = updatedGrades.reduce((sum, g) => sum + g.classCount, 0);
                                const totalStudents = updatedGrades.reduce((sum, g) => sum + g.studentCount, 0);
                                const femaleStudents = updatedGrades.reduce((sum, g) => sum + g.femaleCount, 0);
                                const maleStudents = updatedGrades.reduce((sum, g) => sum + g.maleCount, 0);
                                return { ...s, gradeClassInfo: updatedGrades, totalClasses, totalStudents, femaleStudents, maleStudents, lastModified: new Date().toISOString() };
                              }
                              return s;
                            });
                            setStudentStats(updated);
                          }}
                          className="flex-1 px-3 py-2 border border-slate-200 rounded-lg text-sm text-center focus:outline-none focus:ring-2 focus:ring-cyan-500"
                          min="0"
                        />
                        <button 
                          onClick={() => {
                            const updated = studentStats.map(s => {
                              if (s.id === selectedAcademicYear) {
                                const filtered = s.gradeClassInfo.filter((_, i) => i !== idx);
                                const totalClasses = filtered.reduce((sum, g) => sum + g.classCount, 0);
                                const totalStudents = filtered.reduce((sum, g) => sum + g.studentCount, 0);
                                const femaleStudents = filtered.reduce((sum, g) => sum + g.femaleCount, 0);
                                const maleStudents = filtered.reduce((sum, g) => sum + g.maleCount, 0);
                                return { ...s, gradeClassInfo: filtered, totalClasses, totalStudents, femaleStudents, maleStudents, lastModified: new Date().toISOString() };
                              }
                              return s;
                            });
                            setStudentStats(updated);
                          }}
                          className="px-2 py-2 bg-red-50 text-red-600 hover:bg-red-100 rounded-lg transition-colors"
                          title="Sil"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                  ))}
                  
                  <button 
                    onClick={() => {
                      const updated = studentStats.map(s => {
                        if (s.id === selectedAcademicYear) {
                          const newGrade: GradeClassInfo = {
                            id: Date.now(),
                            grade: '',
                            classCount: 0,
                            studentCount: 0,
                            femaleCount: 0,
                            maleCount: 0
                          };
                          return { ...s, gradeClassInfo: [...s.gradeClassInfo, newGrade], lastModified: new Date().toISOString() };
                        }
                        return s;
                      });
                      setStudentStats(updated);
                    }}
                    className="w-full py-3 border-2 border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50 rounded-xl text-slate-600 hover:text-blue-600 font-semibold transition-all flex items-center justify-center gap-2"
                  >
                    <Plus size={18} /> Yeni Kademe Ekle
                  </button>
                </div>
              )}

              {/* Yabancı Uyruklu Öğrenciler Formu */}
              {studentStatsModalType === 'foreign' && (
                <div className="space-y-4">
                  {/* Excel Yükleme Alanı */}
                  <div className="p-4 bg-purple-50 border-2 border-dashed border-purple-300 rounded-xl">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <FileSpreadsheet size={24} className="text-purple-600" />
                        <div>
                          <h4 className="font-semibold text-purple-800">Excel'den Yükle</h4>
                          <p className="text-xs text-purple-600">Ülke, Cinsiyet, Öğrenci Sayısı sütunları</p>
                        </div>
                      </div>
                      <label className="cursor-pointer px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-sm font-semibold transition-colors flex items-center gap-2">
                        <Download size={16} className="rotate-180" />
                        Dosya Seç
                        <input 
                          type="file" 
                          accept=".xlsx,.xls"
                          className="hidden"
                          onChange={(e) => {
                            const file = e.target.files?.[0];
                            if (!file) return;
                            const reader = new FileReader();
                            reader.onload = (ev) => {
                              try {
                                const data = new Uint8Array(ev.target?.result as ArrayBuffer);
                                const wb = XLSX.read(data, { type: 'array' });
                                const ws = wb.Sheets[wb.SheetNames[0]];
                                const json = XLSX.utils.sheet_to_json(ws) as any[];
                                
                                const findCol = (row: any, ...keys: string[]) => {
                                  for (const key of Object.keys(row)) {
                                    const k = key.toLowerCase().replace(/\s+/g, '');
                                    for (const search of keys) {
                                      if (k.includes(search.toLowerCase())) return row[key];
                                    }
                                  }
                                  return null;
                                };
                                
                                const foreignStudents: ForeignStudent[] = json.map((row, idx) => {
                                  const genderVal = String(findCol(row, 'cinsiyet', 'gender') || 'Kız');
                                  return {
                                    id: Date.now() + idx,
                                    country: String(findCol(row, 'ülke', 'ulke', 'country') || ''),
                                    gender: (genderVal.toLowerCase().includes('erkek') ? 'Erkek' : 'Kız') as 'Kız' | 'Erkek',
                                    count: Number(findCol(row, 'sayı', 'sayi', 'count', 'öğrenci', 'ogrenci') || 0)
                                  };
                                }).filter(f => f.country);
                                
                                if (foreignStudents.length === 0) {
                                  addToast('Hata', 'Excel dosyasından yabancı uyruklu verisi okunamadı.', 'error');
                                  return;
                                }
                                
                                const updated = studentStats.map(s => {
                                  if (s.id === selectedAcademicYear) {
                                    return { ...s, foreignStudents, lastModified: new Date().toISOString() };
                                  }
                                  return s;
                                });
                                setStudentStats(updated);
                                addToast('Başarılı', `${foreignStudents.length} yabancı uyruklu verisi yüklendi.`, 'success');
                              } catch (err) {
                                addToast('Hata', 'Excel dosyası okunamadı.', 'error');
                              }
                            };
                            reader.readAsArrayBuffer(file);
                            e.target.value = '';
                          }}
                        />
                      </label>
                    </div>
                  </div>

                  <div className="grid grid-cols-4 gap-3 p-3 bg-slate-50 rounded-xl font-semibold text-sm text-slate-700">
                    <div>Ülke</div>
                    <div className="text-center">Cinsiyet</div>
                    <div className="text-center">Sayı</div>
                    <div></div>
                  </div>
                  
                  {studentStats.find(s => s.id === selectedAcademicYear)?.foreignStudents.map((foreign, idx) => (
                    <div key={foreign.id} className="grid grid-cols-4 gap-3 items-center">
                      <input 
                        type="text" 
                        value={foreign.country}
                        onChange={(e) => {
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const updatedForeign = [...s.foreignStudents];
                              updatedForeign[idx] = { ...updatedForeign[idx], country: e.target.value };
                              return { ...s, foreignStudents: updatedForeign, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        placeholder="Suriye"
                        className="px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                      />
                      <select 
                        value={foreign.gender}
                        onChange={(e) => {
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const updatedForeign = [...s.foreignStudents];
                              updatedForeign[idx] = { ...updatedForeign[idx], gender: e.target.value as 'Kız' | 'Erkek' };
                              return { ...s, foreignStudents: updatedForeign, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        className="px-3 py-2 border border-slate-200 rounded-lg text-sm text-center focus:outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        <option value="Kız">Kız</option>
                        <option value="Erkek">Erkek</option>
                      </select>
                      <input 
                        type="number" 
                        value={foreign.count}
                        onChange={(e) => {
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const updatedForeign = [...s.foreignStudents];
                              updatedForeign[idx] = { ...updatedForeign[idx], count: Number(e.target.value) };
                              return { ...s, foreignStudents: updatedForeign, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        className="px-3 py-2 border border-slate-200 rounded-lg text-sm text-center focus:outline-none focus:ring-2 focus:ring-blue-500"
                        min="0"
                      />
                      <button 
                        onClick={() => {
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const filtered = s.foreignStudents.filter((_, i) => i !== idx);
                              return { ...s, foreignStudents: filtered, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        className="px-3 py-2 bg-red-50 text-red-600 hover:bg-red-100 rounded-lg transition-colors"
                        title="Sil"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  ))}
                  
                  <button 
                    onClick={() => {
                      const updated = studentStats.map(s => {
                        if (s.id === selectedAcademicYear) {
                          const newForeign: ForeignStudent = {
                            id: Date.now(),
                            country: '',
                            gender: 'Kız',
                            count: 0
                          };
                          return { ...s, foreignStudents: [...s.foreignStudents, newForeign], lastModified: new Date().toISOString() };
                        }
                        return s;
                      });
                      setStudentStats(updated);
                    }}
                    className="w-full py-3 border-2 border-dashed border-slate-300 hover:border-purple-400 hover:bg-purple-50 rounded-xl text-slate-600 hover:text-purple-600 font-semibold transition-all flex items-center justify-center gap-2"
                  >
                    <Plus size={18} /> Yeni Kayıt Ekle
                  </button>
                </div>
              )}

              {/* Mezun Öğrenciler Formu */}
              {studentStatsModalType === 'graduate' && (
                <div className="space-y-4">
                  {studentStats.find(s => s.id === selectedAcademicYear)?.graduates.map((grad, idx) => (
                    <div key={grad.id} className="p-4 border border-slate-200 rounded-xl space-y-3">
                      <div className="grid grid-cols-2 gap-3">
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">Öğrenci Adı</label>
                          <input 
                            type="text" 
                            value={grad.name}
                            onChange={(e) => {
                              const updated = studentStats.map(s => {
                                if (s.id === selectedAcademicYear) {
                                  const updatedGrads = [...s.graduates];
                                  updatedGrads[idx] = { ...updatedGrads[idx], name: e.target.value };
                                  return { ...s, graduates: updatedGrads, lastModified: new Date().toISOString() };
                                }
                                return s;
                              });
                              setStudentStats(updated);
                            }}
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">Cinsiyet</label>
                          <select 
                            value={grad.gender}
                            onChange={(e) => {
                              const updated = studentStats.map(s => {
                                if (s.id === selectedAcademicYear) {
                                  const updatedGrads = [...s.graduates];
                                  updatedGrads[idx] = { ...updatedGrads[idx], gender: e.target.value as 'Kız' | 'Erkek' };
                                  return { ...s, graduates: updatedGrads, lastModified: new Date().toISOString() };
                                }
                                return s;
                              });
                              setStudentStats(updated);
                            }}
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                          >
                            <option value="Kız">Kız</option>
                            <option value="Erkek">Erkek</option>
                          </select>
                        </div>
                      </div>
                      <div className="grid grid-cols-3 gap-3">
                        <div className="flex items-center gap-2">
                          <input 
                            type="checkbox" 
                            checked={grad.tookExam}
                            onChange={(e) => {
                              const updated = studentStats.map(s => {
                                if (s.id === selectedAcademicYear) {
                                  const updatedGrads = [...s.graduates];
                                  updatedGrads[idx] = { ...updatedGrads[idx], tookExam: e.target.checked };
                                  return { ...s, graduates: updatedGrads, lastModified: new Date().toISOString() };
                                }
                                return s;
                              });
                              setStudentStats(updated);
                            }}
                            className="w-4 h-4 text-blue-600 rounded focus:ring-2 focus:ring-blue-500"
                          />
                          <label className="text-sm text-slate-700">Sınava Girdi</label>
                        </div>
                        <div className="flex items-center gap-2">
                          <input 
                            type="checkbox" 
                            checked={grad.passed || false}
                            disabled={!grad.tookExam}
                            onChange={(e) => {
                              const updated = studentStats.map(s => {
                                if (s.id === selectedAcademicYear) {
                                  const updatedGrads = [...s.graduates];
                                  updatedGrads[idx] = { ...updatedGrads[idx], passed: e.target.checked };
                                  return { ...s, graduates: updatedGrads, lastModified: new Date().toISOString() };
                                }
                                return s;
                              });
                              setStudentStats(updated);
                            }}
                            className="w-4 h-4 text-emerald-600 rounded focus:ring-2 focus:ring-emerald-500 disabled:opacity-50"
                          />
                          <label className="text-sm text-slate-700">Kazandı</label>
                        </div>
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">Kazandığı Okul</label>
                        <input 
                          type="text" 
                          value={grad.schoolName || ''}
                          disabled={!grad.passed}
                          onChange={(e) => {
                            const updated = studentStats.map(s => {
                              if (s.id === selectedAcademicYear) {
                                const updatedGrads = [...s.graduates];
                                updatedGrads[idx] = { ...updatedGrads[idx], schoolName: e.target.value };
                                return { ...s, graduates: updatedGrads, lastModified: new Date().toISOString() };
                              }
                              return s;
                            });
                            setStudentStats(updated);
                          }}
                          placeholder="Ankara Fen Lisesi"
                          className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 disabled:bg-slate-50 disabled:text-slate-400"
                        />
                      </div>
                      <button 
                        onClick={() => {
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const filtered = s.graduates.filter((_, i) => i !== idx);
                              return { ...s, graduates: filtered, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        className="w-full py-2 bg-red-50 text-red-600 hover:bg-red-100 rounded-lg transition-colors text-sm font-medium"
                      >
                        Kaydı Sil
                      </button>
                    </div>
                  ))}
                  
                  <button 
                    onClick={() => {
                      const updated = studentStats.map(s => {
                        if (s.id === selectedAcademicYear) {
                          const newGrad: GraduateStudent = {
                            id: Date.now(),
                            name: '',
                            gender: 'Kız',
                            tookExam: false
                          };
                          return { ...s, graduates: [...s.graduates, newGrad], lastModified: new Date().toISOString() };
                        }
                        return s;
                      });
                      setStudentStats(updated);
                    }}
                    className="w-full py-3 border-2 border-dashed border-slate-300 hover:border-amber-400 hover:bg-amber-50 rounded-xl text-slate-600 hover:text-amber-600 font-semibold transition-all flex items-center justify-center gap-2"
                  >
                    <Plus size={18} /> Yeni Mezun Ekle
                  </button>
                </div>
              )}

              {/* Sürekli Devamsızlar Formu */}
              {studentStatsModalType === 'absentee' && (
                <div className="space-y-4">
                  {/* Excel Yükleme Alanı */}
                  <div className="p-4 bg-red-50 border-2 border-dashed border-red-300 rounded-xl">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <FileSpreadsheet size={24} className="text-red-600" />
                        <div>
                          <h4 className="font-semibold text-red-800">Excel'den Yükle</h4>
                          <p className="text-xs text-red-600">Öğrenci Adı, Cinsiyet, Sınıf, Devamsızlık Günü, Sebep sütunları</p>
                        </div>
                      </div>
                      <label className="cursor-pointer px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-sm font-semibold transition-colors flex items-center gap-2">
                        <Download size={16} className="rotate-180" />
                        Dosya Seç
                        <input 
                          type="file" 
                          accept=".xlsx,.xls"
                          className="hidden"
                          onChange={(e) => {
                            const file = e.target.files?.[0];
                            if (!file) return;
                            const reader = new FileReader();
                            reader.onload = (ev) => {
                              try {
                                const data = new Uint8Array(ev.target?.result as ArrayBuffer);
                                const wb = XLSX.read(data, { type: 'array' });
                                const ws = wb.Sheets[wb.SheetNames[0]];
                                const json = XLSX.utils.sheet_to_json(ws) as any[];
                                
                                const findCol = (row: any, ...keys: string[]) => {
                                  for (const key of Object.keys(row)) {
                                    const k = key.toLowerCase().replace(/\s+/g, '');
                                    for (const search of keys) {
                                      if (k.includes(search.toLowerCase())) return row[key];
                                    }
                                  }
                                  return null;
                                };
                                
                                const absentees: AbsenteeStudent[] = json.map((row, idx) => {
                                  const genderVal = String(findCol(row, 'cinsiyet', 'gender') || 'Kız');
                                  return {
                                    id: Date.now() + idx,
                                    name: String(findCol(row, 'ad', 'isim', 'öğrenci', 'ogrenci', 'name') || ''),
                                    gender: (genderVal.toLowerCase().includes('erkek') ? 'Erkek' : 'Kız') as 'Kız' | 'Erkek',
                                    grade: String(findCol(row, 'sınıf', 'sinif', 'kademe', 'grade', 'class') || ''),
                                    absentDays: Number(findCol(row, 'devamsızlık', 'devamsizlik', 'gün', 'gun', 'day', 'absent') || 0),
                                    reason: String(findCol(row, 'sebep', 'neden', 'açıklama', 'reason') || '')
                                  };
                                }).filter(a => a.name);
                                
                                if (absentees.length === 0) {
                                  addToast('Hata', 'Excel dosyasından devamsızlık verisi okunamadı.', 'error');
                                  return;
                                }
                                
                                const updated = studentStats.map(s => {
                                  if (s.id === selectedAcademicYear) {
                                    return { ...s, absentees, lastModified: new Date().toISOString() };
                                  }
                                  return s;
                                });
                                setStudentStats(updated);
                                addToast('Başarılı', `${absentees.length} devamsızlık kaydı yüklendi.`, 'success');
                              } catch (err) {
                                addToast('Hata', 'Excel dosyası okunamadı.', 'error');
                              }
                            };
                            reader.readAsArrayBuffer(file);
                            e.target.value = '';
                          }}
                        />
                      </label>
                    </div>
                  </div>

                  {studentStats.find(s => s.id === selectedAcademicYear)?.absentees.map((absent, idx) => (
                    <div key={absent.id} className="p-4 border border-slate-200 rounded-xl space-y-3">
                      <div className="grid grid-cols-2 gap-3">
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">Öğrenci Adı</label>
                          <input 
                            type="text" 
                            value={absent.name}
                            onChange={(e) => {
                              const updated = studentStats.map(s => {
                                if (s.id === selectedAcademicYear) {
                                  const updatedAbsent = [...s.absentees];
                                  updatedAbsent[idx] = { ...updatedAbsent[idx], name: e.target.value };
                                  return { ...s, absentees: updatedAbsent, lastModified: new Date().toISOString() };
                                }
                                return s;
                              });
                              setStudentStats(updated);
                            }}
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">Cinsiyet</label>
                          <select 
                            value={absent.gender}
                            onChange={(e) => {
                              const updated = studentStats.map(s => {
                                if (s.id === selectedAcademicYear) {
                                  const updatedAbsent = [...s.absentees];
                                  updatedAbsent[idx] = { ...updatedAbsent[idx], gender: e.target.value as 'Kız' | 'Erkek' };
                                  return { ...s, absentees: updatedAbsent, lastModified: new Date().toISOString() };
                                }
                                return s;
                              });
                              setStudentStats(updated);
                            }}
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                          >
                            <option value="Kız">Kız</option>
                            <option value="Erkek">Erkek</option>
                          </select>
                        </div>
                      </div>
                      <div className="grid grid-cols-2 gap-3">
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">Sınıf</label>
                          <input 
                            type="text" 
                            value={absent.grade}
                            onChange={(e) => {
                              const updated = studentStats.map(s => {
                                if (s.id === selectedAcademicYear) {
                                  const updatedAbsent = [...s.absentees];
                                  updatedAbsent[idx] = { ...updatedAbsent[idx], grade: e.target.value };
                                  return { ...s, absentees: updatedAbsent, lastModified: new Date().toISOString() };
                                }
                                return s;
                              });
                              setStudentStats(updated);
                            }}
                            placeholder="5-A"
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">Devamsızlık Günü</label>
                          <input 
                            type="number" 
                            value={absent.absentDays}
                            onChange={(e) => {
                              const updated = studentStats.map(s => {
                                if (s.id === selectedAcademicYear) {
                                  const updatedAbsent = [...s.absentees];
                                  updatedAbsent[idx] = { ...updatedAbsent[idx], absentDays: Number(e.target.value) };
                                  return { ...s, absentees: updatedAbsent, lastModified: new Date().toISOString() };
                                }
                                return s;
                              });
                              setStudentStats(updated);
                            }}
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                            min="0"
                          />
                        </div>
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">Sebep (Opsiyonel)</label>
                        <textarea 
                          value={absent.reason || ''}
                          onChange={(e) => {
                            const updated = studentStats.map(s => {
                              if (s.id === selectedAcademicYear) {
                                const updatedAbsent = [...s.absentees];
                                updatedAbsent[idx] = { ...updatedAbsent[idx], reason: e.target.value };
                                return { ...s, absentees: updatedAbsent, lastModified: new Date().toISOString() };
                              }
                              return s;
                            });
                            setStudentStats(updated);
                          }}
                          rows={2}
                          className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                      </div>
                      <button 
                        onClick={() => {
                          const updated = studentStats.map(s => {
                            if (s.id === selectedAcademicYear) {
                              const filtered = s.absentees.filter((_, i) => i !== idx);
                              return { ...s, absentees: filtered, lastModified: new Date().toISOString() };
                            }
                            return s;
                          });
                          setStudentStats(updated);
                        }}
                        className="w-full py-2 bg-red-50 text-red-600 hover:bg-red-100 rounded-lg transition-colors text-sm font-medium"
                      >
                        Kaydı Sil
                      </button>
                    </div>
                  ))}
                  
                  <button 
                    onClick={() => {
                      const updated = studentStats.map(s => {
                        if (s.id === selectedAcademicYear) {
                          const newAbsent: AbsenteeStudent = {
                            id: Date.now(),
                            name: '',
                            gender: 'Kız',
                            grade: '',
                            absentDays: 0
                          };
                          return { ...s, absentees: [...s.absentees, newAbsent], lastModified: new Date().toISOString() };
                        }
                        return s;
                      });
                      setStudentStats(updated);
                    }}
                    className="w-full py-3 border-2 border-dashed border-slate-300 hover:border-red-400 hover:bg-red-50 rounded-xl text-slate-600 hover:text-red-600 font-semibold transition-all flex items-center justify-center gap-2"
                  >
                    <Plus size={18} /> Yeni Kayıt Ekle
                  </button>
                </div>
              )}
            </div>

            <div className="p-6 border-t border-slate-200 flex justify-end gap-3">
              <button 
                onClick={() => setShowStudentStatsModal(false)}
                className="px-6 py-2.5 bg-slate-200 text-slate-700 hover:bg-slate-300 rounded-xl font-semibold transition-colors"
              >
                Kapat
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// Uygulama açılmadan önce localStorage boşsa otomatik yedekten geri yükle.
// Bu sayede güncelleme/yeniden yükleme/yeni kurulum sonrası veriler korunur.
const bootstrap = async () => {
  let restoredFromBackup = false;
  try {
    const result = await ensureAutoRestoreBeforeRender();
    restoredFromBackup = result.restored;
    if (restoredFromBackup) {
      console.log(`[Auto-Restore] ${result.count} veri kümesi diskteki yedekten geri yüklendi.`);
      // İşaretle: App render edildiğinde toast gösterilebilir
      try {
        sessionStorage.setItem('__pts_just_restored__', String(result.count));
      } catch {}
    }
  } catch (e) {
    console.error('Otomatik geri yükleme başarısız:', e);
  }
  const root = createRoot(document.getElementById('root')!);
  root.render(<App />);
};

bootstrap();
