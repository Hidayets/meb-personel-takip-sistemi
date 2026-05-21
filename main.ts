import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron'
import path from 'path'
import fs from 'fs'
import os from 'os'
import { fileURLToPath } from 'url'
import { autoUpdater } from 'electron-updater'

// ES module'de __dirname yok; import.meta.url ile türetiyoruz
const __dirname = path.dirname(fileURLToPath(import.meta.url))

// Paketlenmiş uygulamada app.getAppPath() kullan (asar içindeki yolu doğru verir)
const APP_ROOT = app.isPackaged ? app.getAppPath() : path.join(__dirname, '..')

process.env.DIST = path.join(APP_ROOT, 'dist')
process.env.VITE_PUBLIC = app.isPackaged ? process.env.DIST : path.join(APP_ROOT, 'public')

// Yeni versiyon kurulumlarında userData yolu değişmesin diye sabit bir
// uygulama adı kullanıyoruz. Bu sayede %APPDATA%\Okul Takip Sistemi
// klasörü versiyondan versiyona aynı kalır ve localStorage/yedekler korunur.
// (app.setName ready event'inden ÖNCE çağrılmalı.)
const STABLE_APP_NAME = 'Okul Takip Sistemi'
try {
  app.setName(STABLE_APP_NAME)
} catch (e) {
  console.error('App adı sabitlenemedi:', e)
}

let win: BrowserWindow | null

// Auto Updater Ayarları
autoUpdater.autoDownload = true
autoUpdater.autoInstallOnAppQuit = true
// VITE_DEV_SERVER_URL değişkeni vite-plugin-electron tarafından otomatik tanımlanır
const VITE_DEV_SERVER_URL = process.env['VITE_DEV_SERVER_URL']

function createWindow() {
  win = new BrowserWindow({
    width: 1400,
    height: 900,
    title: "Okul Takip Sistemi",
    icon: app.isPackaged
      ? path.join(app.getAppPath(), 'build', 'icon.ico')
      : path.join(__dirname, '../build/icon.ico'),
    frame: false, // Çerçevesiz pencere (Custom Frame)
    titleBarStyle: 'hidden', // macOS uyumluluğu için
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  // Development modunda Vite sunucusunu yükle, değilse html dosyasını yükle
  if (VITE_DEV_SERVER_URL) {
    win.loadURL(VITE_DEV_SERVER_URL)
  } else {
    win.loadFile(path.join(process.env.DIST || '', 'index.html'))
  }
}

app.on('window-all-closed', () => {
  if ((process as any).platform !== 'darwin') {
    app.quit()
    win = null
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

// IPC Listeners for Window Controls
// Bu dinleyiciler render sürecinden (index.tsx) gelen komutları işler
ipcMain.on('window-minimize', (event) => {
  const webContents = event.sender;
  const win = BrowserWindow.fromWebContents(webContents);
  win?.minimize();
});

ipcMain.on('window-maximize', (event) => {
  const webContents = event.sender;
  const win = BrowserWindow.fromWebContents(webContents);
  if (win) {
    if (win.isMaximized()) {
      win.unmaximize();
    } else {
      win.maximize();
    }
  }
});

ipcMain.on('window-close', (event) => {
  const webContents = event.sender;
  const win = BrowserWindow.fromWebContents(webContents);
  win?.close();
});

app.whenReady().then(() => {
  // Yeni versiyon kurulumlarında userData yolu değişmiş olabilir.
  // Mevcut konumda yedek yoksa eski konumlardan otomatik kopyala.
  // Bu işlem renderer açılmadan önce yapıldığı için ensureAutoRestoreBeforeRender
  // mevcut userData/backups/auto-backup.json'ı bulur ve verileri geri yükler.
  try {
    migrateLegacyBackupsIfNeeded()
  } catch (e) {
    console.error('Yedek taşıma sırasında hata:', e)
  }

  createWindow()
  
  // Paketlenmiş uygulamada güncelleme kontrolü yap
  if (app.isPackaged) {
    // Uygulama başladıktan 3 saniye sonra güncelleme kontrolü
    setTimeout(() => {
      autoUpdater.checkForUpdates().catch(err => {
        console.log('Güncelleme kontrolü başarısız:', err)
      })
    }, 3000)
    
    // Her 30 dakikada bir güncelleme kontrolü
    setInterval(() => {
      autoUpdater.checkForUpdates().catch(err => {
        console.log('Güncelleme kontrolü başarısız:', err)
      })
    }, 30 * 60 * 1000)
  }
})

// Auto Updater Event Listeners
autoUpdater.on('checking-for-update', () => {
  console.log('Güncelleme kontrol ediliyor...')
  win?.webContents.send('update-status', { status: 'checking', message: 'Güncelleme kontrol ediliyor...' })
})

autoUpdater.on('update-available', (info) => {
  console.log('Güncelleme mevcut:', info.version)
  win?.webContents.send('update-status', { 
    status: 'available', 
    message: `Yeni sürüm mevcut: v${info.version}`,
    version: info.version
  })
})

autoUpdater.on('update-not-available', () => {
  console.log('Uygulama güncel.')
  win?.webContents.send('update-status', { status: 'not-available', message: 'Uygulama güncel.' })
})

autoUpdater.on('download-progress', (progress) => {
  const percent = Math.round(progress.percent)
  console.log(`İndiriliyor: ${percent}%`)
  win?.webContents.send('update-status', { 
    status: 'downloading', 
    message: `Güncelleme indiriliyor: ${percent}%`,
    percent
  })
})

autoUpdater.on('update-downloaded', (info) => {
  console.log('Güncelleme indirildi:', info.version)
  win?.webContents.send('update-status', { 
    status: 'downloaded', 
    message: `v${info.version} indirildi. Yeniden başlatıldığında yüklenecek.`,
    version: info.version
  })
})

autoUpdater.on('error', (err) => {
  console.log('Güncelleme hatası:', err)
  win?.webContents.send('update-status', { 
    status: 'error', 
    message: 'Güncelleme kontrolünde hata oluştu.'
  })
})

// Manuel güncelleme kontrolü için IPC
ipcMain.on('check-for-updates', () => {
  if (app.isPackaged) {
    autoUpdater.checkForUpdates().catch(err => {
      console.log('Manuel güncelleme kontrolü başarısız:', err)
    })
  } else {
    win?.webContents.send('update-status', { 
      status: 'dev-mode', 
      message: 'Geliştirme modunda güncelleme kontrolü devre dışı.'
    })
  }
})

// Güncellemeyi şimdi yükle
ipcMain.on('install-update', () => {
  autoUpdater.quitAndInstall(false, true)
})

// Uygulama versiyonunu al
ipcMain.handle('get-app-version', () => {
  return app.getVersion()
})

// ============================================================
// YEDEKLEME SİSTEMİ
// ============================================================
// Veriler localStorage'da tutulur, fakat ek güvence olarak
// her veri değişiminde userData/backups klasörüne JSON dosyası
// olarak yedeklenir. userData klasörü electron-updater tarafından
// güncelleme sırasında silinmediği için veriler korunmuş olur.
// ============================================================

const getBackupDir = (): string => {
  const dir = path.join(app.getPath('userData'), 'backups')
  try {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  } catch (e) {
    console.error('Yedek klasörü oluşturulamadı:', e)
  }
  return dir
}

const getDocumentsBackupDir = (): string => {
  const dir = path.join(app.getPath('documents'), 'Okul Takip Sistemi Yedekler')
  try {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  } catch (e) {
    console.error('Belgeler yedek klasörü oluşturulamadı:', e)
  }
  return dir
}

// Olası eski userData yolları (önceki versiyonlarda farklı app adı
// kullanıldıysa veriler buralarda kalmış olabilir). Yeni versiyon kurulumunda
// bu klasörlerdeki yedek dosyalarını mevcut userData'ya kopyalayarak
// kullanıcının verilerini koruruz.
const getAlternateUserDataDirs = (): string[] => {
  const candidates = new Set<string>()
  const home = os.homedir()
  const appData = process.env.APPDATA || path.join(home, 'AppData', 'Roaming')
  // Eski olası uygulama adları
  const oldNames = [
    'Okul Takip Sistemi',
    'meb-personel-takip',
    'okul-takip-sistemi',
    'OkulTakipSistemi',
    'MEB Personel Takip',
    'MEB-Personel-Takip',
  ]
  for (const name of oldNames) {
    candidates.add(path.join(appData, name))
  }
  // Mevcut userData'yı listeden çıkar (kendisi olmasın)
  const current = app.getPath('userData')
  candidates.delete(current)
  return Array.from(candidates).filter(p => {
    try {
      return fs.existsSync(p)
    } catch {
      return false
    }
  })
}

// Eski userData klasörlerinde auto-backup.json bul ve içeriğini döndür.
// En yeni mtime'a sahip olanı tercih eder.
const findLegacyAutoBackup = (): string | null => {
  try {
    const altDirs = getAlternateUserDataDirs()
    let bestPath: string | null = null
    let bestMtime = 0
    for (const ud of altDirs) {
      const candidates = [
        path.join(ud, 'backups', 'auto-backup.json'),
        path.join(ud, 'backups', 'auto-backup.previous.json'),
      ]
      for (const c of candidates) {
        try {
          if (fs.existsSync(c)) {
            const stat = fs.statSync(c)
            if (stat.mtimeMs > bestMtime) {
              bestMtime = stat.mtimeMs
              bestPath = c
            }
          }
        } catch {}
      }
    }
    if (bestPath) {
      return fs.readFileSync(bestPath, 'utf-8')
    }
  } catch (e) {
    console.error('Eski yedek arama hatası:', e)
  }
  return null
}

// Belgeler klasöründeki en güncel manuel yedek (okul-takip-yedek-*.json) dosyasını oku
const findLatestDocumentsBackup = (): string | null => {
  try {
    const dir = getDocumentsBackupDir()
    if (!fs.existsSync(dir)) return null
    const files = fs.readdirSync(dir)
      .filter(f => f.toLowerCase().endsWith('.json'))
      .map(f => {
        try {
          const full = path.join(dir, f)
          return { full, mtime: fs.statSync(full).mtimeMs }
        } catch {
          return null
        }
      })
      .filter((x): x is { full: string; mtime: number } => x !== null)
      .sort((a, b) => b.mtime - a.mtime)
    if (files.length > 0) {
      return fs.readFileSync(files[0].full, 'utf-8')
    }
  } catch (e) {
    console.error('Belgeler yedek arama hatası:', e)
  }
  return null
}

// Uygulama başlarken: mevcut userData/backups boşsa, eski userData
// yollarındaki yedek dosyalarını mevcut konuma kopyala. Bu sayede önceki
// versiyondan kalma veriler kaybolmaz.
const migrateLegacyBackupsIfNeeded = (): void => {
  try {
    const currentDir = getBackupDir()
    const currentAuto = path.join(currentDir, 'auto-backup.json')
    if (fs.existsSync(currentAuto)) return // mevcut yedek var, taşımaya gerek yok

    const altDirs = getAlternateUserDataDirs()
    for (const ud of altDirs) {
      const srcDir = path.join(ud, 'backups')
      if (!fs.existsSync(srcDir)) continue
      try {
        const files = fs.readdirSync(srcDir).filter(f => f.endsWith('.json'))
        for (const f of files) {
          const src = path.join(srcDir, f)
          const dst = path.join(currentDir, f)
          if (!fs.existsSync(dst)) {
            try {
              fs.copyFileSync(src, dst)
            } catch {}
          }
        }
        if (fs.existsSync(currentAuto)) {
          console.log('Eski yedekler yeni konuma taşındı:', srcDir)
          return
        }
      } catch {}
    }
  } catch (e) {
    console.error('Eski yedek taşıma hatası:', e)
  }
}

// Otomatik yedek dosyasını oku (başlangıçta geri yükleme için)
ipcMain.handle('backup:read-auto', () => {
  try {
    const filePath = path.join(getBackupDir(), 'auto-backup.json')
    if (fs.existsSync(filePath)) {
      return fs.readFileSync(filePath, 'utf-8')
    }
    const prev = path.join(getBackupDir(), 'auto-backup.previous.json')
    if (fs.existsSync(prev)) {
      return fs.readFileSync(prev, 'utf-8')
    }
    // Mevcut konumda yedek yoksa eski userData yollarını tara.
    // Bu, app.setName/app.getName veya productName değişikliği nedeniyle
    // userData yolu farklılaşmış olsa bile verilerin korunmasını sağlar.
    const legacy = findLegacyAutoBackup()
    if (legacy) return legacy
    // Son çare: Belgeler klasöründeki en güncel manuel yedek
    const docs = findLatestDocumentsBackup()
    if (docs) return docs
  } catch (e) {
    console.error('Otomatik yedek okuma hatası:', e)
  }
  return null
})

// Otomatik yedek yaz (her veri değişiminde, debounce'lu olarak çağrılır)
ipcMain.handle('backup:write-auto', (_event, data: string) => {
  try {
    const dir = getBackupDir()
    const filePath = path.join(dir, 'auto-backup.json')
    if (fs.existsSync(filePath)) {
      try {
        fs.copyFileSync(filePath, path.join(dir, 'auto-backup.previous.json'))
      } catch {}
    }
    fs.writeFileSync(filePath, data, 'utf-8')
    return { success: true, path: filePath }
  } catch (e: any) {
    console.error('Otomatik yedek yazma hatası:', e)
    return { success: false, error: e?.message || String(e) }
  }
})

// Günlük rotation yedek (7 günlük rotasyon)
ipcMain.handle('backup:write-daily', (_event, data: string) => {
  try {
    const dir = getBackupDir()
    const today = new Date().toISOString().slice(0, 10)
    const filePath = path.join(dir, `daily-${today}.json`)
    fs.writeFileSync(filePath, data, 'utf-8')

    // 7 günden eski daily yedekleri temizle
    const sevenDaysAgo = Date.now() - 7 * 24 * 60 * 60 * 1000
    try {
      const files = fs.readdirSync(dir).filter(f => f.startsWith('daily-'))
      files.forEach(f => {
        try {
          const full = path.join(dir, f)
          const stat = fs.statSync(full)
          if (stat.mtimeMs < sevenDaysAgo) fs.unlinkSync(full)
        } catch {}
      })
    } catch {}

    return { success: true, path: filePath }
  } catch (e: any) {
    return { success: false, error: e?.message || String(e) }
  }
})

// Belgeler klasörüne yedek dışa aktar (kullanıcı dosyayı taşıyabilsin)
ipcMain.handle('backup:export-to-documents', (_event, data: string) => {
  try {
    const dir = getDocumentsBackupDir()
    const ts = new Date().toISOString().replace(/[:.]/g, '-')
    const filePath = path.join(dir, `okul-takip-yedek-${ts}.json`)
    fs.writeFileSync(filePath, data, 'utf-8')
    return { success: true, path: filePath }
  } catch (e: any) {
    return { success: false, error: e?.message || String(e) }
  }
})

// Kullanıcının seçtiği dosyaya yedek kaydet
ipcMain.handle('backup:save-as', async (_event, data: string) => {
  try {
    const defaultName = `okul-takip-yedek-${new Date().toISOString().replace(/[:.]/g, '-')}.json`
    const result = await dialog.showSaveDialog({
      title: 'Yedek dosyasını kaydet',
      defaultPath: path.join(getDocumentsBackupDir(), defaultName),
      filters: [{ name: 'JSON Yedek', extensions: ['json'] }]
    })
    if (result.canceled || !result.filePath) {
      return { success: false, canceled: true }
    }
    fs.writeFileSync(result.filePath, data, 'utf-8')
    return { success: true, path: result.filePath }
  } catch (e: any) {
    return { success: false, error: e?.message || String(e) }
  }
})

// Kullanıcının seçtiği dosyadan yedek yükle
ipcMain.handle('backup:import', async () => {
  try {
    const result = await dialog.showOpenDialog({
      title: 'Yedek dosyasını seçin',
      filters: [{ name: 'JSON Yedek', extensions: ['json'] }],
      properties: ['openFile'],
      defaultPath: getDocumentsBackupDir()
    })
    if (result.canceled || result.filePaths.length === 0) {
      return { success: false, canceled: true }
    }
    const data = fs.readFileSync(result.filePaths[0], 'utf-8')
    return { success: true, data, path: result.filePaths[0] }
  } catch (e: any) {
    return { success: false, error: e?.message || String(e) }
  }
})

// Yedek klasörünü dosya gezgininde aç
ipcMain.handle('backup:open-folder', () => {
  const dir = getBackupDir()
  shell.openPath(dir).catch(() => {})
  return dir
})

// Belgeler yedek klasörünü aç
ipcMain.handle('backup:open-documents-folder', () => {
  const dir = getDocumentsBackupDir()
  shell.openPath(dir).catch(() => {})
  return dir
})

// Mevcut yedeklerin listesi
ipcMain.handle('backup:list', () => {
  try {
    const dir = getBackupDir()
    if (!fs.existsSync(dir)) return []
    return fs.readdirSync(dir)
      .filter(f => f.endsWith('.json'))
      .map(f => {
        const full = path.join(dir, f)
        const stat = fs.statSync(full)
        return { name: f, size: stat.size, mtime: stat.mtimeMs, path: full }
      })
      .sort((a, b) => b.mtime - a.mtime)
  } catch {
    return []
  }
})

// Yedek dosyasının içeriğini oku (listeden seçilenleri geri yüklemek için)
ipcMain.handle('backup:read-file', (_event, filePath: string) => {
  try {
    // Güvenlik: yalnızca backup klasöründeki dosyalar okunabilsin
    const backupDir = getBackupDir()
    const docsDir = getDocumentsBackupDir()
    const resolved = path.resolve(filePath)
    if (!resolved.startsWith(path.resolve(backupDir)) && !resolved.startsWith(path.resolve(docsDir))) {
      return { success: false, error: 'Yetkisiz dosya yolu' }
    }
    const data = fs.readFileSync(resolved, 'utf-8')
    return { success: true, data }
  } catch (e: any) {
    return { success: false, error: e?.message || String(e) }
  }
})

// ============================================================
// YAZDIRMA / PDF KAYDETME
// ============================================================
// html2pdf kütüphanesi Electron'da büyük tablolarda donmaya yol açıyor.
// Bunun yerine Electron'un kendi webContents.printToPDF() API'sini
// kullanarak çok daha hızlı ve güvenilir PDF üretiyoruz.
// ============================================================

interface PrintOptions {
  orientation?: 'portrait' | 'landscape'
  filename?: string
  marginsType?: number
}

ipcMain.handle('printer:save-pdf', async (event, options: PrintOptions = {}) => {
  try {
    const webContents = event.sender
    const browserWindow = BrowserWindow.fromWebContents(webContents)
    if (!browserWindow) {
      return { success: false, error: 'Pencere bulunamadı' }
    }

    const filename = (options.filename || 'cizelge.pdf').replace(/[\\/:*?"<>|]/g, '_')
    const isLandscape = options.orientation === 'landscape'

    // Önce kullanıcıya kaydetme konumunu sor
    const saveResult = await dialog.showSaveDialog(browserWindow, {
      title: 'PDF olarak kaydet',
      defaultPath: path.join(app.getPath('documents'), filename),
      filters: [{ name: 'PDF Dosyası', extensions: ['pdf'] }],
    })

    if (saveResult.canceled || !saveResult.filePath) {
      return { success: false, canceled: true }
    }

    // Electron'un native PDF üretimi (çok daha hızlı ve güvenilir).
    // preferCSSPageSize: true ile renderer tarafında dinamik enjekte edilen
    // @page size kuralı kullanılır — böylece her form için doğru orientation.
    // scale: 0.92 ile form A4'e tam sığacak şekilde hafifçe küçültülür.
    const pdfBuffer = await webContents.printToPDF({
      landscape: isLandscape,
      pageSize: 'A4',
      printBackground: true,
      preferCSSPageSize: true,
      scale: 0.92,
      margins: {
        marginType: 'custom',
        top: 0.3,
        bottom: 0.3,
        left: 0.2,
        right: 0.2,
      },
    } as any)

    fs.writeFileSync(saveResult.filePath, pdfBuffer)

    // Dosyayı varsayılan PDF uygulamasında aç
    shell.openPath(saveResult.filePath).catch(() => {})

    return { success: true, path: saveResult.filePath }
  } catch (e: any) {
    console.error('PDF kaydetme hatası:', e)
    return { success: false, error: e?.message || String(e) }
  }
})

// Yazıcıdan yazdırma (sistem yazdırma diyaloğunu açar)
ipcMain.handle('printer:print', async (event, options: PrintOptions = {}) => {
  try {
    const webContents = event.sender
    const isLandscape = options.orientation === 'landscape'

    return await new Promise<{ success: boolean; error?: string; canceled?: boolean }>((resolve) => {
      webContents.print(
        {
          silent: false,
          printBackground: true,
          landscape: isLandscape,
          pageSize: 'A4',
          // Form 6 tablo + imza içerdiği için A4'e tam sığması için %85 ölçek.
          // Bu Chromium'un print zoom'u — yazıcı sayfasını küçültür, içerik
          // tek sayfaya sığar.
          scaleFactor: 85,
          margins: {
            marginType: 'default',
          },
        } as any,
        (success, failureReason) => {
          if (success) {
            resolve({ success: true })
          } else if (failureReason === 'cancelled') {
            resolve({ success: false, canceled: true })
          } else {
            resolve({ success: false, error: failureReason || 'Yazdırma başarısız' })
          }
        }
      )
    })
  } catch (e: any) {
    console.error('Yazdırma hatası:', e)
    return { success: false, error: e?.message || String(e) }
  }
})