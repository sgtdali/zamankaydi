import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { supabase } from './supabase'

const GUNLER = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar']
const GUN_KEYS = ['pazartesi', 'sali', 'carsamba', 'persembe', 'cuma', 'cumartesi', 'pazar'] as const
const SATIR_SAYISI = 10
const ONE_DAY_MS = 24 * 60 * 60 * 1000

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Row = Record<string, any>

// ─── Stil yardımcıları ────────────────────────────────────────────────────────

const THIN: ExcelJS.BorderStyle = 'thin'

const BORDER_ALL: Partial<ExcelJS.Borders> = {
  top: { style: THIN },
  bottom: { style: THIN },
  left: { style: THIN },
  right: { style: THIN },
}

function applyStyle(
  cell: ExcelJS.Cell,
  opts: {
    bold?: boolean
    italic?: boolean
    sz?: number
    color?: string        // hex, ör: '1D4ED8'
    bgColor?: string      // hex, ör: 'DBEAFE'
    hAlign?: ExcelJS.Alignment['horizontal']
    vAlign?: ExcelJS.Alignment['vertical']
    wrap?: boolean
    border?: boolean
    numFmt?: string
  }
) {
  cell.font = {
    name: 'Calibri',
    size: opts.sz ?? 10,
    bold: opts.bold ?? false,
    italic: opts.italic ?? false,
    ...(opts.color ? { color: { argb: 'FF' + opts.color } } : {}),
  }
  cell.alignment = {
    horizontal: opts.hAlign ?? 'left',
    vertical: opts.vAlign ?? 'middle',
    wrapText: opts.wrap ?? false,
  }
  if (opts.bgColor) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF' + opts.bgColor },
    }
  }
  if (opts.border !== false) {
    cell.border = BORDER_ALL
  }
  if (opts.numFmt) cell.numFmt = opts.numFmt
}

// ─── Workbook oluşturucu ──────────────────────────────────────────────────────

async function buildWorkbook(ts: Row, rows: Row[]): Promise<ExcelJS.Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'ZamanKaydi'
  const ws = wb.addWorksheet('Zaman Kaydı')

  // Sütun genişlikleri (11 sütun)
  ws.columns = [
    { width: 20 }, // A  İş Tipi
    { width: 6 }, // B  KOD
    { width: 13 }, // C  Çalışılan Makine Kodu
    { width: 11 }, // D  Pazartesi
    { width: 8 }, // E  Salı
    { width: 10 }, // F  Çarşamba
    { width: 10 }, // G  Perşembe
    { width: 8 }, // H  Cuma
    { width: 11 }, // I  Cumartesi
    { width: 8 }, // J  Pazar
    { width: 22 }, // K  NOTLAR
  ]

  // ── Satır 1: Çalışan Adı | Masraf Yeri | Hafta No ────────────────────────
  const r1 = ws.getRow(1)
  r1.height = 20

  ws.mergeCells('A1:B1')
  ws.mergeCells('C1:E1')
  ws.mergeCells('F1:G1')
  ws.mergeCells('H1:I1')

  const a1 = ws.getCell('A1'); a1.value = 'Çalışan Adı:'
  applyStyle(a1, { bold: true })

  const c1 = ws.getCell('C1'); c1.value = String(ts.calisan_adi ?? '')
  applyStyle(c1, {})

  const f1 = ws.getCell('F1'); f1.value = 'Masraf Yeri:'
  applyStyle(f1, { bold: true })

  const h1 = ws.getCell('H1'); h1.value = String(ts.masraf_yeri ?? '')
  applyStyle(h1, {})

  const j1 = ws.getCell('J1'); j1.value = 'Hafta No:'
  applyStyle(j1, { bold: true })

  const k1 = ws.getCell('K1'); k1.value = ts.hafta_no ? Number(ts.hafta_no) : ''
  applyStyle(k1, { bold: true, color: '1D4ED8', hAlign: 'center' })

  // ── Satır 2: Çalışan No | Masraf Yeri Kodu | Tarih ───────────────────────
  const r2 = ws.getRow(2)
  r2.height = 20

  ws.mergeCells('A2:B2')
  ws.mergeCells('C2:E2')
  ws.mergeCells('F2:G2')
  ws.mergeCells('H2:I2')

  const a2 = ws.getCell('A2'); a2.value = 'Çalışan No:'
  applyStyle(a2, { bold: true })

  const c2 = ws.getCell('C2'); c2.value = String(ts.calisan_no ?? '')
  applyStyle(c2, {})

  const f2 = ws.getCell('F2'); f2.value = 'Masraf Yeri Kodu:'
  applyStyle(f2, { bold: true })

  const h2 = ws.getCell('H2'); h2.value = String(ts.masraf_yeri_kodu ?? '')
  applyStyle(h2, {})

  const j2 = ws.getCell('J2'); j2.value = 'Tarih:'
  applyStyle(j2, { bold: true })

  const k2 = ws.getCell('K2'); k2.value = ts.tarih ? String(ts.tarih).split('T')[0] : ''
  applyStyle(k2, {})

  // ── Satır 3: Tablo başlık üst ────────────────────────────────────────────
  ws.getRow(3).height = 30
  ws.mergeCells('A3:A4')
  ws.mergeCells('B3:B4')
  ws.mergeCells('C3:C4')
  ws.mergeCells('D3:J3')
  ws.mergeCells('K3:K4')

  const thStyle = { bold: true, bgColor: 'DBEAFE', hAlign: 'center' as const, vAlign: 'middle' as const, wrap: true }

  const a3 = ws.getCell('A3'); a3.value = 'İş Tipi'; applyStyle(a3, thStyle)
  const b3 = ws.getCell('B3'); b3.value = 'KOD'; applyStyle(b3, thStyle)
  const c3 = ws.getCell('C3'); c3.value = 'Çalışılan\nMakine Kodu'; applyStyle(c3, thStyle)
  const d3 = ws.getCell('D3'); d3.value = 'Çalışılan Süre (saat)'; applyStyle(d3, thStyle)
  const k3 = ws.getCell('K3'); k3.value = 'NOTLAR'; applyStyle(k3, thStyle)

  // ── Satır 4: Gün başlıkları ───────────────────────────────────────────────
  ws.getRow(4).height = 20
  // merged hücre dolguları (border görünsün diye)
  applyStyle(ws.getCell('A4'), thStyle)
  applyStyle(ws.getCell('B4'), thStyle)
  applyStyle(ws.getCell('C4'), thStyle)
  applyStyle(ws.getCell('K4'), thStyle)

  const gunCols = ['D', 'E', 'F', 'G', 'H', 'I', 'J']
  GUNLER.forEach((gun, i) => {
    const cell = ws.getCell(`${gunCols[i]}4`)
    cell.value = gun
    applyStyle(cell, thStyle)
  })

  // ── Veri satırları (5..14) ────────────────────────────────────────────────
  const sortedRows = [...rows].sort((a, b) => Number(a.sira_no) - Number(b.sira_no))

  for (let i = 0; i < SATIR_SAYISI; i++) {
    const r = sortedRows[i]
    const rowNum = 5 + i
    const exRow = ws.getRow(rowNum)
    exRow.height = 17

    const bgColor = i % 2 === 1 ? 'F9FAFB' : undefined

    const cellA = ws.getCell(`A${rowNum}`); cellA.value = r ? String(r.is_tipi ?? '') : ''
    applyStyle(cellA, { bgColor })

    const cellB = ws.getCell(`B${rowNum}`); cellB.value = r ? String(r.kod ?? '') : ''
    applyStyle(cellB, { hAlign: 'center', bgColor })

    const cellC = ws.getCell(`C${rowNum}`); cellC.value = r ? String(r.makine_kodu ?? '') : ''
    applyStyle(cellC, { hAlign: 'center', bgColor })

    GUN_KEYS.forEach((g, gi) => {
      const val = r ? Number(r[g] ?? 0) : 0
      const c = ws.getCell(`${gunCols[gi]}${rowNum}`)
      if (val > 0) {
        c.value = val
        // Tam sayıysa virgülsüz, ondalıklıysa 0.## formatı
        c.numFmt = Number.isInteger(val) ? '0' : '0.##'
      } else {
        c.value = null
      }
      applyStyle(c, { hAlign: 'center', bgColor })
    })

    const cellK = ws.getCell(`K${rowNum}`); cellK.value = r ? String(r.notlar ?? '') : ''
    applyStyle(cellK, { bgColor })
  }

  // ── TOPLAM SÜRE (satır 15) ────────────────────────────────────────────────
  const totRow = 5 + SATIR_SAYISI
  ws.getRow(totRow).height = 20
  ws.mergeCells(`A${totRow}:C${totRow}`)

  const totLabel = ws.getCell(`A${totRow}`)
  totLabel.value = 'TOPLAM SÜRE'
  applyStyle(totLabel, { bold: true, color: '1D4ED8', bgColor: 'DBEAFE', hAlign: 'center' })

  applyStyle(ws.getCell(`B${totRow}`), { bold: true, bgColor: 'DBEAFE' })
  applyStyle(ws.getCell(`C${totRow}`), { bold: true, bgColor: 'DBEAFE' })

  GUN_KEYS.forEach((g, gi) => {
    const toplam = rows.reduce((acc, row) => acc + Number(row[g] ?? 0), 0)
    const c = ws.getCell(`${gunCols[gi]}${totRow}`)
    c.value = toplam
    c.numFmt = '0.00'
    applyStyle(c, { bold: true, color: '1D4ED8', bgColor: 'DBEAFE', hAlign: 'center' })
  })

  applyStyle(ws.getCell(`K${totRow}`), { bold: true, bgColor: 'DBEAFE' })

  return wb.xlsx.writeBuffer()
}

// ─── Tüm personel listesi ─────────────────────────────────────────────────────

const TUM_PERSONEL = [
  'ABDULSAMEt ÖZTÜRK', 'ABDURRAHMAN ALDEMİR', 'ADEM SELİM', 'AHMET KESKİN',
  'AHMET UYGUR', 'AYHAN ŞAHAN', 'AYKUT ARSLANALP', 'BARAN KORKMAZ',
  'BARIŞ DURAN', 'BERK BABACAN', 'CABİR KOÇ', 'CENGİZ ÜSTÜN',
  'CİHAT BIÇKI', 'ÇAĞRI CAN ÇOLAK', 'DOĞAN EROL', 'ERKAN KÜLAHLΙ',
  'FATİH UZUNAL', 'FERHAT ÇOBAN', 'HALİL İBRAHİM DEMİREL', 'HALİT ÇELİK',
  'İBRAHİM KARA', 'KADİR YÜKSELEN', 'KEMAL ÜSTÜN', 'MEHMET CAN AKAR',
  'METEHAN ARGUT', 'MUHSİN UYSAL', 'MUSTAFA ŞAHİN', 'MUSTAFA YILDIZ',
  'MÜCAHİT TOPTAŞ', 'OKAN CEYHAN', 'ONUR AKCI', 'ÖZGÜR KALAYCI',
  'RESUL KEKLİK', 'SEDAT KARAKAYA', 'SERHAT FATİH KALYONCU', 'SUAT TUNÇ',
  'ŞENEL ÇELİK', 'TANER ÇELİK', 'UĞUR BOZYURT', 'ULAŞ ÇELİK',
  'VOLKAN MADEN', 'YASİN DURSUN', 'YİĞİT ALİ ÜNAL', 'ZAFER ÇAĞMAN',
]

// ─── Toplu özet Excel ─────────────────────────────────────────────────────────

async function buildSummaryWorkbook(haftaNo: number, sheets: Row[]): Promise<ExcelJS.Buffer> {
  // kişi adı → gün toplamları map'i
  const dataMap: Record<string, Record<string, number>> = {}
  for (const ts of sheets) {
    const rows = (ts.zamankay_timesheet_rows ?? []) as Row[]
    const gunToplam: Record<string, number> = {}
    GUN_KEYS.forEach(g => {
      gunToplam[g] = rows.reduce((acc, r) => acc + Number(r[g] ?? 0), 0)
    })
    dataMap[String(ts.calisan_adi)] = gunToplam
  }

  const wb = new ExcelJS.Workbook()
  wb.creator = 'ZamanKaydi'
  const ws = wb.addWorksheet('Özet')

  // Sütun genişlikleri: Personel + 7 gün + Toplam
  ws.columns = [
    { width: 26 }, // Personel adı
    ...GUNLER.map(() => ({ width: 12 })),
    { width: 10 }, // Haftalık toplam
  ]

  // ── Başlık satırı ────────────────────────────────────────────────────────
  ws.getRow(1).height = 22
  const baslikStyle = { bold: true, bgColor: 'DBEAFE', hAlign: 'center' as const, vAlign: 'middle' as const }

  const h0 = ws.getCell('A1'); h0.value = `Hafta ${haftaNo} — Personel Özeti`
  applyStyle(h0, { ...baslikStyle, hAlign: 'left' })
  ws.mergeCells(`A1:${String.fromCharCode(65 + 1 + GUNLER.length)}1`)

  // ── Kolon başlıkları ─────────────────────────────────────────────────────
  ws.getRow(2).height = 20
  const c0 = ws.getCell('A2'); c0.value = 'Personel Adı'
  applyStyle(c0, baslikStyle)

  GUNLER.forEach((gun, i) => {
    const c = ws.getCell(2, 2 + i); c.value = gun
    applyStyle(c, baslikStyle)
  })

  const cToplam = ws.getCell(2, 2 + GUNLER.length); cToplam.value = 'Haftalık\nToplam'
  applyStyle(cToplam, { ...baslikStyle, wrap: true })

  // ── Personel satırları ───────────────────────────────────────────────────
  TUM_PERSONEL.forEach((ad, idx) => {
    const rowNum = 3 + idx
    const exRow = ws.getRow(rowNum)
    exRow.height = 17

    const bgColor = idx % 2 === 1 ? 'F9FAFB' : undefined
    const gunToplam = dataMap[ad]

    // Personel adı
    const nameCell = ws.getCell(rowNum, 1)
    nameCell.value = ad
    applyStyle(nameCell, { bgColor })

    // Günler
    let haftaToplam = 0
    GUN_KEYS.forEach((g, gi) => {
      const val = gunToplam ? (gunToplam[g] ?? 0) : 0
      haftaToplam += val
      const c = ws.getCell(rowNum, 2 + gi)
      if (val > 0) {
        c.value = val
        c.numFmt = Number.isInteger(val) ? '0' : '0.##'
        applyStyle(c, { hAlign: 'center', bgColor })
      } else if (gunToplam) {
        // Kayıt var ama o gün 0 → boş bırak
        applyStyle(c, { hAlign: 'center', bgColor })
      } else {
        // Hiç kayıt yok → gri arka plan
        applyStyle(c, { hAlign: 'center', bgColor: bgColor ?? 'F3F4F6' })
      }
    })

    // Haftalık toplam
    const totCell = ws.getCell(rowNum, 2 + GUNLER.length)
    if (gunToplam && haftaToplam > 0) {
      totCell.value = haftaToplam
      totCell.numFmt = Number.isInteger(haftaToplam) ? '0' : '0.##'
      applyStyle(totCell, { bold: true, hAlign: 'center', bgColor })
    } else if (!gunToplam) {
      totCell.value = 'Giriş Yok'
      applyStyle(totCell, { hAlign: 'center', color: 'EF4444', bgColor: bgColor ?? 'F3F4F6' })
    } else {
      applyStyle(totCell, { hAlign: 'center', bgColor })
    }
  })

  // ── Alt toplam satırı ────────────────────────────────────────────────────
  const altRow = 3 + TUM_PERSONEL.length
  ws.getRow(altRow).height = 20

  const altLabel = ws.getCell(altRow, 1)
  altLabel.value = 'GENEL TOPLAM'
  applyStyle(altLabel, { bold: true, bgColor: 'DBEAFE', color: '1D4ED8' })

  let genelToplam = 0
  GUN_KEYS.forEach((g, gi) => {
    const gunelT = Object.values(dataMap).reduce((acc, d) => acc + (d[g] ?? 0), 0)
    genelToplam += gunelT
    const c = ws.getCell(altRow, 2 + gi)
    c.value = gunelT
    c.numFmt = '0.00'
    applyStyle(c, { bold: true, hAlign: 'center', bgColor: 'DBEAFE', color: '1D4ED8' })
  })

  const genelTotCell = ws.getCell(altRow, 2 + GUNLER.length)
  genelTotCell.value = genelToplam
  genelTotCell.numFmt = '0.00'
  applyStyle(genelTotCell, { bold: true, hAlign: 'center', bgColor: 'DBEAFE', color: '1D4ED8' })

  return wb.xlsx.writeBuffer()
}

function getYearFromTimesheet(ts: Row) {
  const rawDate = String(ts.tarih ?? ts.created_at ?? '')
  const yearText = rawDate.slice(0, 4)
  const year = Number(yearText)
  return Number.isFinite(year) && year > 1900 ? year : new Date().getFullYear()
}

function getIsoWeekDate(year: number, weekNo: number, dayIndex: number) {
  const jan4 = new Date(Date.UTC(year, 0, 4))
  const jan4Day = jan4.getUTCDay() || 7
  const firstMonday = jan4.getTime() - (jan4Day - 1) * ONE_DAY_MS
  return new Date(firstMonday + ((weekNo - 1) * 7 + dayIndex) * ONE_DAY_MS)
}

function formatReportDate(date: Date) {
  return `${date.getUTCDate()}.${date.getUTCMonth() + 1}.${date.getUTCFullYear()}`
}

function getLatestSheetsByPersonWeekYear(sheets: Row[]) {
  const latest = new Map<string, Row>()

  for (const sheet of sheets) {
    const person = String(sheet.calisan_adi ?? '').trim()
    const weekNo = Number(sheet.hafta_no ?? 0)
    if (!person || !weekNo) continue

    const key = `${person}|${getYearFromTimesheet(sheet)}|${weekNo}`
    const current = latest.get(key)
    const currentCreatedAt = current ? new Date(String(current.created_at ?? '')).getTime() : 0
    const sheetCreatedAt = new Date(String(sheet.created_at ?? '')).getTime()

    if (!current || sheetCreatedAt >= currentCreatedAt) {
      latest.set(key, sheet)
    }
  }

  return Array.from(latest.values())
}

async function buildDetailedAllDataWorkbook(sheets: Row[]): Promise<ExcelJS.Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'ZamanKaydi'
  const ws = wb.addWorksheet('Tum Veri')

  ws.columns = [
    { width: 24 }, // KISI
    { width: 10 }, // SICIL NO
    { width: 8 },  // HAFTA
    { width: 12 }, // TARIH
    { width: 16 }, // DEPARTMAN
    { width: 15 }, // MAKINA KODU
    { width: 10 }, // IS KODU
    { width: 16 }, // IS TIPI
    { width: 12 }, // SURE
    { width: 18 }, // NOTLAR
    { width: 24 }, // ACIKLAMA
  ]

  const headers = ['KİŞİ', 'SİCİL NO', 'HAFTA', 'TARİH', 'LOKASYON', 'MAKİNA KODU', 'İŞ KODU', 'İŞ TİPİ', 'SÜRE (SAAT)', 'NOTLAR', 'AÇIKLAMA']
  const headerRow = ws.getRow(1)
  headerRow.height = 20
  headers.forEach((header, idx) => {
    const cell = ws.getCell(1, idx + 1)
    cell.value = header
    applyStyle(cell, { bold: true, bgColor: 'FACC15', hAlign: 'center', vAlign: 'middle' })
  })

  const detailRows: Row[] = []
  for (const ts of sheets) {
    const rows = (ts.zamankay_timesheet_rows ?? []) as Row[]
    const weekNo = Number(ts.hafta_no ?? 0)
    if (!weekNo) continue

    const year = getYearFromTimesheet(ts)
    for (const row of rows) {
      GUN_KEYS.forEach((gun, dayIndex) => {
        const sure = Number(row[gun] ?? 0)
        if (sure <= 0) return

        const date = getIsoWeekDate(year, weekNo, dayIndex)
        detailRows.push({
          kisi: String(ts.calisan_adi ?? ''),
          sicilNo: String(ts.calisan_no ?? ''),
          hafta: weekNo,
          tarih: formatReportDate(date),
          tarihSort: date.getTime(),
          departman: String(ts.masraf_yeri ?? ''),
          makineKodu: String(row.makine_kodu ?? ''),
          isKodu: String(row.kod ?? ''),
          isTipi: String(row.is_tipi ?? ''),
          sure,
          notlar: String(row.notlar ?? ''),
          aciklama: '',
        })
      })
    }
  }

  detailRows.sort((a, b) =>
    Number(a.tarihSort) - Number(b.tarihSort) ||
    String(a.kisi).localeCompare(String(b.kisi), 'tr') ||
    String(a.makineKodu).localeCompare(String(b.makineKodu), 'tr')
  )

  detailRows.forEach((row, idx) => {
    const rowNum = idx + 2
    const bgColor = idx % 2 === 1 ? 'F9FAFB' : undefined
    const values = [
      row.kisi,
      row.sicilNo,
      row.hafta,
      row.tarih,
      row.departman,
      row.makineKodu,
      row.isKodu,
      row.isTipi,
      row.sure,
      row.notlar,
      row.aciklama,
    ]

    values.forEach((value, colIdx) => {
      const cell = ws.getCell(rowNum, colIdx + 1)
      cell.value = value
      if (colIdx === 8 && typeof value === 'number') {
        cell.numFmt = Number.isInteger(value) ? '0' : '0.##'
      }
      applyStyle(cell, {
        bgColor,
        hAlign: [1, 2, 8].includes(colIdx) ? 'center' : 'left',
      })
    })
  })

  ws.autoFilter = {
    from: 'A1',
    to: 'K1',
  }
  ws.views = [{ state: 'frozen', ySplit: 1 }]

  return wb.xlsx.writeBuffer()
}

// ─── Supabase fetch ───────────────────────────────────────────────────────────

async function fetchTimesheetData(haftaNo: number, calisanAdi?: string) {
  let q = supabase
    .from('zamankay_timesheets')
    .select('*, zamankay_timesheet_rows(*)')
    .eq('hafta_no', haftaNo)
    .order('created_at', { ascending: false })

  if (calisanAdi) q = q.eq('calisan_adi', calisanAdi)

  const { data, error } = await q
  if (error) throw new Error(error.message)
  return data ?? []
}

async function fetchAllTimesheetData() {
  const pageSize = 1000
  let from = 0
  const allRows: Row[] = []

  while (true) {
    const { data, error } = await supabase
      .from('zamankay_timesheets')
      .select('*, zamankay_timesheet_rows(*)')
      .order('tarih', { ascending: true })
      .order('created_at', { ascending: true })
      .range(from, from + pageSize - 1)

    if (error) throw new Error(error.message)
    allRows.push(...((data ?? []) as Row[]))
    if (!data || data.length < pageSize) break
    from += pageSize
  }

  return allRows
}

function safeName(name: string) {
  return name.replace(/[\\/:*?"<>|]/g, '_')
}

function downloadBuffer(buf: ExcelJS.Buffer, filename: string) {
  const blob = new Blob([buf as ArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

// ─── Public API ───────────────────────────────────────────────────────────────

export async function exportOne(haftaNo: number, calisanAdi: string) {
  const sheets = await fetchTimesheetData(haftaNo, calisanAdi)
  if (!sheets.length) throw new Error(`Hafta ${haftaNo} için ${calisanAdi} kaydı bulunamadı.`)

  const ts = sheets[0] as Row
  const rows = (ts.zamankay_timesheet_rows ?? []) as Row[]
  const buf = await buildWorkbook(ts, rows)
  downloadBuffer(buf, `ZamanKaydi_Hafta${haftaNo}_${safeName(calisanAdi)}.xlsx`)
}

export async function exportAll(haftaNo: number) {
  const sheets = await fetchTimesheetData(haftaNo)
  if (!sheets.length) throw new Error(`Hafta ${haftaNo} için hiç kayıt bulunamadı.`)

  const zip = new JSZip()
  const folder = zip.folder(`ZamanKaydi_Hafta${haftaNo}`)!

  for (const ts of sheets as Row[]) {
    const rows = (ts.zamankay_timesheet_rows ?? []) as Row[]
    const buf = await buildWorkbook(ts, rows)
    folder.file(`${safeName(String(ts.calisan_adi))}.xlsx`, buf as ArrayBuffer)
  }

  // Toplu özet Excel
  const summaryBuf = await buildSummaryWorkbook(haftaNo, sheets as Row[])
  folder.file(`_OZET_Hafta${haftaNo}.xlsx`, summaryBuf as ArrayBuffer)

  const zipBlob = await zip.generateAsync({ type: 'blob' })
  const url = URL.createObjectURL(zipBlob)
  const a = document.createElement('a')
  a.href = url
  a.download = `ZamanKaydi_Hafta${haftaNo}_TumPersonel.zip`
  a.click()
  URL.revokeObjectURL(url)
}

export async function exportDetailedAllData() {
  const sheets = await fetchAllTimesheetData()
  if (!sheets.length) throw new Error('Dışa aktarılacak kayıt bulunamadı.')

  const buf = await buildDetailedAllDataWorkbook(getLatestSheetsByPersonWeekYear(sheets as Row[]))
  downloadBuffer(buf, 'ZamanKaydi_TumVeri_Detayli.xlsx')
}
