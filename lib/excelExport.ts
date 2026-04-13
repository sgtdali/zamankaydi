import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { supabase } from './supabase'

const GUNLER = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar']
const GUN_KEYS = ['pazartesi', 'sali', 'carsamba', 'persembe', 'cuma', 'cumartesi', 'pazar'] as const
const SATIR_SAYISI = 10

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Row = Record<string, any>

// ─── Stil yardımcıları ────────────────────────────────────────────────────────

const THIN: ExcelJS.BorderStyle = 'thin'

const BORDER_ALL: Partial<ExcelJS.Borders> = {
  top:    { style: THIN },
  bottom: { style: THIN },
  left:   { style: THIN },
  right:  { style: THIN },
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
    vertical:   opts.vAlign ?? 'middle',
    wrapText:   opts.wrap ?? false,
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
    { width: 6  }, // B  KOD
    { width: 13 }, // C  Çalışılan Makine Kodu
    { width: 11 }, // D  Pazartesi
    { width: 8  }, // E  Salı
    { width: 10 }, // F  Çarşamba
    { width: 10 }, // G  Perşembe
    { width: 8  }, // H  Cuma
    { width: 11 }, // I  Cumartesi
    { width: 8  }, // J  Pazar
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

  const a3 = ws.getCell('A3'); a3.value = 'İş Tipi';                    applyStyle(a3, thStyle)
  const b3 = ws.getCell('B3'); b3.value = 'KOD';                        applyStyle(b3, thStyle)
  const c3 = ws.getCell('C3'); c3.value = 'Çalışılan\nMakine Kodu';     applyStyle(c3, thStyle)
  const d3 = ws.getCell('D3'); d3.value = 'Çalışılan Süre (saat)';      applyStyle(d3, thStyle)
  const k3 = ws.getCell('K3'); k3.value = 'NOTLAR';                     applyStyle(k3, thStyle)

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

  const zipBlob = await zip.generateAsync({ type: 'blob' })
  const url = URL.createObjectURL(zipBlob)
  const a = document.createElement('a')
  a.href = url
  a.download = `ZamanKaydi_Hafta${haftaNo}_TumPersonel.zip`
  a.click()
  URL.revokeObjectURL(url)
}
