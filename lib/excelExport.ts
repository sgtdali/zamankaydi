import * as XLSX from 'xlsx'
import JSZip from 'jszip'
import { supabase } from './supabase'

const GUNLER = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar']
const GUN_KEYS = ['pazartesi', 'sali', 'carsamba', 'persembe', 'cuma', 'cumartesi', 'pazar'] as const
const SATIR_SAYISI = 10

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Row = Record<string, any>
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type CellStyle = Record<string, any>

// ─── Stil sabitleri ────────────────────────────────────────────────────────────

const BORDER_THIN = {
  top:    { style: 'thin' },
  bottom: { style: 'thin' },
  left:   { style: 'thin' },
  right:  { style: 'thin' },
}

/** Başlık label hücresi: bold, sola yaslı, border */
const S_LABEL: CellStyle = {
  font: { bold: true, sz: 10, name: 'Calibri' },
  alignment: { horizontal: 'left', vertical: 'center' },
  border: BORDER_THIN,
}

/** Başlık value hücresi: normal, sola yaslı, border */
const S_VALUE: CellStyle = {
  font: { sz: 10, name: 'Calibri' },
  alignment: { horizontal: 'left', vertical: 'center' },
  border: BORDER_THIN,
}

/** Hafta No değeri: mavi, bold */
const S_HAFTA: CellStyle = {
  font: { bold: true, sz: 10, color: { rgb: '1D4ED8' }, name: 'Calibri' },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: BORDER_THIN,
}

/** Tablo başlık hücresi: bold, ortalı, mavi arka plan, border, word-wrap */
const S_TH: CellStyle = {
  font: { bold: true, sz: 10, name: 'Calibri' },
  fill: { fgColor: { rgb: 'DBEAFE' }, patternType: 'solid' },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: BORDER_THIN,
}

/** Veri hücresi: normal, border */
const S_TD: CellStyle = {
  font: { sz: 10, name: 'Calibri' },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: BORDER_THIN,
}

/** Veri hücresi sol yaslı (İş Tipi, Notlar) */
const S_TD_LEFT: CellStyle = {
  font: { sz: 10, name: 'Calibri' },
  alignment: { horizontal: 'left', vertical: 'center' },
  border: BORDER_THIN,
}

/** Toplam satırı: bold, mavi, border */
const S_TOTAL: CellStyle = {
  font: { bold: true, sz: 10, color: { rgb: '1D4ED8' }, name: 'Calibri' },
  fill: { fgColor: { rgb: 'DBEAFE' }, patternType: 'solid' },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: BORDER_THIN,
}

const S_TOTAL_LABEL: CellStyle = {
  font: { bold: true, sz: 10, color: { rgb: '1D4ED8' }, name: 'Calibri' },
  fill: { fgColor: { rgb: 'DBEAFE' }, patternType: 'solid' },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: BORDER_THIN,
}

// ─── Yardımcılar ───────────────────────────────────────────────────────────────

function cell(r: number, c: number): string {
  return XLSX.utils.encode_cell({ r, c })
}

function sc(ws: XLSX.WorkSheet, r: number, c: number, v: XLSX.CellObject) {
  ws[cell(r, c)] = v
}

function merge(merges: XLSX.Range[], r1: number, c1: number, r2: number, c2: number) {
  if (r1 === r2 && c1 === c2) return
  merges.push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } })
}

// Boş border'lı hücre (dolgu için)
function emptyCell(style: CellStyle = S_TD): XLSX.CellObject {
  return { t: 's', v: '', s: style }
}

// ─── Workbook oluşturucu ───────────────────────────────────────────────────────

function buildWorkbook(ts: Row, rows: Row[]): XLSX.WorkBook {
  const wb = XLSX.utils.book_new()
  const ws: XLSX.WorkSheet = {}
  const merges: XLSX.Range[] = []

  // Sütun genişlikleri (11 sütun: 0-10)
  ws['!cols'] = [
    { wch: 20 }, // 0  İş Tipi
    { wch: 6  }, // 1  KOD
    { wch: 12 }, // 2  Çalışılan Makine Kodu
    { wch: 11 }, // 3  Pazartesi
    { wch: 8  }, // 4  Salı
    { wch: 10 }, // 5  Çarşamba
    { wch: 10 }, // 6  Perşembe
    { wch: 8  }, // 7  Cuma
    { wch: 11 }, // 8  Cumartesi
    { wch: 8  }, // 9  Pazar
    { wch: 22 }, // 10 NOTLAR
  ]

  // Satır yükseklikleri
  ws['!rows'] = [
    { hpx: 22 }, // 0 header row 1
    { hpx: 22 }, // 1 header row 2
    { hpx: 30 }, // 2 tablo başlık üst
    { hpx: 22 }, // 3 tablo başlık alt (günler)
    ...Array(SATIR_SAYISI).fill({ hpx: 18 }),
    { hpx: 22 }, // toplam
  ]

  // ── Satır 0: Çalışan Adı | Masraf Yeri | Hafta No ──────────────────────────
  //  A:B = label, C:E = value | F:G = label, H:I = value | J = label, K = value
  sc(ws, 0, 0, { t: 's', v: 'Çalışan Adı:', s: S_LABEL })
  merge(merges, 0, 0, 0, 1)
  sc(ws, 0, 1, emptyCell(S_LABEL))
  sc(ws, 0, 2, { t: 's', v: String(ts.calisan_adi ?? ''), s: S_VALUE })
  merge(merges, 0, 2, 0, 4)
  sc(ws, 0, 3, emptyCell(S_VALUE)); sc(ws, 0, 4, emptyCell(S_VALUE))
  sc(ws, 0, 5, { t: 's', v: 'Masraf Yeri:', s: S_LABEL })
  merge(merges, 0, 5, 0, 6)
  sc(ws, 0, 6, emptyCell(S_LABEL))
  sc(ws, 0, 7, { t: 's', v: String(ts.masraf_yeri ?? ''), s: S_VALUE })
  merge(merges, 0, 7, 0, 8)
  sc(ws, 0, 8, emptyCell(S_VALUE))
  sc(ws, 0, 9, { t: 's', v: 'Hafta No:', s: S_LABEL })
  sc(ws, 0, 10, { t: 'n', v: Number(ts.hafta_no ?? 0), s: S_HAFTA })

  // ── Satır 1: Çalışan No | Masraf Yeri Kodu | Tarih ─────────────────────────
  sc(ws, 1, 0, { t: 's', v: 'Çalışan No:', s: S_LABEL })
  merge(merges, 1, 0, 1, 1)
  sc(ws, 1, 1, emptyCell(S_LABEL))
  sc(ws, 1, 2, { t: 's', v: String(ts.calisan_no ?? ''), s: S_VALUE })
  merge(merges, 1, 2, 1, 4)
  sc(ws, 1, 3, emptyCell(S_VALUE)); sc(ws, 1, 4, emptyCell(S_VALUE))
  sc(ws, 1, 5, { t: 's', v: 'Masraf Yeri Kodu:', s: S_LABEL })
  merge(merges, 1, 5, 1, 6)
  sc(ws, 1, 6, emptyCell(S_LABEL))
  sc(ws, 1, 7, { t: 's', v: String(ts.masraf_yeri_kodu ?? ''), s: S_VALUE })
  merge(merges, 1, 7, 1, 8)
  sc(ws, 1, 8, emptyCell(S_VALUE))
  sc(ws, 1, 9, { t: 's', v: 'Tarih:', s: S_LABEL })
  sc(ws, 1, 10, { t: 's', v: ts.tarih ? String(ts.tarih).split('T')[0] : '', s: S_VALUE })

  // ── Satır 2: Tablo başlık üst ───────────────────────────────────────────────
  sc(ws, 2, 0, { t: 's', v: 'İş Tipi', s: S_TH })
  merge(merges, 2, 0, 3, 0)
  sc(ws, 2, 1, { t: 's', v: 'KOD', s: S_TH })
  merge(merges, 2, 1, 3, 1)
  sc(ws, 2, 2, { t: 's', v: 'Çalışılan\nMakine Kodu', s: S_TH })
  merge(merges, 2, 2, 3, 2)
  sc(ws, 2, 3, { t: 's', v: 'Çalışılan Süre (saat)', s: S_TH })
  merge(merges, 2, 3, 2, 9)
  for (let c = 4; c <= 9; c++) sc(ws, 2, c, emptyCell(S_TH))
  sc(ws, 2, 10, { t: 's', v: 'NOTLAR', s: S_TH })
  merge(merges, 2, 10, 3, 10)

  // ── Satır 3: Gün başlıkları ─────────────────────────────────────────────────
  GUNLER.forEach((g, i) => sc(ws, 3, 3 + i, { t: 's', v: g, s: S_TH }))
  // Rowspan dolgu
  sc(ws, 3, 0, emptyCell(S_TH))
  sc(ws, 3, 1, emptyCell(S_TH))
  sc(ws, 3, 2, emptyCell(S_TH))
  sc(ws, 3, 10, emptyCell(S_TH))

  // ── Veri satırları ──────────────────────────────────────────────────────────
  const sortedRows = [...rows].sort((a, b) => Number(a.sira_no) - Number(b.sira_no))
  for (let i = 0; i < SATIR_SAYISI; i++) {
    const r = sortedRows[i]
    const rIdx = 4 + i
    const bg = i % 2 === 0
      ? S_TD_LEFT
      : { ...S_TD_LEFT, fill: { fgColor: { rgb: 'F9FAFB' }, patternType: 'solid' } } as CellStyle

    sc(ws, rIdx, 0, { t: 's', v: r ? String(r.is_tipi ?? '')    : '', s: bg })
    sc(ws, rIdx, 1, { t: 's', v: r ? String(r.kod ?? '')        : '', s: { ...S_TD } })
    sc(ws, rIdx, 2, { t: 's', v: r ? String(r.makine_kodu ?? '') : '', s: { ...S_TD } })

    GUN_KEYS.forEach((g, gi) => {
      const val = r ? Number(r[g] ?? 0) : 0
      if (val > 0) {
        sc(ws, rIdx, 3 + gi, { t: 'n', v: val, z: '0.##', s: S_TD })
      } else {
        sc(ws, rIdx, 3 + gi, emptyCell(S_TD))
      }
    })

    sc(ws, rIdx, 10, { t: 's', v: r ? String(r.notlar ?? '') : '', s: S_TD_LEFT })
  }

  // ── TOPLAM SÜRE ─────────────────────────────────────────────────────────────
  const totRow = 4 + SATIR_SAYISI
  sc(ws, totRow, 0, { t: 's', v: 'TOPLAM SÜRE', s: S_TOTAL_LABEL })
  merge(merges, totRow, 0, totRow, 2)
  sc(ws, totRow, 1, emptyCell(S_TOTAL_LABEL))
  sc(ws, totRow, 2, emptyCell(S_TOTAL_LABEL))

  GUN_KEYS.forEach((g, gi) => {
    const toplam = rows.reduce((acc, row) => acc + Number(row[g] ?? 0), 0)
    sc(ws, totRow, 3 + gi, { t: 'n', v: toplam, z: '0.00', s: S_TOTAL })
  })
  sc(ws, totRow, 10, emptyCell(S_TOTAL))

  ws['!merges'] = merges
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: totRow, c: 10 } })

  XLSX.utils.book_append_sheet(wb, ws, 'Zaman Kaydı')
  return wb
}

// ─── Buffer / Download ─────────────────────────────────────────────────────────

function wbToBuffer(wb: XLSX.WorkBook): ArrayBuffer {
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true }) as ArrayBuffer
}

function download(buf: ArrayBuffer, filename: string) {
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

function safeName(name: string) {
  return name.replace(/[\\/:*?"<>|]/g, '_')
}

// ─── Supabase fetch ────────────────────────────────────────────────────────────

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

// ─── Public API ────────────────────────────────────────────────────────────────

export async function exportOne(haftaNo: number, calisanAdi: string) {
  const sheets = await fetchTimesheetData(haftaNo, calisanAdi)
  if (!sheets.length) throw new Error(`Hafta ${haftaNo} için ${calisanAdi} kaydı bulunamadı.`)

  const ts = sheets[0] as Row
  const rows = (ts.zamankay_timesheet_rows ?? []) as Row[]
  const wb = buildWorkbook(ts, rows)
  download(wbToBuffer(wb), `ZamanKaydi_Hafta${haftaNo}_${safeName(calisanAdi)}.xlsx`)
}

export async function exportAll(haftaNo: number) {
  const sheets = await fetchTimesheetData(haftaNo)
  if (!sheets.length) throw new Error(`Hafta ${haftaNo} için hiç kayıt bulunamadı.`)

  const zip = new JSZip()
  const folder = zip.folder(`ZamanKaydi_Hafta${haftaNo}`)!

  for (const ts of sheets as Row[]) {
    const rows = (ts.zamankay_timesheet_rows ?? []) as Row[]
    const wb = buildWorkbook(ts, rows)
    folder.file(`${safeName(String(ts.calisan_adi))}.xlsx`, wbToBuffer(wb))
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' })
  const url = URL.createObjectURL(zipBlob)
  const a = document.createElement('a')
  a.href = url
  a.download = `ZamanKaydi_Hafta${haftaNo}_TumPersonel.zip`
  a.click()
  URL.revokeObjectURL(url)
}
