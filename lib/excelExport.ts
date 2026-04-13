import * as XLSX from 'xlsx'
import JSZip from 'jszip'
import { supabase } from './supabase'

const GUNLER = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar']
const GUN_KEYS = ['pazartesi', 'sali', 'carsamba', 'persembe', 'cuma', 'cumartesi', 'pazar'] as const
const SATIR_SAYISI = 10

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Row = Record<string, any>

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

function buildWorkbook(ts: Row, rows: Row[]): XLSX.WorkBook {
  const wb = XLSX.utils.book_new()
  const ws: XLSX.WorkSheet = {}

  // Merge ve stil yardımcıları
  const merges: XLSX.Range[] = []
  const addMerge = (r1: number, c1: number, r2: number, c2: number) =>
    merges.push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } })

  // Sütun genişlikleri
  ws['!cols'] = [
    { wch: 22 }, // İş Tipi
    { wch: 6  }, // KOD
    { wch: 14 }, // Makine Kodu
    { wch: 11 }, // Pazartesi
    { wch: 8  }, // Salı
    { wch: 10 }, // Çarşamba
    { wch: 10 }, // Perşembe
    { wch: 8  }, // Cuma
    { wch: 11 }, // Cumartesi
    { wch: 8  }, // Pazar
    { wch: 20 }, // Notlar
  ]

  const C = (r: number, c: number) => XLSX.utils.encode_cell({ r, c })

  const hStyle = {
    font: { bold: true, sz: 10 },
    fill: { fgColor: { rgb: 'DBEAFE' }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top:    { style: 'thin', color: { rgb: '9CA3AF' } },
      bottom: { style: 'thin', color: { rgb: '9CA3AF' } },
      left:   { style: 'thin', color: { rgb: '9CA3AF' } },
      right:  { style: 'thin', color: { rgb: '9CA3AF' } },
    },
  }
  const labelStyle = {
    font: { bold: true, sz: 10 },
    alignment: { horizontal: 'left', vertical: 'center' },
    border: {
      top:    { style: 'thin', color: { rgb: '9CA3AF' } },
      bottom: { style: 'thin', color: { rgb: '9CA3AF' } },
      left:   { style: 'thin', color: { rgb: '9CA3AF' } },
      right:  { style: 'thin', color: { rgb: '9CA3AF' } },
    },
  }
  const valueStyle = {
    font: { sz: 10 },
    alignment: { horizontal: 'left', vertical: 'center' },
    border: {
      top:    { style: 'thin', color: { rgb: '9CA3AF' } },
      bottom: { style: 'thin', color: { rgb: '9CA3AF' } },
      left:   { style: 'thin', color: { rgb: '9CA3AF' } },
      right:  { style: 'thin', color: { rgb: '9CA3AF' } },
    },
  }
  const numHStyle = {
    font: { bold: true, color: { rgb: '1D4ED8' }, sz: 10 },
    fill: { fgColor: { rgb: 'DBEAFE' }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top:    { style: 'thin', color: { rgb: '9CA3AF' } },
      bottom: { style: 'thin', color: { rgb: '9CA3AF' } },
      left:   { style: 'thin', color: { rgb: '9CA3AF' } },
      right:  { style: 'thin', color: { rgb: '9CA3AF' } },
    },
  }
  const dataStyle = {
    font: { sz: 10 },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top:    { style: 'thin', color: { rgb: 'D1D5DB' } },
      bottom: { style: 'thin', color: { rgb: 'D1D5DB' } },
      left:   { style: 'thin', color: { rgb: 'D1D5DB' } },
      right:  { style: 'thin', color: { rgb: 'D1D5DB' } },
    },
  }

  const set = (r: number, c: number, v: XLSX.CellObject) => { ws[C(r, c)] = v }

  // ── Satır 0: Çalışan Adı | Masraf Yeri | Hafta No ──
  set(0, 0, { t: 's', v: 'Çalışan Adı:', s: labelStyle })
  addMerge(0, 0, 0, 1)
  set(0, 2, { t: 's', v: String(ts.calisan_adi ?? ''), s: valueStyle })
  addMerge(0, 2, 0, 3)
  set(0, 4, { t: 's', v: 'Masraf Yeri:', s: labelStyle })
  addMerge(0, 4, 0, 5)
  set(0, 6, { t: 's', v: String(ts.masraf_yeri ?? ''), s: valueStyle })
  addMerge(0, 6, 0, 7)
  set(0, 8, { t: 's', v: 'Hafta No:', s: labelStyle })
  addMerge(0, 8, 0, 9)
  set(0, 10, { t: 'n', v: Number(ts.hafta_no ?? 0), s: { ...numHStyle } })

  // ── Satır 1: Çalışan No | Masraf Yeri Kodu | Tarih ──
  set(1, 0, { t: 's', v: 'Çalışan No:', s: labelStyle })
  addMerge(1, 0, 1, 1)
  set(1, 2, { t: 's', v: String(ts.calisan_no ?? ''), s: valueStyle })
  addMerge(1, 2, 1, 3)
  set(1, 4, { t: 's', v: 'Masraf Yeri Kodu:', s: labelStyle })
  addMerge(1, 4, 1, 5)
  set(1, 6, { t: 's', v: String(ts.masraf_yeri_kodu ?? ''), s: valueStyle })
  addMerge(1, 6, 1, 7)
  set(1, 8, { t: 's', v: 'Tarih:', s: labelStyle })
  addMerge(1, 8, 1, 9)
  set(1, 10, { t: 's', v: ts.tarih ? String(ts.tarih).split('T')[0] : '', s: valueStyle })

  // ── Satır 2: Tablo başlık üst ──
  set(2, 0, { t: 's', v: 'İş Tipi',              s: hStyle }); addMerge(2, 0, 3, 0)
  set(2, 1, { t: 's', v: 'KOD',                  s: hStyle }); addMerge(2, 1, 3, 1)
  set(2, 2, { t: 's', v: 'Çalışılan\nMakine Kodu', s: hStyle }); addMerge(2, 2, 3, 2)
  set(2, 3, { t: 's', v: 'Çalışılan Süre (saat)', s: hStyle }); addMerge(2, 3, 2, 9)
  set(2, 10, { t: 's', v: 'NOTLAR',               s: hStyle }); addMerge(2, 10, 3, 10)

  // ── Satır 3: Gün başlıkları ──
  GUNLER.forEach((g, i) => set(3, 3 + i, { t: 's', v: g, s: hStyle }))

  // ── Veri satırları (4..13) ──
  const sortedRows = [...rows].sort((a, b) => Number(a.sira_no) - Number(b.sira_no))
  for (let i = 0; i < SATIR_SAYISI; i++) {
    const r = sortedRows[i]
    const rIdx = 4 + i
    set(rIdx, 0,  { t: 's', v: r ? String(r.is_tipi ?? '')    : '', s: dataStyle })
    set(rIdx, 1,  { t: 's', v: r ? String(r.kod ?? '')        : '', s: { ...dataStyle, alignment: { horizontal: 'center', vertical: 'center' } } })
    set(rIdx, 2,  { t: 's', v: r ? String(r.makine_kodu ?? '') : '', s: { ...dataStyle, alignment: { horizontal: 'center', vertical: 'center' } } })
    GUN_KEYS.forEach((g, gi) => {
      const val = r ? Number(r[g] ?? 0) : 0
      set(rIdx, 3 + gi, { t: 'n', v: val, s: dataStyle })
    })
    set(rIdx, 10, { t: 's', v: r ? String(r.notlar ?? '') : '', s: dataStyle })
  }

  // ── Satır 14: TOPLAM SÜRE ──
  const totRow = 4 + SATIR_SAYISI
  set(totRow, 0, { t: 's', v: 'TOPLAM SÜRE', s: numHStyle }); addMerge(totRow, 0, totRow, 2)
  GUN_KEYS.forEach((g, gi) => {
    const toplam = rows.reduce((acc, r) => acc + Number(r[g] ?? 0), 0)
    set(totRow, 3 + gi, { t: 'n', v: toplam, s: numHStyle })
  })
  set(totRow, 10, { t: 's', v: '', s: hStyle })

  ws['!merges'] = merges
  ws['!rows'] = [
    { hpx: 20 }, { hpx: 20 }, { hpx: 28 }, { hpx: 20 },
    ...Array(SATIR_SAYISI).fill({ hpx: 18 }),
    { hpx: 20 },
  ]

  const ref = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: totRow, c: 10 } })
  ws['!ref'] = ref

  XLSX.utils.book_append_sheet(wb, ws, 'Zaman Kaydı')
  return wb
}

function wbToBuffer(wb: XLSX.WorkBook): ArrayBuffer {
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true }) as ArrayBuffer
}

function safeName(name: string) {
  return name.replace(/[\\/:*?"<>|]/g, '_')
}

// ── Tek personel export ──
export async function exportOne(haftaNo: number, calisanAdi: string) {
  const sheets = await fetchTimesheetData(haftaNo, calisanAdi)
  if (!sheets.length) throw new Error(`Hafta ${haftaNo} için ${calisanAdi} kaydı bulunamadı.`)

  const ts = sheets[0] as Row
  const rows = ((ts.zamankay_timesheet_rows as Row[]) ?? [])
  const wb = buildWorkbook(ts, rows)
  const buf = wbToBuffer(wb)

  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `ZamanKaydi_Hafta${haftaNo}_${safeName(calisanAdi)}.xlsx`
  a.click()
  URL.revokeObjectURL(url)
}

// ── Tüm personel export (ZIP) ──
export async function exportAll(haftaNo: number) {
  const sheets = await fetchTimesheetData(haftaNo)
  if (!sheets.length) throw new Error(`Hafta ${haftaNo} için hiç kayıt bulunamadı.`)

  const zip = new JSZip()
  const folder = zip.folder(`ZamanKaydi_Hafta${haftaNo}`)!

  for (const ts of sheets as Row[]) {
    const rows = ((ts.zamankay_timesheet_rows as Row[]) ?? [])
    const wb = buildWorkbook(ts, rows)
    const buf = wbToBuffer(wb)
    const fileName = `${safeName(String(ts.calisan_adi))}.xlsx`
    folder.file(fileName, buf)
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' })
  const url = URL.createObjectURL(zipBlob)
  const a = document.createElement('a')
  a.href = url
  a.download = `ZamanKaydi_Hafta${haftaNo}_TumPersonel.zip`
  a.click()
  URL.revokeObjectURL(url)
}
