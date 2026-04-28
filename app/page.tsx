'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect, useRef } from 'react'
import { supabase, type TimesheetRow } from '@/lib/supabase'
import ExcelModal from './ExcelModal'

// Çalışan listesi artık Supabase'den çekiliyor.
// Tip tanımlaması lib/supabase.ts içinde Employee olarak yapıldı.
import type { Employee, ProjectLocation } from '@/lib/supabase'

const GUNLER = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar'] as const
const GUN_KEYS = ['pazartesi', 'sali', 'carsamba', 'persembe', 'cuma', 'cumartesi', 'pazar'] as const

const IS_TIPLERI: { label: string; kod: string }[] = [
  { label: 'Montaj', kod: 'MM' },
  { label: 'Kurulum', kod: 'KU' },
  { label: 'Sevkiyat', kod: 'SE' },
  { label: 'Rework', kod: 'RE' },
  { label: 'Eğitim', kod: 'EG' },
  { label: 'Yalın Yönetim', kod: 'YG' },
  { label: 'Kalite Güvence', kod: 'KG' },
  { label: 'İş Bekleme', kod: 'IB' },
  { label: 'Parça Bekleme', kod: 'PB' },
  { label: 'Proses', kod: 'PR' },
  { label: 'Devreye Alma', kod: 'DA' },
  { label: 'Destek Proses', kod: 'DP' },
  { label: 'Destek Montaj', kod: 'DM' },
  { label: 'Destek Devreye Alma', kod: 'DD' },
  { label: 'Servis/Bakım', kod: 'SV' },
  { label: 'Raporlama', kod: 'RP' },
]

const KOD_MAP = Object.fromEntries(IS_TIPLERI.map(t => [t.label, t.kod]))

function isoHaftaNo(tarihStr: string): number {
  const d = new Date(tarihStr)
  d.setHours(0, 0, 0, 0)
  // ISO 8601: Perşembe'yi içeren haftanın haftası
  d.setDate(d.getDate() + 4 - (d.getDay() || 7))
  const yilBaslangic = new Date(d.getFullYear(), 0, 1)
  return Math.ceil(((d.getTime() - yilBaslangic.getTime()) / 86400000 + 1) / 7)
}

const BOŞ_SATIR = (): TimesheetRow => ({
  sira_no: 0,
  is_tipi: '',
  kod: '',
  makine_kodu: '',
  pazartesi: 0,
  sali: 0,
  carsamba: 0,
  persembe: 0,
  cuma: 0,
  cumartesi: 0,
  pazar: 0,
  notlar: '',
})

const SATIR_SAYISI = 10

async function mevcutKaydiBul(calisanAdi: string, haftaNo: number) {
  const { data, error } = await supabase
    .from('zamankay_timesheets')
    .select('id')
    .eq('calisan_adi', calisanAdi)
    .eq('hafta_no', haftaNo)
    .order('created_at', { ascending: false })
    .limit(1)
    .maybeSingle()

  if (error) throw error
  return data?.id ? String(data.id) : null
}

export default function ZamanKaydiForm() {
  const [calisan_adi, setCalisanAdi] = useState('')
  const [calisan_no, setCalisanNo] = useState('')
  const [masraf_yeri, setMasrafYeri] = useState('')
  const [masraf_yeri_kodu, setMasrafYeriKodu] = useState('')
  const [hafta_no, setHaftaNo] = useState<string>('')
  const [tarih, setTarih] = useState('')
  const [satirlar, setSatirlar] = useState<TimesheetRow[]>(
    Array.from({ length: SATIR_SAYISI }, (_, i) => ({ ...BOŞ_SATIR(), sira_no: i + 1 }))
  )
  const [kayit, setKayit] = useState<'idle' | 'loading' | 'success' | 'error'>('idle')
  const [hataMsg, setHataMsg] = useState('')
  const [excelModal, setExcelModal] = useState(false)
  const [mevcutId, setMevcutId] = useState<string | null>(null)
  const [sorgulanıyor, setSorgulanıyor] = useState(false)
  const [employees, setEmployees] = useState<Employee[]>([])
  const [locations, setLocations] = useState<ProjectLocation[]>([])
  const sorguRef = useRef<AbortController | null>(null)

  // Sayfa açıldığında verileri çek
  useEffect(() => {
    // Çalışanlar
    supabase.from('zamankay_employees').select('*').order('ad', { ascending: true })
      .then(({ data, error }) => {
        if (!error && data) setEmployees(data)
      })

    // Lokasyonlar
    supabase.from('zamankay_locations').select('*').order('ad', { ascending: true })
      .then(({ data, error }) => {
        if (!error && data) setLocations(data)
      })
  }, [])

  // Çalışan + hafta kombinasyonu değişince mevcut kaydı sorgula
  useEffect(() => {
    if (!calisan_adi || !hafta_no) {
      setMevcutId(null)
      return
    }

    // Önceki sorguyu iptal et
    sorguRef.current?.abort()
    const ctrl = new AbortController()
    sorguRef.current = ctrl

    setSorgulanıyor(true)

    supabase
      .from('zamankay_timesheets')
      .select('*, zamankay_timesheet_rows(*)')
      .eq('calisan_adi', calisan_adi)
      .eq('hafta_no', parseInt(hafta_no))
      .order('created_at', { ascending: false })
      .limit(1)
      .single()
      .then(({ data }) => {
        if (ctrl.signal.aborted) return
        if (data) {
          setMevcutId(data.id)
          setMasrafYeri(data.masraf_yeri ?? '')
          setMasrafYeriKodu(data.masraf_yeri_kodu ?? '')
          setTarih(prev => data.tarih ? String(data.tarih).split('T')[0] : prev)
          // Mevcut satırları forma yükle
          const dbRows: TimesheetRow[] = Array.from({ length: SATIR_SAYISI }, (_, i) => ({ ...BOŞ_SATIR(), sira_no: i + 1 }))
          const sorted = [...(data.zamankay_timesheet_rows ?? [])].sort((a: TimesheetRow, b: TimesheetRow) => a.sira_no - b.sira_no)
          sorted.forEach((r: TimesheetRow, i: number) => { if (i < SATIR_SAYISI) dbRows[i] = { ...r } })
          setSatirlar(dbRows)
        } else {
          setMevcutId(null)
          setMasrafYeri('')
          setMasrafYeriKodu('')
          setSatirlar(Array.from({ length: SATIR_SAYISI }, (_, i) => ({ ...BOŞ_SATIR(), sira_no: i + 1 })))
        }
        setSorgulanıyor(false)
      })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [calisan_adi, hafta_no])

  const satirGuncelle = (idx: number, alan: keyof TimesheetRow, deger: string | number) => {
    setSatirlar(prev => prev.map((s, i) => i === idx ? { ...s, [alan]: deger } : s))
  }

  const toplamHesapla = (gun: typeof GUN_KEYS[number]) =>
    satirlar.reduce((acc, s) => acc + (Number(s[gun]) || 0), 0)

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault()
    if (!calisan_adi.trim()) {
      setHataMsg('Çalışan adı zorunludur.')
      return
    }
    if (!masraf_yeri.trim()) {
      setHataMsg('Lokasyon seçimi zorunludur.')
      return
    }
    if (!hafta_no || Number(hafta_no) < 1 || Number(hafta_no) > 53) {
      setHataMsg('Hafta no zorunludur.')
      return
    }
    if (!tarih) {
      setHataMsg('Tarih zorunludur.')
      return
    }

    // Satır doğrulaması: Süre girilmişse İş Tipi zorunlu
    for (let i = 0; i < satirlar.length; i++) {
      const row = satirlar[i]
      const hasHours = GUN_KEYS.some(key => Number(row[key] || 0) > 0)
      if (hasHours && !row.is_tipi) {
        setHataMsg(`${i + 1}. satırda çalışma süresi girilmiş ancak "İş Tipi" seçilmemiş.`)
        return
      }
    }

    setKayit('loading')
    setHataMsg('')

    const doluSatirlar = satirlar.filter(s =>
      s.is_tipi || s.kod || s.makine_kodu ||
      GUN_KEYS.some(g => Number(s[g]) > 0) || s.notlar
    )

    const headerPayload = {
      calisan_adi: calisan_adi.trim(),
      calisan_no: calisan_no.trim(),
      masraf_yeri: masraf_yeri.trim(),
      masraf_yeri_kodu: masraf_yeri_kodu.trim(),
      hafta_no: parseInt(hafta_no),
      tarih,
      updated_at: new Date().toISOString(),
    }

    let tsId = mevcutId

    try {
      tsId = await mevcutKaydiBul(headerPayload.calisan_adi, headerPayload.hafta_no)
    } catch (e: unknown) {
      setKayit('error')
      setHataMsg('Mevcut kayıt kontrol edilemedi: ' + (e instanceof Error ? e.message : 'Bilinmeyen hata'))
      return
    }

    if (tsId) {
      // Güncelle
      const { error: upErr } = await supabase
        .from('zamankay_timesheets')
        .update(headerPayload)
        .eq('id', tsId)
      if (upErr) {
        setKayit('error')
        setHataMsg('Güncelleme hatası: ' + upErr.message)
        return
      }
      // Eski satırları sil, yenilerini yaz
      const { error: delErr } = await supabase.from('zamankay_timesheet_rows').delete().eq('timesheet_id', tsId)
      if (delErr) {
        setKayit('error')
        setHataMsg('Eski satırlar silinemedi: ' + delErr.message)
        return
      }
    } else {
      // Yeni kayıt
      const { data: ts, error: tsErr } = await supabase
        .from('zamankay_timesheets')
        .insert(headerPayload)
        .select()
        .single()
      if (tsErr || !ts) {
        setKayit('error')
        setHataMsg('Kayıt oluşturulamadı: ' + (tsErr?.message || 'Bilinmeyen hata'))
        return
      }
      tsId = ts.id
    }
    setMevcutId(tsId)

    if (doluSatirlar.length > 0) {
      const { error: rowErr } = await supabase
        .from('zamankay_timesheet_rows')
        .insert(doluSatirlar.map((s, i) => {
          // eslint-disable-next-line @typescript-eslint/no-unused-vars
          const { id: _id, timesheet_id: _tid, ...rest } = s
          return { ...rest, sira_no: i + 1, timesheet_id: tsId }
        }))

      if (rowErr) {
        setKayit('error')
        setHataMsg('Satırlar kaydedilemedi: ' + rowErr.message)
        return
      }
    }

    setKayit('success')
    setTimeout(() => setKayit('idle'), 3000)
  }


  return (
    <div className="min-h-screen bg-gray-50 py-8 px-4">
      <form onSubmit={handleSubmit} className="max-w-[1300px] mx-auto">
        {/* Başlık */}
        <div className="text-center mb-4">
          <h1 className="text-xl font-bold text-gray-800 tracking-wide uppercase">
            Haftalık Zaman Kaydı
          </h1>
        </div>

        {/* Header Bilgileri */}
        <div className="border border-gray-400 bg-white mb-0 divide-y divide-gray-400">
          <div className="flex items-center px-3 py-2 gap-2">
            <label className="text-sm font-semibold whitespace-nowrap w-36">Çalışan Adı:</label>
            <select
              value={calisan_adi}
              onChange={e => {
                const val = e.target.value
                setCalisanAdi(val)
                const emp = employees.find(c => c.ad === val)
                setCalisanNo(emp?.no ?? '')
              }}
              className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent cursor-pointer"
            >
              <option value="" />
              {employees.map(c => (
                <option key={c.ad} value={c.ad}>{c.ad}</option>
              ))}
            </select>
          </div>
          <div className="flex items-center px-3 py-2 gap-2">
            <label className="text-sm font-semibold whitespace-nowrap w-36">Çalışan No:</label>
            <input
              type="text"
              value={calisan_no}
              readOnly
              className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent text-gray-600"
            />
          </div>
          <div className="flex items-center px-3 py-2 gap-2">
            <label className="text-sm font-semibold whitespace-nowrap w-36">Lokasyon:</label>
            <select
              value={masraf_yeri}
              required
              onChange={e => setMasrafYeri(e.target.value)}
              className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent cursor-pointer"
            >
              <option value="" />
              {locations.map(loc => (
                <option key={loc.id} value={loc.ad}>{loc.ad}</option>
              ))}
            </select>
          </div>
          <div className="flex items-center px-3 py-2 gap-2">
            <label className="text-sm font-semibold whitespace-nowrap w-36">Hafta No:</label>
            <input
              type="number"
              min={1}
              max={53}
              required
              value={hafta_no}
              onChange={e => setHaftaNo(e.target.value)}
              className="w-20 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent text-blue-600 font-semibold text-center"
            />
          </div>
          <div className="flex items-center px-3 py-2 gap-2">
            <label className="text-sm font-semibold whitespace-nowrap w-36">Tarih:</label>
            <input
              type="date"
              required
              value={tarih}
              onChange={e => {
                const val = e.target.value
                setTarih(val)
                if (val) setHaftaNo(String(isoHaftaNo(val)))
              }}
              className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent"
            />
          </div>
        </div>

        {/* Tablo */}
        <div className="overflow-x-auto border border-t-0 border-gray-400 bg-white">
          <table className="w-full border-collapse text-sm">
            <thead>
              <tr className="bg-[#dbeafe]">
                <th rowSpan={2} className="border border-gray-400 px-2 py-2 text-center align-middle font-semibold w-44">
                  İş Tipi
                </th>
                <th rowSpan={2} className="border border-gray-400 px-2 py-2 text-center align-middle font-semibold w-20">
                  KOD
                </th>
                <th rowSpan={2} className="border border-gray-400 px-2 py-2 text-center align-middle font-semibold w-28">
                  Çalışılan<br />Makine Kodu
                </th>
                <th colSpan={7} className="border border-gray-400 px-2 py-2 text-center font-semibold">
                  Çalışılan Süre (saat)
                </th>
                <th rowSpan={2} className="border border-gray-400 px-2 py-2 text-center align-middle font-semibold w-32">
                  NOTLAR
                </th>
              </tr>
              <tr className="bg-[#dbeafe]">
                {GUNLER.map(gun => (
                  <th key={gun} className="border border-gray-400 px-1 py-2 text-center font-semibold w-20 text-xs">
                    {gun}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {satirlar.map((satir, idx) => (
                <tr key={idx} className={idx % 2 === 0 ? 'bg-white hover:bg-blue-50' : 'bg-gray-50 hover:bg-blue-50'}>
                  <td className="border border-gray-300 p-0">
                    <select
                      value={satir.is_tipi}
                      required={GUN_KEYS.some(gun => Number(satir[gun] || 0) > 0)}
                      onChange={e => {
                        const val = e.target.value
                        setSatirlar(prev => prev.map((s, i) =>
                          i === idx ? { ...s, is_tipi: val, kod: KOD_MAP[val] ?? s.kod } : s
                        ))
                      }}
                      className="w-full h-8 px-1 outline-none text-sm bg-transparent cursor-pointer"
                    >
                      <option value="" />
                      {IS_TIPLERI.map(t => (
                        <option key={t.kod} value={t.label}>{t.label}</option>
                      ))}
                    </select>
                  </td>
                  <td className="border border-gray-300 p-0">
                    <input
                      type="text"
                      value={satir.kod}
                      onChange={e => satirGuncelle(idx, 'kod', e.target.value)}
                      className="w-full h-8 px-2 outline-none text-sm bg-transparent text-center"
                    />
                  </td>
                  <td className="border border-gray-300 p-0">
                    <input
                      type="text"
                      value={satir.makine_kodu}
                      onChange={e => satirGuncelle(idx, 'makine_kodu', e.target.value)}
                      className="w-full h-8 px-2 outline-none text-sm bg-transparent text-center"
                    />
                  </td>
                  {GUN_KEYS.map(gun => (
                    <td key={gun} className="border border-gray-300 p-0">
                      <input
                        type="number"
                        min={0}
                        max={24}
                        step={0.25}
                        value={satir[gun] === 0 ? '' : satir[gun]}
                        onChange={e => satirGuncelle(idx, gun, e.target.value === '' ? 0 : parseFloat(e.target.value))}
                        className="w-full h-8 px-1 outline-none text-sm bg-transparent text-center"
                        placeholder="0"
                      />
                    </td>
                  ))}
                  <td className="border border-gray-300 p-0">
                    <input
                      type="text"
                      value={satir.notlar}
                      onChange={e => satirGuncelle(idx, 'notlar', e.target.value)}
                      className="w-full h-8 px-2 outline-none text-sm bg-transparent"
                    />
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="bg-[#dbeafe]">
                <td colSpan={3} className="border border-gray-400 px-3 py-2 text-center font-bold text-blue-700 text-sm uppercase tracking-wide">
                  TOPLAM SÜRE
                </td>
                {GUN_KEYS.map(gun => (
                  <td key={gun} className="border border-gray-400 px-1 py-2 text-center font-semibold text-sm text-blue-700">
                    {toplamHesapla(gun).toFixed(2)}
                  </td>
                ))}
                <td className="border border-gray-400" />
              </tr>
            </tfoot>
          </table>
        </div>

        {/* Alt Butonlar */}
        <div className="mt-4 flex items-center gap-4">
          <button
            type="submit"
            disabled={kayit === 'loading'}
            className="bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white font-semibold px-8 py-2 rounded transition-colors text-sm"
          >
            {kayit === 'loading' ? 'Kaydediliyor...' : 'Kaydet'}
          </button>
          <button
            type="button"
            onClick={() => {
              setCalisanAdi(''); setCalisanNo(''); setMasrafYeri(''); setMasrafYeriKodu('')
              setHaftaNo(''); setTarih('')
              setSatirlar(Array.from({ length: SATIR_SAYISI }, (_, i) => ({ ...BOŞ_SATIR(), sira_no: i + 1 })))
              setKayit('idle'); setHataMsg('')
            }}
            className="bg-gray-200 hover:bg-gray-300 text-gray-700 font-semibold px-8 py-2 rounded transition-colors text-sm"
          >
            Temizle
          </button>

          <button
            type="button"
            onClick={() => setExcelModal(true)}
            className="bg-green-600 hover:bg-green-700 text-white font-semibold px-8 py-2 rounded transition-colors text-sm"
          >
            Excel Çıktısı
          </button>


          {sorgulanıyor && (
            <span className="text-blue-500 text-sm">Kayıt sorgulanıyor...</span>
          )}
          {!sorgulanıyor && mevcutId && kayit === 'idle' && (
            <span className="text-amber-600 text-sm font-medium">Mevcut kayıt yüklendi — üzerine yazılacak</span>
          )}

          {kayit === 'success' && (
            <span className="text-green-600 font-semibold text-sm">
              {mevcutId ? 'Kayıt güncellendi!' : 'Kayıt oluşturuldu!'}
            </span>
          )}
          {(kayit === 'error' || hataMsg) && (
            <span className="text-red-600 font-semibold text-sm">
              {hataMsg}
            </span>
          )}
        </div>
      </form>

      {excelModal && <ExcelModal onClose={() => setExcelModal(false)} />}
    </div>
  )
}
