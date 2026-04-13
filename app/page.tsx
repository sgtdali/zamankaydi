'use client'

import { useState } from 'react'
import { supabase, type TimesheetRow } from '@/lib/supabase'

const GUNLER = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar'] as const
const GUN_KEYS = ['pazartesi', 'sali', 'carsamba', 'persembe', 'cuma', 'cumartesi', 'pazar'] as const

const IS_TIPLERI: { label: string; kod: string }[] = [
  { label: 'Montaj',           kod: 'MM' },
  { label: 'Kurulum',          kod: 'KU' },
  { label: 'Sevkiyat',         kod: 'SE' },
  { label: 'Rework',           kod: 'RE' },
  { label: 'Eğitim',           kod: 'EG' },
  { label: 'Yalın Yönetim',    kod: 'YG' },
  { label: 'Kalite Güvence',   kod: 'KG' },
  { label: 'İş Bekleme',       kod: 'IB' },
  { label: 'Parça Bekleme',    kod: 'PB' },
  { label: 'Proses',           kod: 'PR' },
  { label: 'Devreye Alma',     kod: 'DA' },
  { label: 'Destek Proses',    kod: 'DP' },
  { label: 'Destek Montaj',    kod: 'DM' },
  { label: 'Destek Devreye Alma', kod: 'DD' },
  { label: 'Servis/Bakım',     kod: 'SV' },
  { label: 'Raporlama',        kod: 'RP' },
]

const KOD_MAP = Object.fromEntries(IS_TIPLERI.map(t => [t.label, t.kod]))

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

    setKayit('loading')
    setHataMsg('')

    const doluSatirlar = satirlar.filter(s =>
      s.is_tipi || s.kod || s.makine_kodu ||
      GUN_KEYS.some(g => Number(s[g]) > 0) || s.notlar
    )

    const { data: ts, error: tsErr } = await supabase
      .from('zamankay_timesheets')
      .insert({
        calisan_adi: calisan_adi.trim(),
        calisan_no: calisan_no.trim(),
        masraf_yeri: masraf_yeri.trim(),
        masraf_yeri_kodu: masraf_yeri_kodu.trim(),
        hafta_no: hafta_no ? parseInt(hafta_no) : null,
        tarih: tarih || null,
      })
      .select()
      .single()

    if (tsErr || !ts) {
      setKayit('error')
      setHataMsg('Kayıt oluşturulamadı: ' + (tsErr?.message || 'Bilinmeyen hata'))
      return
    }

    if (doluSatirlar.length > 0) {
      const { error: rowErr } = await supabase
        .from('zamankay_timesheet_rows')
        .insert(doluSatirlar.map((s, i) => ({ ...s, sira_no: i + 1, timesheet_id: ts.id })))

      if (rowErr) {
        setKayit('error')
        setHataMsg('Satırlar kaydedilemedi: ' + rowErr.message)
        return
      }
    }

    setKayit('success')
    setTimeout(() => {
      setCalisanAdi('')
      setCalisanNo('')
      setMasrafYeri('')
      setMasrafYeriKodu('')
      setHaftaNo('')
      setTarih('')
      setSatirlar(Array.from({ length: SATIR_SAYISI }, (_, i) => ({ ...BOŞ_SATIR(), sira_no: i + 1 })))
      setKayit('idle')
    }, 2500)
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
        <div className="border border-gray-400 bg-white mb-0">
          <div className="grid grid-cols-3 divide-x divide-gray-400">
            {/* Sol */}
            <div className="divide-y divide-gray-400">
              <div className="flex items-center px-3 py-2 gap-2">
                <label className="text-sm font-semibold whitespace-nowrap w-36">Çalışan Adı:</label>
                <input
                  type="text"
                  value={calisan_adi}
                  onChange={e => setCalisanAdi(e.target.value)}
                  className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent"
                />
              </div>
              <div className="flex items-center px-3 py-2 gap-2">
                <label className="text-sm font-semibold whitespace-nowrap w-36">Çalışan No:</label>
                <input
                  type="text"
                  value={calisan_no}
                  onChange={e => setCalisanNo(e.target.value)}
                  className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent"
                />
              </div>
            </div>
            {/* Orta */}
            <div className="divide-y divide-gray-400">
              <div className="flex items-center px-3 py-2 gap-2">
                <label className="text-sm font-semibold whitespace-nowrap w-36">Masraf Yeri:</label>
                <input
                  type="text"
                  value={masraf_yeri}
                  onChange={e => setMasrafYeri(e.target.value)}
                  className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent"
                />
              </div>
              <div className="flex items-center px-3 py-2 gap-2">
                <label className="text-sm font-semibold whitespace-nowrap w-36">Masraf Yeri Kodu:</label>
                <input
                  type="text"
                  value={masraf_yeri_kodu}
                  onChange={e => setMasrafYeriKodu(e.target.value)}
                  className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent"
                />
              </div>
            </div>
            {/* Sağ */}
            <div className="divide-y divide-gray-400">
              <div className="flex items-center px-3 py-2 gap-2">
                <label className="text-sm font-semibold whitespace-nowrap w-24">Hafta No:</label>
                <input
                  type="number"
                  min={1}
                  max={53}
                  value={hafta_no}
                  onChange={e => setHaftaNo(e.target.value)}
                  className="w-16 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent text-blue-600 font-semibold text-center"
                />
              </div>
              <div className="flex items-center px-3 py-2 gap-2">
                <label className="text-sm font-semibold whitespace-nowrap w-24">Tarih:</label>
                <input
                  type="date"
                  value={tarih}
                  onChange={e => setTarih(e.target.value)}
                  className="flex-1 border-b border-gray-400 outline-none text-sm px-1 py-0.5 bg-transparent"
                />
              </div>
            </div>
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

          {kayit === 'success' && (
            <span className="text-green-600 font-semibold text-sm">
              Kayıt başarıyla oluşturuldu!
            </span>
          )}
          {(kayit === 'error' || hataMsg) && (
            <span className="text-red-600 font-semibold text-sm">
              {hataMsg}
            </span>
          )}
        </div>
      </form>
    </div>
  )
}
