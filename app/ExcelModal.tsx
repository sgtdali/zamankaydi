'use client'

import { useState } from 'react'
import { exportOne, exportAll } from '@/lib/excelExport'

const CALISANLAR = [
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

type Props = { onClose: () => void }

export default function ExcelModal({ onClose }: Props) {
  const [haftaNo, setHaftaNo] = useState('')
  const [kisi, setKisi] = useState<'tumü' | 'tekil'>('tekil')
  const [seciliKisi, setSeciliKisi] = useState('')
  const [durum, setDurum] = useState<'idle' | 'loading' | 'error'>('idle')
  const [hata, setHata] = useState('')

  const handleIndir = async () => {
    if (!haftaNo || Number(haftaNo) < 1) {
      setHata('Geçerli bir hafta numarası giriniz.'); return
    }
    if (kisi === 'tekil' && !seciliKisi) {
      setHata('Lütfen bir personel seçiniz.'); return
    }

    setDurum('loading')
    setHata('')
    try {
      if (kisi === 'tumü') {
        await exportAll(Number(haftaNo))
      } else {
        await exportOne(Number(haftaNo), seciliKisi)
      }
      onClose()
    } catch (e: unknown) {
      setHata(e instanceof Error ? e.message : 'Bilinmeyen hata')
      setDurum('error')
    } finally {
      setDurum(prev => prev === 'loading' ? 'idle' : prev)
    }
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      onClick={e => { if (e.target === e.currentTarget) onClose() }}
    >
      <div className="bg-white rounded-lg shadow-2xl w-full max-w-md p-6">
        <div className="flex items-center justify-between mb-5">
          <h2 className="text-base font-bold text-gray-800">Excel Çıktısı Al</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-xl leading-none">&times;</button>
        </div>

        {/* Hafta No */}
        <div className="mb-4">
          <label className="block text-sm font-semibold text-gray-700 mb-1">Hafta No</label>
          <input
            type="number"
            min={1}
            max={53}
            value={haftaNo}
            onChange={e => setHaftaNo(e.target.value)}
            placeholder="Örn: 17"
            className="w-full border border-gray-300 rounded px-3 py-2 text-sm outline-none focus:border-blue-500"
          />
        </div>

        {/* Personel Seçimi */}
        <div className="mb-4">
          <label className="block text-sm font-semibold text-gray-700 mb-2">Personel</label>
          <div className="flex gap-4 mb-3">
            <label className="flex items-center gap-2 cursor-pointer text-sm">
              <input
                type="radio"
                value="tekil"
                checked={kisi === 'tekil'}
                onChange={() => setKisi('tekil')}
                className="accent-blue-600"
              />
              Tek personel
            </label>
            <label className="flex items-center gap-2 cursor-pointer text-sm">
              <input
                type="radio"
                value="tumü"
                checked={kisi === 'tumü'}
                onChange={() => setKisi('tumü')}
                className="accent-blue-600"
              />
              Tüm personel <span className="text-gray-400 text-xs">(ZIP)</span>
            </label>
          </div>

          {kisi === 'tekil' && (
            <select
              value={seciliKisi}
              onChange={e => setSeciliKisi(e.target.value)}
              className="w-full border border-gray-300 rounded px-3 py-2 text-sm outline-none focus:border-blue-500"
            >
              <option value="">— Personel seçiniz —</option>
              {CALISANLAR.map(ad => (
                <option key={ad} value={ad}>{ad}</option>
              ))}
            </select>
          )}
        </div>

        {/* Hata */}
        {hata && (
          <p className="text-red-600 text-sm mb-3">{hata}</p>
        )}

        {/* Butonlar */}
        <div className="flex gap-3 mt-5">
          <button
            onClick={handleIndir}
            disabled={durum === 'loading'}
            className="flex-1 bg-green-600 hover:bg-green-700 disabled:bg-green-400 text-white font-semibold py-2 rounded text-sm transition-colors"
          >
            {durum === 'loading'
              ? kisi === 'tumü' ? 'ZIP oluşturuluyor...' : 'Excel oluşturuluyor...'
              : 'İndir'}
          </button>
          <button
            onClick={onClose}
            className="flex-1 bg-gray-200 hover:bg-gray-300 text-gray-700 font-semibold py-2 rounded text-sm transition-colors"
          >
            İptal
          </button>
        </div>
      </div>
    </div>
  )
}
