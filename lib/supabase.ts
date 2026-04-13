import { createClient } from '@supabase/supabase-js'

// Build sırasında env var olmasa bile hata vermemesi için ?? '' kullanılıyor.
// Gerçek istekler runtime'da geldiğinde env var zaten mevcut olur.
export const supabase = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL ?? '',
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ?? ''
)

export type Timesheet = {
  id?: string
  calisan_adi: string
  calisan_no: string
  masraf_yeri: string
  masraf_yeri_kodu: string
  hafta_no: number | null
  tarih: string
  created_at?: string
}

export type TimesheetRow = {
  id?: string
  timesheet_id?: string
  sira_no: number
  is_tipi: string
  kod: string
  makine_kodu: string
  pazartesi: number
  sali: number
  carsamba: number
  persembe: number
  cuma: number
  cumartesi: number
  pazar: number
  notlar: string
}
