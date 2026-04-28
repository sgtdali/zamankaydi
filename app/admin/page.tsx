"use client";

import { useState, useEffect } from "react";
import { supabase, type Employee, type ProjectLocation } from "@/lib/supabase";
import { exportDetailedAllData } from "@/lib/excelExport";

export default function AdminPage() {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [error, setError] = useState("");
  
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [newEmpName, setNewEmpName] = useState("");
  const [newEmpNo, setNewEmpNo] = useState("");
  
  const [locations, setLocations] = useState<ProjectLocation[]>([]);
  const [newLocName, setNewLocName] = useState("");

  const [loading, setLoading] = useState(false);
  const [locLoading, setLocLoading] = useState(false);
  const [excelLoading, setExcelLoading] = useState(false);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (username === "repkon" && password === "repkonmontaj") {
      setIsLoggedIn(true);
      setError("");
    } else {
      setError("Hatalı kullanıcı adı veya şifre!");
    }
  };

  useEffect(() => {
    if (isLoggedIn) {
      fetchEmployees();
      fetchLocations();
    }
  }, [isLoggedIn]);

  async function fetchEmployees() {
    setLoading(true);
    const { data, error } = await supabase
      .from("zamankay_employees")
      .select("*")
      .order("ad", { ascending: true });
    
    if (!error && data) {
      setEmployees(data);
    }
    setLoading(false);
  }

  async function fetchLocations() {
    setLocLoading(true);
    const { data, error } = await supabase
      .from("zamankay_locations")
      .select("*")
      .order("ad", { ascending: true });
    
    if (!error && data) {
      setLocations(data);
    }
    setLocLoading(false);
  }

  async function addEmployee(e: React.FormEvent) {
    e.preventDefault();
    if (!newEmpName || !newEmpNo) return;

    const { error } = await supabase
      .from("zamankay_employees")
      .insert([{ ad: newEmpName.toUpperCase(), no: newEmpNo }]);

    if (error) {
      alert("Hata: " + error.message);
    } else {
      setNewEmpName("");
      setNewEmpNo("");
      fetchEmployees();
    }
  }

  async function deleteEmployee(id: string) {
    if (!confirm("Bu kişiyi silmek istediğinize emin misiniz?")) return;

    const { error } = await supabase
      .from("zamankay_employees")
      .delete()
      .eq("id", id);

    if (error) {
      alert("Hata: " + error.message);
    } else {
      fetchEmployees();
    }
  }

  async function addLocation(e: React.FormEvent) {
    e.preventDefault();
    if (!newLocName) return;

    const { error } = await supabase
      .from("zamankay_locations")
      .insert([{ ad: newLocName.toUpperCase() }]);

    if (error) {
      alert("Hata: " + error.message);
    } else {
      setNewLocName("");
      fetchLocations();
    }
  }

  async function deleteLocation(id: string) {
    if (!confirm("Bu lokasyonu silmek istediğinize emin misiniz?")) return;

    const { error } = await supabase
      .from("zamankay_locations")
      .delete()
      .eq("id", id);

    if (error) {
      alert("Hata: " + error.message);
    } else {
      fetchLocations();
    }
  }

  async function handleDetayExcel() {
    setExcelLoading(true);
    try {
      await exportDetailedAllData();
    } catch (e: unknown) {
      alert("Excel hatası: " + (e instanceof Error ? e.message : "Bilinmeyen hata"));
    } finally {
      setExcelLoading(false);
    }
  }

  if (isLoggedIn) {
    return (
      <div className="min-h-screen bg-zinc-950 flex flex-col items-center p-8 text-white font-[family-name:var(--font-geist-sans)]">
        <div className="w-full max-w-6xl animate-in fade-in duration-700">
          <header className="mb-12 flex flex-col md:flex-row justify-between items-start md:items-end gap-4">
            <div>
              <h1 className="text-4xl font-bold bg-gradient-to-r from-blue-400 to-emerald-400 bg-clip-text text-transparent">
                Admin Paneli
              </h1>
              <p className="text-zinc-400 mt-2">Hoş geldiniz, Repkon Admin.</p>
            </div>
            <button 
              onClick={() => setIsLoggedIn(false)}
              className="text-zinc-500 hover:text-white transition-colors flex items-center gap-2 group mb-1"
            >
              <span className="group-hover:-translate-x-1 transition-transform">←</span>
              <span>Çıkış Yap</span>
            </button>
          </header>
          
          <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
            {/* SOL KOLON: EKLEME FORMLARI VE EXCEL */}
            <div className="lg:col-span-1 space-y-8">
              {/* Kişi Ekleme */}
              <div className="bg-zinc-900/50 border border-zinc-800 p-6 rounded-2xl backdrop-blur-sm">
                <h3 className="text-lg font-semibold mb-6 flex items-center gap-2">
                  <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-emerald-400"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" x2="19" y1="8" y2="14"/><line x1="16" x2="22" y1="11" y2="11"/></svg>
                  Yeni Kişi Ekle
                </h3>
                <form onSubmit={addEmployee} className="space-y-4">
                  <div>
                    <label className="text-xs font-medium text-zinc-500 uppercase tracking-wider ml-1">Ad Soyad</label>
                    <input
                      type="text"
                      value={newEmpName}
                      onChange={(e) => setNewEmpName(e.target.value)}
                      placeholder="Örn: AHMET YILMAZ"
                      className="w-full bg-zinc-800/50 border border-zinc-700 rounded-xl px-4 py-2.5 mt-1 focus:outline-none focus:ring-2 focus:ring-blue-500/50 transition-all text-sm"
                    />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-zinc-500 uppercase tracking-wider ml-1">Personel No</label>
                    <input
                      type="text"
                      value={newEmpNo}
                      onChange={(e) => setNewEmpNo(e.target.value)}
                      placeholder="Örn: 1234"
                      className="w-full bg-zinc-800/50 border border-zinc-700 rounded-xl px-4 py-2.5 mt-1 focus:outline-none focus:ring-2 focus:ring-blue-500/50 transition-all text-sm"
                    />
                  </div>
                  <button
                    type="submit"
                    className="w-full bg-blue-600 hover:bg-blue-500 text-white font-semibold py-2.5 rounded-xl transition-all shadow-lg shadow-blue-600/20 active:scale-[0.98]"
                  >
                    Kişiyi Kaydet
                  </button>
                </form>
              </div>

              {/* Lokasyon Ekleme */}
              <div className="bg-zinc-900/50 border border-zinc-800 p-6 rounded-2xl backdrop-blur-sm">
                <h3 className="text-lg font-semibold mb-6 flex items-center gap-2">
                  <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-blue-400"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
                  Yeni Lokasyon Ekle
                </h3>
                <form onSubmit={addLocation} className="space-y-4">
                  <div>
                    <label className="text-xs font-medium text-zinc-500 uppercase tracking-wider ml-1">Lokasyon Adı</label>
                    <input
                      type="text"
                      value={newLocName}
                      onChange={(e) => setNewLocName(e.target.value)}
                      placeholder="Örn: ATÖLYE 1"
                      className="w-full bg-zinc-800/50 border border-zinc-700 rounded-xl px-4 py-2.5 mt-1 focus:outline-none focus:ring-2 focus:ring-blue-500/50 transition-all text-sm"
                    />
                  </div>
                  <button
                    type="submit"
                    className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-semibold py-2.5 rounded-xl transition-all shadow-lg shadow-emerald-600/20 active:scale-[0.98]"
                  >
                    Lokasyonu Kaydet
                  </button>
                </form>
              </div>

              {/* Excel */}
              <div className="bg-zinc-900/50 border border-zinc-800 p-6 rounded-2xl backdrop-blur-sm">
                <h4 className="text-xs font-medium text-zinc-400 mb-4 uppercase tracking-wider">Veri Yönetimi</h4>
                <button
                  onClick={handleDetayExcel}
                  disabled={excelLoading}
                  className="w-full bg-zinc-800 hover:bg-zinc-700 disabled:bg-zinc-900 disabled:text-zinc-600 text-zinc-200 font-semibold py-2.5 rounded-xl transition-all flex items-center justify-center gap-2 border border-zinc-700"
                >
                  <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-emerald-500">
                    <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/>
                    <polyline points="14 2 14 8 20 8"/>
                    <path d="M8 13h2"/>
                    <path d="M8 17h10"/>
                    <path d="M14 13h4"/>
                  </svg>
                  {excelLoading ? "Hazırlanıyor..." : "Excel İndir"}
                </button>
              </div>
            </div>

            {/* SAĞ KOLON: LİSTELER */}
            <div className="lg:col-span-3 grid grid-cols-1 md:grid-cols-2 gap-8 h-fit">
              {/* Kişi Listesi */}
              <div className="bg-zinc-900/50 border border-zinc-800 rounded-2xl backdrop-blur-sm overflow-hidden flex flex-col h-[600px]">
                <div className="p-6 border-b border-zinc-800 flex justify-between items-center bg-zinc-800/20">
                  <h3 className="font-semibold flex items-center gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-blue-400"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>
                    Kişi Listesi
                  </h3>
                  <span className="text-[10px] bg-zinc-800 px-2 py-1 rounded-full text-zinc-400 uppercase tracking-tighter">
                    {employees.length} Kayıt
                  </span>
                </div>
                
                <div className="flex-1 overflow-y-auto custom-scrollbar">
                  {loading ? (
                    <div className="p-12 text-center text-zinc-500">Yükleniyor...</div>
                  ) : employees.length === 0 ? (
                    <div className="p-12 text-center text-zinc-500 italic text-sm">Kayıt yok.</div>
                  ) : (
                    <table className="w-full text-left border-collapse">
                      <thead className="bg-zinc-800/30 text-[10px] text-zinc-500 uppercase tracking-widest sticky top-0 backdrop-blur-md">
                        <tr>
                          <th className="px-6 py-3 font-medium">Ad Soyad</th>
                          <th className="px-6 py-3 font-medium">No</th>
                          <th className="px-6 py-3 font-medium text-right"></th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-zinc-800/50 text-sm">
                        {employees.map((emp) => (
                          <tr key={emp.id} className="hover:bg-zinc-800/20 transition-colors group">
                            <td className="px-6 py-3 font-medium text-zinc-300">{emp.ad}</td>
                            <td className="px-6 py-3 text-zinc-500">{emp.no}</td>
                            <td className="px-6 py-3 text-right">
                              <button
                                onClick={() => deleteEmployee(emp.id!)}
                                className="text-zinc-700 hover:text-red-400 transition-colors p-1.5 rounded-lg hover:bg-red-400/10"
                              >
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/></svg>
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>

              {/* Lokasyon Listesi */}
              <div className="bg-zinc-900/50 border border-zinc-800 rounded-2xl backdrop-blur-sm overflow-hidden flex flex-col h-[600px]">
                <div className="p-6 border-b border-zinc-800 flex justify-between items-center bg-zinc-800/20">
                  <h3 className="font-semibold flex items-center gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-emerald-400"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
                    Lokasyon Listesi
                  </h3>
                  <span className="text-[10px] bg-zinc-800 px-2 py-1 rounded-full text-zinc-400 uppercase tracking-tighter">
                    {locations.length} Kayıt
                  </span>
                </div>
                
                <div className="flex-1 overflow-y-auto custom-scrollbar">
                  {locLoading ? (
                    <div className="p-12 text-center text-zinc-500">Yükleniyor...</div>
                  ) : locations.length === 0 ? (
                    <div className="p-12 text-center text-zinc-500 italic text-sm">Kayıt yok.</div>
                  ) : (
                    <table className="w-full text-left border-collapse">
                      <thead className="bg-zinc-800/30 text-[10px] text-zinc-500 uppercase tracking-widest sticky top-0 backdrop-blur-md">
                        <tr>
                          <th className="px-6 py-3 font-medium">Lokasyon Adı</th>
                          <th className="px-6 py-3 font-medium text-right"></th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-zinc-800/50 text-sm">
                        {locations.map((loc) => (
                          <tr key={loc.id} className="hover:bg-zinc-800/20 transition-colors group">
                            <td className="px-6 py-3 font-medium text-zinc-300">{loc.ad}</td>
                            <td className="px-6 py-3 text-right">
                              <button
                                onClick={() => deleteLocation(loc.id!)}
                                className="text-zinc-700 hover:text-red-400 transition-colors p-1.5 rounded-lg hover:bg-red-400/10"
                              >
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/></svg>
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
        <style jsx global>{`
          .custom-scrollbar::-webkit-scrollbar { width: 4px; }
          .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
          .custom-scrollbar::-webkit-scrollbar-thumb { background: #3f3f46; border-radius: 10px; }
          .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #52525b; }
        `}</style>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-zinc-950 flex items-center justify-center p-4 font-[family-name:var(--font-geist-sans)]">
      <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,_var(--tw-gradient-stops))] from-blue-900/20 via-zinc-950 to-zinc-950 -z-10" />
      
      <div className="w-full max-w-md bg-zinc-900/40 backdrop-blur-xl border border-zinc-800 p-8 rounded-3xl shadow-2xl animate-in fade-in zoom-in duration-500">
        <div className="mb-8 text-center">
          <div className="w-16 h-16 bg-gradient-to-br from-blue-500 to-emerald-500 rounded-2xl mx-auto mb-4 flex items-center justify-center shadow-lg shadow-blue-500/20">
            <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-white">
              <rect width="18" height="11" x="3" y="11" rx="2" ry="2"/>
              <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
            </svg>
          </div>
          <h2 className="text-2xl font-bold text-white">Admin Girişi</h2>
          <p className="text-zinc-400 mt-1">Devam etmek için kimlik bilgilerinizi girin</p>
        </div>

        <form onSubmit={handleLogin} className="space-y-6">
          <div className="space-y-2">
            <label className="text-sm font-medium text-zinc-300 ml-1">Kullanıcı Adı</label>
            <input
              type="text"
              required
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="w-full bg-zinc-800/50 border border-zinc-700 text-white rounded-xl px-4 py-3 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500 transition-all"
              placeholder="Kullanıcı adınızı girin"
            />
          </div>

          <div className="space-y-2">
            <label className="text-sm font-medium text-zinc-300 ml-1">Şifre</label>
            <input
              type="password"
              required
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full bg-zinc-800/50 border border-zinc-700 text-white rounded-xl px-4 py-3 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500 transition-all"
              placeholder="••••••••"
            />
          </div>

          {error && (
            <div className="bg-red-500/10 border border-red-500/20 text-red-400 text-sm p-3 rounded-xl animate-pulse">
              {error}
            </div>
          )}

          <button
            type="submit"
            className="w-full bg-gradient-to-r from-blue-600 to-emerald-600 hover:from-blue-500 hover:to-emerald-500 text-white font-semibold py-3 rounded-xl shadow-lg shadow-blue-600/20 active:scale-[0.98] transition-all"
          >
            Giriş Yap
          </button>
        </form>
      </div>
    </div>
  );
}
