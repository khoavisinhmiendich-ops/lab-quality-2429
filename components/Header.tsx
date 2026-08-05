'use client';

import React, { useState, useEffect } from 'react';
import { Bell, Sun, Moon, ChevronDown, ShieldCheck, Key, Clock } from 'lucide-react';

export const Header: React.FC = () => {
  const [isDarkMode, setIsDarkMode] = useState<boolean>(false);
  const [currentTime, setCurrentTime] = useState<string>('');
  const [passKey, setPassKey] = useState<string>('2429-ADMIN');
  const [showPassKey, setShowPassKey] = useState<boolean>(false);

  useEffect(() => {
    const updateTime = () => {
      const now = new Date();
      setCurrentTime(
        now.toLocaleDateString('vi-VN', {
          weekday: 'short',
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
        })
      );
    };

    updateTime();
    const interval = setInterval(updateTime, 1000);
    return () => clearInterval(interval);
  }, []);

  return (
    <header className="relative bg-white border-b border-slate-200 px-6 py-3 shadow-sm shrink-0 overflow-hidden">
      <div
        className="absolute inset-0 opacity-10 bg-cover bg-center pointer-events-none"
        style={{
          backgroundImage: `url('https://images.unsplash.com/photo-1587354246490-7e26e63fdb6f?auto=format&fit=crop&w=1200&q=80')`,
        }}
      />

      <div className="relative flex items-center justify-between z-10">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <div className="w-12 h-12 rounded-full bg-rose-50 border border-rose-200 flex items-center justify-center p-1 shadow-sm">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img
                src="https://upload.wikimedia.org/wikipedia/commons/e/e8/Logo_Bo_Y_Te_Huong_Phai_Xanh.svg"
                alt="Logo Bệnh viện"
                className="w-full h-full object-contain"
                onError={(e) => {
                  (e.target as HTMLImageElement).style.display = 'none';
                }}
              />
            </div>
            <div className="w-12 h-12 rounded-full bg-sky-50 border border-sky-200 flex items-center justify-center p-1 shadow-sm">
              <div className="w-full h-full rounded-full bg-sky-700 flex items-center justify-center text-white text-[10px] font-bold text-center leading-none">
                KVSMD
              </div>
            </div>
          </div>

          <div>
            <h2 className="text-xs font-black tracking-wide text-slate-800 uppercase">
              BỆNH VIỆN PHONG - DA LIỄU TRUNG ƯƠNG QUY HÒA
            </h2>
            <h1 className="text-xl font-black text-sky-900 tracking-tight leading-tight">
              KHOA VI SINH - MIỄN DỊCH
            </h1>
            <p className="text-[11px] font-semibold text-slate-500 tracking-wide">
              QUẢN LÝ CHẤT LƯỢNG 2429 - SMARTLAB 360
            </p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="hidden lg:flex items-center gap-1.5 bg-slate-100 px-3 py-1.5 rounded-full border border-slate-200 text-slate-600 text-xs">
            <Clock className="w-3.5 h-3.5 text-sky-600" />
            <span className="font-mono font-medium">{currentTime || 'Đang tải...'}</span>
          </div>

          <div className="flex items-center gap-1 bg-slate-100 px-2.5 py-1 rounded-full border border-slate-200">
            <Key className="w-3.5 h-3.5 text-amber-500" />
            <input
              type={showPassKey ? 'text' : 'password'}
              value={passKey}
              onChange={(e) => setPassKey(e.target.value)}
              className="bg-transparent text-amber-600 font-mono w-20 outline-none text-xs"
              placeholder="Pass Key..."
            />
            <button
              type="button"
              onClick={() => setShowPassKey(!showPassKey)}
              className="text-[10px] text-slate-400 hover:text-slate-600 underline ml-0.5"
            >
              {showPassKey ? 'Ẩn' : 'Hiện'}
            </button>
          </div>

          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="w-9 h-9 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center transition-all border border-slate-200"
            title="Chuyển chế độ Sáng/Tối"
          >
            {isDarkMode ? <Moon className="w-4 h-4 text-amber-500" /> : <Sun className="w-4 h-4 text-amber-500" />}
          </button>

          <button
            className="relative w-9 h-9 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center transition-all border border-slate-200"
            title="Thông báo hệ thống"
          >
            <Bell className="w-4 h-4" />
            <span className="absolute -top-1 -right-1 w-4 h-4 bg-rose-500 text-white text-[9px] font-bold rounded-full flex items-center justify-center border-2 border-white">
              5
            </span>
          </button>

          <div className="flex items-center gap-2.5 bg-slate-50 hover:bg-slate-100 p-1.5 pr-3 rounded-full border border-slate-200 transition-all cursor-pointer">
            <div className="w-8 h-8 rounded-full bg-sky-600 text-white flex items-center justify-center font-bold text-xs shadow-sm">
              NV
            </div>
            <div className="text-left">
              <div className="text-xs font-bold text-slate-800 leading-tight flex items-center gap-1">
                <span>Nguyễn Thái Hòa</span>
                <ShieldCheck className="w-3 h-3 text-sky-600" />
              </div>
              <div className="text-[10px] font-medium text-slate-500 leading-tight">Quản trị viên</div>
            </div>
            <ChevronDown className="w-3.5 h-3.5 text-slate-400 ml-1" />
          </div>
        </div>
      </div>
    </header>
  );
};