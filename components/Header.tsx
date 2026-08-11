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
    <header className="relative bg-white/90 backdrop-blur-md border-b border-slate-200/80 px-6 py-3 shadow-[0_1px_0_rgba(15,50,55,0.04),0_8px_24px_-16px_rgba(15,50,55,0.15)] shrink-0 overflow-hidden font-ui">
      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500&display=swap');
        .font-display { font-family: 'Fraunces', 'Times New Roman', serif; }
        .font-ui { font-family: 'Inter', ui-sans-serif, system-ui, sans-serif; }

        @keyframes riseIn {
          from { opacity: 0; transform: translateY(-6px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-riseIn { animation: riseIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes iconFlip {
          from { opacity: 0; transform: rotate(-90deg) scale(0.5); }
          to { opacity: 1; transform: rotate(0deg) scale(1); }
        }
        .animate-iconFlip { animation: iconFlip 0.35s cubic-bezier(0.34, 1.56, 0.64, 1) both; }

        @keyframes badgePop {
          0% { transform: scale(0.6); opacity: 0; }
          60% { transform: scale(1.15); opacity: 1; }
          100% { transform: scale(1); }
        }
        .animate-badgePop { animation: badgePop 0.4s cubic-bezier(0.34, 1.56, 0.64, 1) both; }

        @media (prefers-reduced-motion: reduce) {
          .animate-riseIn, .animate-iconFlip, .animate-badgePop { animation: none !important; }
        }
      `}</style>

      <div
        className="absolute inset-0 opacity-[0.06] bg-cover bg-center pointer-events-none"
        style={{
          backgroundImage: `url('https://images.unsplash.com/photo-1587354246490-7e26e63fdb6f?auto=format&fit=crop&w=1200&q=80')`,
        }}
      />
      <div className="absolute inset-x-0 top-0 h-px bg-gradient-to-r from-transparent via-teal-500/40 to-transparent" />

      <div className="relative flex items-center justify-between z-10 animate-riseIn">
        <div className="flex items-center gap-4">
          <div className="flex items-center -space-x-2">
            <div className="w-12 h-12 rounded-full bg-rose-50 border-2 border-white ring-1 ring-rose-200 flex items-center justify-center p-1.5 shadow-sm z-10">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img
                src="https://upload.wikimedia.org/wikipedia/commons/e/e8/Logo_Bo_Y_Te_Huong_Phai_Xanh.svg"
                alt="https://bvquyhoa.vn/wp-content/uploads/2022/03/LOGO-390x220.png"
                className="w-full h-full object-contain"
                onError={(e) => {
                  (e.target as HTMLImageElement).style.display = 'none';
                }}
              />
            </div>
            <div className="w-12 h-12 rounded-full bg-teal-50 border-2 border-white ring-1 ring-teal-200 flex items-center justify-center p-1 shadow-sm">
              <div className="w-full h-full rounded-full bg-gradient-to-br from-[#0E3A41] to-teal-700 flex items-center justify-center text-white text-[10px] font-bold text-center leading-none tracking-tight">
                KVSMD
              </div>
            </div>
          </div>

          <div className="pl-1">
            <h2 className="text-[10.5px] font-bold tracking-wide text-slate-500 uppercase">
              Bệnh viện Phong - Da liễu Trung ương Quy Hòa
            </h2>
            <h1 className="font-display text-xl font-semibold text-[#0E3A41] tracking-tight leading-tight">
              Khoa Vi Sinh - Miễn Dịch
            </h1>
            <p className="text-[10.5px] font-semibold text-teal-700/70 tracking-wide uppercase">
              Quản lý chất lượng 2429 &middot; Khoa Vi Sinh - Miễn Dịch
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <div
            className="hidden lg:flex items-center gap-1.5 bg-slate-50 px-3 py-1.5 rounded-full border border-slate-200 text-slate-600 text-xs"
            style={{ fontFamily: '"IBM Plex Mono", ui-monospace, monospace' }}
          >
            <span className="relative flex h-1.5 w-1.5 mr-0.5">
              <span className="absolute inline-flex h-full w-full rounded-full bg-teal-400 opacity-60 animate-ping" />
              <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-teal-500" />
            </span>
            <Clock className="w-3.5 h-3.5 text-teal-600" />
            <span className="font-medium">{currentTime || 'Đang tải...'}</span>
          </div>

          <div className="flex items-center gap-1 bg-amber-50/70 px-2.5 py-1 rounded-full border border-amber-200/80 transition-all duration-200 focus-within:ring-4 focus-within:ring-amber-500/15 focus-within:border-amber-300">
            <Key className="w-3.5 h-3.5 text-amber-500 shrink-0" />
            <input
              type={showPassKey ? 'text' : 'password'}
              value={passKey}
              onChange={(e) => setPassKey(e.target.value)}
              className="bg-transparent text-amber-700 font-medium w-20 outline-none text-xs"
              style={{ fontFamily: '"IBM Plex Mono", ui-monospace, monospace' }}
              placeholder="Pass Key..."
            />
            <button
              type="button"
              onClick={() => setShowPassKey(!showPassKey)}
              className="text-[10px] text-amber-500/80 hover:text-amber-700 font-semibold ml-0.5 cursor-pointer transition-colors"
            >
              {showPassKey ? 'Ẩn' : 'Hiện'}
            </button>
          </div>

          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="w-9 h-9 rounded-full bg-slate-50 hover:bg-slate-100 text-slate-600 flex items-center justify-center transition-all duration-200 border border-slate-200 hover:-translate-y-px cursor-pointer overflow-hidden"
            title="Chuyển chế độ Sáng/Tối"
          >
            <span key={isDarkMode ? 'moon' : 'sun'} className="animate-iconFlip flex items-center justify-center">
              {isDarkMode ? <Moon className="w-4 h-4 text-teal-600" /> : <Sun className="w-4 h-4 text-amber-500" />}
            </span>
          </button>

          <button
            className="relative w-9 h-9 rounded-full bg-slate-50 hover:bg-slate-100 text-slate-600 flex items-center justify-center transition-all duration-200 border border-slate-200 hover:-translate-y-px cursor-pointer"
            title="Thông báo hệ thống"
          >
            <Bell className="w-4 h-4" />
            <span className="absolute -top-1 -right-1 w-4 h-4 bg-rose-500 text-white text-[9px] font-bold rounded-full flex items-center justify-center border-2 border-white animate-badgePop">
              5
            </span>
          </button>

          <div className="group flex items-center gap-2.5 bg-slate-50 hover:bg-slate-100 p-1.5 pr-3 rounded-full border border-slate-200 transition-all duration-200 hover:-translate-y-px cursor-pointer">
            <div className="w-8 h-8 rounded-full bg-gradient-to-br from-[#0E3A41] to-teal-700 text-white flex items-center justify-center font-bold text-xs shadow-sm ring-2 ring-white">
              NV
            </div>
            <div className="text-left">
              <div className="text-xs font-bold text-slate-800 leading-tight flex items-center gap-1">
                <span>Nguyễn Thái Hòa</span>
                <ShieldCheck className="w-3 h-3 text-teal-600" />
              </div>
              <div className="text-[10px] font-medium text-slate-500 leading-tight">Quản trị viên</div>
            </div>
            <ChevronDown className="w-3.5 h-3.5 text-slate-400 ml-1 transition-transform duration-200 group-hover:rotate-180" />
          </div>
        </div>
      </div>
    </header>
  );
};
