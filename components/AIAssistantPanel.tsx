// components/AIAssistantPanel.tsx
import React, { useState } from 'react';
import { Bot, Search, Sparkles, Send } from 'lucide-react';

export const AIAssistantPanel: React.FC = () => {
  const [searchTerm, setSearchTerm] = useState('');
  const [query, setQuery] = useState('');
  const [isThinking, setIsThinking] = useState(false);
  const [messages, setMessages] = useState<Array<{ role: 'user' | 'ai'; text: string }>>([
    { role: 'ai', text: 'Xin chào! Tôi là Trợ lý AI Quản lý Chất lượng 2429. Bạn cần hỗ trợ gì về các tiêu chí hoặc hành động khắc phục (CAPA)?' }
  ]);

  const handleAISend = () => {
    if (!query.trim()) return;
    const userMsg = query;
    setMessages(prev => [...prev, { role: 'user', text: userMsg }]);
    setQuery('');
    setIsThinking(true);

    // Simulated AI CAPA & Gap Analysis response logic
    setTimeout(() => {
      setMessages(prev => [
        ...prev,
        {
          role: 'ai',
          text: `Theo Tiêu chí QĐ 2429 (Mục 1.1), đối với vấn đề "${userMsg}", bạn cần lập biên bản CAPA (Biểu mẫu XN-BM 5.8.1.01), ghi nhận nguyên nhân gốc rễ và phân công nhân sự khắc phục trong vòng 5 ngày làm việc.`
        }
      ]);
      setIsThinking(false);
    }, 600);
  };

  return (
    <div
      className="w-80 text-slate-100 flex flex-col h-full border-l border-teal-900/50 print:hidden font-ui relative overflow-hidden"
      style={{ background: 'linear-gradient(180deg, #0B2A30 0%, #0E3A41 55%, #0B2A30 100%)' }}
    >
      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500&display=swap');
        .font-ui { font-family: 'Inter', ui-sans-serif, system-ui, sans-serif; }

        @keyframes riseIn {
          from { opacity: 0; transform: translateY(-6px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-riseIn { animation: riseIn 0.45s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes bubbleIn {
          from { opacity: 0; transform: translateY(8px) scale(0.98); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .animate-bubbleIn { animation: bubbleIn 0.32s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes glowPulse {
          0%, 100% { opacity: 0.5; transform: scale(1); }
          50% { opacity: 0.9; transform: scale(1.12); }
        }
        .animate-glowPulse { animation: glowPulse 2.4s ease-in-out infinite; }

        @keyframes dotBounce {
          0%, 80%, 100% { transform: translateY(0); opacity: 0.5; }
          40% { transform: translateY(-3px); opacity: 1; }
        }
        .animate-dotBounce { animation: dotBounce 1.1s ease-in-out infinite; }

        @media (prefers-reduced-motion: reduce) {
          .animate-riseIn, .animate-bubbleIn, .animate-glowPulse, .animate-dotBounce { animation: none !important; }
        }
      `}</style>

      {/* Ambient glow */}
      <div className="pointer-events-none absolute -top-16 -right-16 w-56 h-56 rounded-full bg-teal-500/10 blur-3xl" />

      {/* Tra cứu nhanh */}
      <div className="relative p-3.5 border-b border-teal-900/40 animate-riseIn">
        <label className="text-[10.5px] font-bold text-teal-300/70 uppercase tracking-wider mb-2 block">
          Tra cứu tài liệu 2429
        </label>
        <div className="relative">
          <Search className="w-4 h-4 text-teal-400/60 absolute left-2.5 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            placeholder="Tìm mã SOP, biểu mẫu, tên..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full bg-white/5 text-xs pl-8 pr-3 py-2 rounded-lg text-slate-100 placeholder-slate-500 border border-white/10 focus:outline-none focus:ring-4 focus:ring-teal-500/15 focus:border-teal-400/50 transition-all duration-200"
          />
        </div>
      </div>

      {/* Header trợ lý AI */}
      <div className="relative p-3.5 border-b border-teal-900/40 flex items-center gap-2.5 bg-white/[0.03]">
        <div className="relative flex items-center justify-center w-8 h-8 shrink-0">
          <span className="absolute inset-0 rounded-full bg-teal-400/20 animate-glowPulse" />
          <span className="relative flex items-center justify-center w-8 h-8 rounded-full bg-teal-500/15 border border-teal-400/30">
            <Bot className="w-4 h-4 text-teal-300" />
          </span>
        </div>
        <span className="font-semibold text-xs text-teal-50 flex items-center gap-1.5">
          Trợ lý AI Chất Lượng <Sparkles className="w-3.5 h-3.5 text-amber-400" />
        </span>
      </div>

      {/* Khung chat */}
      <div className="relative flex-1 p-3 overflow-y-auto space-y-2.5 text-xs">
        {messages.map((m, idx) => (
          <div
            key={idx}
            className={`p-2.5 rounded-xl leading-relaxed animate-bubbleIn ${
              m.role === 'user'
                ? 'bg-teal-600 text-white ml-6 rounded-br-md shadow-sm shadow-teal-950/30'
                : 'bg-white/[0.06] text-slate-300 mr-6 border border-white/10 rounded-bl-md'
            }`}
          >
            {m.text}
          </div>
        ))}

        {isThinking && (
          <div className="mr-6 bg-white/[0.06] border border-white/10 rounded-xl rounded-bl-md px-3.5 py-3 flex items-center gap-1 w-fit animate-bubbleIn">
            <span className="w-1.5 h-1.5 rounded-full bg-teal-300 animate-dotBounce" style={{ animationDelay: '0ms' }} />
            <span className="w-1.5 h-1.5 rounded-full bg-teal-300 animate-dotBounce" style={{ animationDelay: '150ms' }} />
            <span className="w-1.5 h-1.5 rounded-full bg-teal-300 animate-dotBounce" style={{ animationDelay: '300ms' }} />
          </div>
        )}
      </div>

      {/* Ô nhập câu hỏi */}
      <div className="relative p-3 border-t border-teal-900/40 flex gap-2 bg-white/[0.02]">
        <input
          type="text"
          placeholder="Hỏi về chuẩn 2429, CAPA..."
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && handleAISend()}
          className="flex-1 bg-white/5 text-xs px-3 py-2 rounded-lg text-slate-100 placeholder-slate-500 border border-white/10 focus:outline-none focus:ring-4 focus:ring-teal-500/15 focus:border-teal-400/50 transition-all duration-200"
        />
        <button
          onClick={handleAISend}
          className="p-2 bg-teal-600 hover:bg-teal-500 rounded-lg text-white transition-all duration-200 hover:-translate-y-px active:translate-y-0 shadow-sm shadow-teal-950/30 cursor-pointer"
        >
          <Send className="w-3.5 h-3.5" />
        </button>
      </div>
    </div>
  );
};
