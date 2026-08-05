// components/AIAssistantPanel.tsx
import React, { useState } from 'react';
import { Bot, Search, Sparkles, Send } from 'lucide-react';

export const AIAssistantPanel: React.FC = () => {
  const [searchTerm, setSearchTerm] = useState('');
  const [query, setQuery] = useState('');
  const [messages, setMessages] = useState<Array<{ role: 'user' | 'ai'; text: string }>>([
    { role: 'ai', text: 'Xin chào! Tôi là Trợ lý AI Quản lý Chất lượng 2429. Bạn cần hỗ trợ gì về các tiêu chí hoặc hành động khắc phục (CAPA)?' }
  ]);

  const handleAISend = () => {
    if (!query.trim()) return;
    const userMsg = query;
    setMessages(prev => [...prev, { role: 'user', text: userMsg }]);
    setQuery('');

    // Simulated AI CAPA & Gap Analysis response logic
    setTimeout(() => {
      setMessages(prev => [
        ...prev,
        {
          role: 'ai',
          text: `Theo Tiêu chí QĐ 2429 (Mục 1.1), đối với vấn đề "${userMsg}", bạn cần lập biên bản CAPA (Biểu mẫu XN-BM 5.8.1.01), ghi nhận nguyên nhân gốc rễ và phân công nhân sự khắc phục trong vòng 5 ngày làm việc.`
        }
      ]);
    }, 600);
  };

  return (
    <div className="w-80 bg-slate-900 text-slate-100 flex flex-col h-full border-l border-slate-800 print:hidden">
      {/* Fast Search Engine */}
      <div className="p-3 border-b border-slate-800">
        <label className="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2 block">Tra cứu tài liệu 2429</label>
        <div className="relative">
          <Search className="w-4 h-4 text-slate-400 absolute left-2.5 top-2.5" />
          <input
            type="text"
            placeholder="Tìm mã SOP, biểu mẫu, tên..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full bg-slate-800 text-xs pl-8 pr-3 py-2 rounded text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
          />
        </div>
      </div>

      {/* AI Assistant Chat UI */}
      <div className="p-3 border-b border-slate-800 flex items-center gap-2 bg-slate-800/50">
        <Bot className="w-5 h-5 text-blue-400" />
        <span className="font-semibold text-xs text-blue-100 flex items-center gap-1">
          Trợ lý AI Chất Lượng <Sparkles className="w-3 h-3 text-amber-400" />
        </span>
      </div>

      <div className="flex-1 p-3 overflow-y-auto space-y-3 text-xs">
        {messages.map((m, idx) => (
          <div key={idx} className={`p-2.5 rounded-lg ${m.role === 'user' ? 'bg-blue-600 text-white ml-4' : 'bg-slate-800 text-slate-300 mr-4 border border-slate-700'}`}>
            {m.text}
          </div>
        ))}
      </div>

      {/* AI Input */}
      <div className="p-3 border-t border-slate-800 flex gap-2">
        <input
          type="text"
          placeholder="Hỏi về chuẩn 2429, CAPA..."
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && handleAISend()}
          className="flex-1 bg-slate-800 text-xs px-3 py-2 rounded text-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-500"
        />
        <button onClick={handleAISend} className="p-2 bg-blue-600 hover:bg-blue-500 rounded text-white">
          <Send className="w-3.5 h-3.5" />
        </button>
      </div>
    </div>
  );
};