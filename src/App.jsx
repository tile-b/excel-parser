import { useState } from 'react';
import LtlParser from './components/LtlParser';
import NerutiranoParser from './components/NerutiranoParser';

function App() {
  const [activeParser, setActiveParser] = useState('LTL');

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 font-sans">
      
      {/* Navigation / Switcher */}
      <div className="mb-8 flex bg-white rounded-xl shadow-sm p-1 border border-slate-200">
        <button
          onClick={() => setActiveParser('LTL')}
          className={`
            px-6 py-2 rounded-lg font-semibold text-sm transition-all cursor-pointer
            ${activeParser === 'LTL' 
              ? 'bg-emerald-500 text-white shadow-md' 
              : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'}
          `}
        >
          LTL Izveštaj
        </button>
        <button
          onClick={() => setActiveParser('NERUTIRANO')}
          className={`
            px-6 py-2 rounded-lg font-semibold text-sm transition-all cursor-pointer
            ${activeParser === 'NERUTIRANO' 
              ? 'bg-rose-500 text-white shadow-md' 
              : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'}
          `}
        >
          NERUTIRANO
        </button>
      </div>

      {/* Active Parser Component */}
      {activeParser === 'LTL' ? <LtlParser /> : <NerutiranoParser />}

    </div>
  );
}

export default App;
