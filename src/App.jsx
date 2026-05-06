import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Download, FilePlus, ChevronLeft, ChevronRight, Calendar, LayoutGrid, RotateCcw, Trash2 } from 'lucide-react';

const DAY_NAMES = ['LUN', 'MAR', 'MER', 'JEU', 'VEN'];
const DAY_NAMES_FULL = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi'];
const MONTHS_FR = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];

function parseYMD(str) {
  const [y, m, d] = str.split('-').map(Number);
  return new Date(y, m - 1, d);
}
function formatYMD(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}
function getMonday(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = d.getDay();
  d.setDate(d.getDate() - day + (day === 0 ? -6 : 1));
  return d;
}
function getWeekDays(monday) {
  return Array.from({ length: 5 }, (_, i) => {
    const d = new Date(monday);
    d.setDate(monday.getDate() + i);
    return d;
  });
}
function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}
function formatDayShort(date) {
  return String(date.getDate()).padStart(2, '0') + '.' + String(date.getMonth() + 1).padStart(2, '0');
}
const todayKey = formatYMD(new Date());

export default function App() {
  const [data, setData] = useState(() => {
    try { return JSON.parse(localStorage.getItem('pointage-data') || '{}'); }
    catch { return {}; }
  });
  const [currentMonday, setCurrentMonday] = useState(() => getMonday(new Date()));
  const [fileName, setFileName] = useState(
    () => localStorage.getItem('pointage-filename') || 'pointage_alternance.xlsx'
  );
  const [statusMessage, setStatusMessage] = useState('');
  const [statusKind, setStatusKind] = useState('info');
  const [view, setView] = useState('week');
  const [workStart, setWorkStart] = useState(() => Number(localStorage.getItem('pointage-wstart') ?? 8));
  const [workEnd, setWorkEnd]     = useState(() => Number(localStorage.getItem('pointage-wend')   ?? 17));
  const [currentMonth, setCurrentMonth] = useState(() => {
    const n = new Date();
    return { year: n.getFullYear(), month: n.getMonth() };
  });
  const fileInputRef  = useRef(null);
  const historyRef    = useRef([]);
  const [canUndo, setCanUndo] = useState(false);

  useEffect(() => { localStorage.setItem('pointage-data', JSON.stringify(data)); }, [data]);
  useEffect(() => { localStorage.setItem('pointage-filename', fileName); }, [fileName]);
  useEffect(() => { localStorage.setItem('pointage-wstart', workStart); }, [workStart]);
  useEffect(() => { localStorage.setItem('pointage-wend', workEnd); }, [workEnd]);

  const workHours = useMemo(
    () => Array.from({ length: Math.max(0, workEnd - workStart + 1) }, (_, i) => workStart + i),
    [workStart, workEnd]
  );
  const weekDays   = useMemo(() => getWeekDays(currentMonday), [currentMonday]);
  const weekNum    = getWeekNumber(currentMonday);
  const totalSlots = workHours.length * 5;

  const allActivities = useMemo(() => {
    const counts = {};
    Object.values(data).forEach(day =>
      Object.values(day).forEach(act => {
        const v = (act || '').trim();
        if (v) counts[v] = (counts[v] || 0) + 1;
      })
    );
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([k]) => k);
  }, [data]);

  const filledCount = useMemo(() =>
    weekDays.reduce((acc, day) => {
      const k = formatYMD(day);
      return acc + workHours.filter(h => (data[k]?.[h] || '').trim()).length;
    }, 0),
    [data, weekDays, workHours]
  );

  const monthWeeks = useMemo(() => {
    const { year, month } = currentMonth;
    const lastDay = new Date(year, month + 1, 0);
    let cur = getMonday(new Date(year, month, 1));
    const weeks = [];
    while (cur <= lastDay) {
      const days = getWeekDays(cur);
      if (days.some(d => d.getMonth() === month)) weeks.push({ monday: new Date(cur), days });
      cur = new Date(cur);
      cur.setDate(cur.getDate() + 7);
    }
    return weeks;
  }, [currentMonth]);

  function flash(msg, kind = 'info') {
    setStatusMessage(msg);
    setStatusKind(kind);
    setTimeout(() => setStatusMessage(''), 4500);
  }

  function pushHistory(snapshot) {
    historyRef.current = [...historyRef.current.slice(-49), snapshot];
    setCanUndo(true);
  }

  function setCell(dayKey, hour, value) {
    setData(prev => {
      pushHistory(prev);
      const next = { ...prev, [dayKey]: { ...(prev[dayKey] || {}) } };
      if (value?.trim()) next[dayKey][hour] = value;
      else {
        delete next[dayKey][hour];
        if (!Object.keys(next[dayKey]).length) delete next[dayKey];
      }
      return next;
    });
  }

  function handleUndo() {
    if (!historyRef.current.length) return;
    const prev = historyRef.current.pop();
    setData(prev);
    setCanUndo(historyRef.current.length > 0);
    flash('Action annulée', 'info');
  }

  function handleClearWeek() {
    const keys = weekDays.map(formatYMD);
    const hasData = keys.some(k => data[k]);
    if (!hasData) return flash('Semaine déjà vide', 'info');
    if (!window.confirm('Effacer toute la semaine en cours ?')) return;
    pushHistory(data);
    setData(prev => {
      const next = { ...prev };
      keys.forEach(k => delete next[k]);
      return next;
    });
    flash('Semaine effacée', 'success');
  }

  function handleClearAll() {
    if (!Object.keys(data).length) return flash('Aucune donnée à effacer', 'info');
    if (!window.confirm('Effacer toutes les données ? Cette action est annulable via Undo.')) return;
    pushHistory(data);
    setData({});
    flash('Toutes les données effacées', 'success');
  }

  function navigateWeek(delta) {
    const next = new Date(currentMonday);
    next.setDate(currentMonday.getDate() + delta * 7);
    setCurrentMonday(getMonday(next));
  }

  function navigateMonth(delta) {
    setCurrentMonth(prev => {
      let m = prev.month + delta, y = prev.year;
      if (m < 0)  { m = 11; y--; }
      if (m > 11) { m = 0;  y++; }
      return { year: y, month: m };
    });
  }

  async function handleUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      const newData = {};
      let lastWeekFound = null;
      wb.SheetNames.forEach(sheetName => {
        const match = sheetName.match(/^S\d+[-_](\d{4}-\d{2}-\d{2})$/);
        if (!match) return;
        const monday = parseYMD(match[1]);
        if (!lastWeekFound || monday > lastWeekFound) lastWeekFound = monday;
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
        const days = getWeekDays(monday).map(formatYMD);
        for (let r = 1; r < rows.length; r++) {
          const row = rows[r];
          if (!row?.[0]) continue;
          const hourMatch = String(row[0]).match(/^(\d{1,2})/);
          if (!hourMatch) continue;
          const hour = parseInt(hourMatch[1], 10);
          if (hour < 0 || hour > 23) continue;
          for (let c = 0; c < 5; c++) {
            const val = row[c + 1];
            if (val != null && String(val).trim()) {
              if (!newData[days[c]]) newData[days[c]] = {};
              newData[days[c]][hour] = String(val).trim();
            }
          }
        }
      });
      setData(newData);
      setFileName(file.name);
      if (lastWeekFound) setCurrentMonday(getMonday(lastWeekFound));
      const n = wb.SheetNames.filter(s => /^S\d+[-_]\d{4}-\d{2}-\d{2}$/.test(s)).length;
      flash(`${file.name} importé · ${n} semaine${n > 1 ? 's' : ''} reconnue${n > 1 ? 's' : ''}`, 'success');
    } catch (err) {
      flash(`Erreur de lecture : ${err.message}`, 'error');
    }
    event.target.value = '';
  }

  function handleDownload() {
    const wb = XLSX.utils.book_new();
    const byWeek = {};
    Object.entries(data).forEach(([dayKey, hours]) => {
      const wk = formatYMD(getMonday(parseYMD(dayKey)));
      if (!byWeek[wk]) byWeek[wk] = {};
      byWeek[wk][dayKey] = hours;
    });
    const currentKey = formatYMD(currentMonday);
    if (!byWeek[currentKey]) byWeek[currentKey] = {};

    Object.keys(byWeek).sort().forEach(wk => {
      const monday = parseYMD(wk);
      const days = getWeekDays(monday);
      const wn = getWeekNumber(monday);
      const sheetName = `S${String(wn).padStart(2, '0')}-${wk}`;
      const aoa = [['Heure', ...days.map((d, i) => `${DAY_NAMES_FULL[i]} ${formatDayShort(d)}`)]];
      workHours.forEach(hour => {
        const row = [`${String(hour).padStart(2, '0')}h-${String(hour + 1).padStart(2, '0')}h`];
        days.forEach(d => row.push(data[formatYMD(d)]?.[hour] || ''));
        aoa.push(row);
      });
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [{ wch: 12 }, { wch: 32 }, { wch: 32 }, { wch: 32 }, { wch: 32 }, { wch: 32 }];
      XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
    });
    XLSX.writeFile(wb, fileName);
    flash(`${fileName} téléchargé`, 'success');
  }

  function handleNew() {
    if (Object.keys(data).length > 0 && !window.confirm('Créer un nouveau fichier ? Les données non téléchargées seront perdues.')) return;
    const def = `pointage_${formatYMD(new Date())}.xlsx`;
    const newName = window.prompt('Nom du nouveau fichier :', def);
    if (!newName) return;
    setData({});
    const final = newName.endsWith('.xlsx') ? newName : newName + '.xlsx';
    setFileName(final);
    flash(`Nouveau classeur · ${final}`, 'success');
  }

  function getDayStats(dateKey) {
    const dayData = data[dateKey] || {};
    const filled = workHours.filter(h => (dayData[h] || '').trim()).length;
    const acts = [...new Set(Object.values(dayData).filter(v => v?.trim()))];
    return { filled, acts, pct: workHours.length ? (filled / workHours.length) * 100 : 0 };
  }

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
        .pointage-app * { box-sizing: border-box; }
        .pointage-app {
          font-family: 'JetBrains Mono', ui-monospace, monospace;
          background: #ece3d0; color: #1a1715; min-height: 100vh;
          padding: 2.2rem 1.5rem 3rem;
          background-image:
            radial-gradient(ellipse 80% 60% at 15% 10%, rgba(161,58,42,0.06), transparent 60%),
            radial-gradient(ellipse 80% 60% at 90% 90%, rgba(60,80,100,0.06), transparent 60%),
            repeating-linear-gradient(0deg, rgba(0,0,0,0.012) 0 1px, transparent 1px 4px);
        }
        .container { max-width: 1240px; margin: 0 auto; position: relative; }

        .top-banner {
          display: flex; align-items: flex-end; justify-content: space-between;
          padding-bottom: 1.4rem; border-bottom: 2px solid #1a1715; gap: 2rem; position: relative;
        }
        .double-rule { height: 1px; background: #1a1715; margin: 4px 0 1.6rem; }
        .title-block { display: flex; flex-direction: column; min-width: 0; }
        .title-eyebrow { font-size: 0.7rem; letter-spacing: 0.32em; text-transform: uppercase; color: #a13a2a; font-weight: 600; }
        .title {
          font-family: 'Instrument Serif', 'Times New Roman', serif; font-weight: 400;
          font-size: clamp(2.8rem, 7vw, 5.2rem); line-height: 0.92; margin: 0.18em 0 0; letter-spacing: -0.025em;
        }
        .title em { font-style: italic; color: #a13a2a; font-weight: 400; margin-left: 0.05em; }
        .week-block { text-align: right; flex-shrink: 0; }
        .week-num { font-family: 'Instrument Serif', serif; font-style: italic; font-size: 4.2rem; line-height: 0.9; color: #a13a2a; letter-spacing: -0.02em; }
        .week-num span { color: #1a1715; font-style: normal; font-size: 0.6em; }
        .week-dates { font-size: 0.78rem; letter-spacing: 0.05em; margin-top: 0.4rem; opacity: 0.85; }
        .week-nav { display: flex; gap: 0.35rem; margin-top: 0.7rem; justify-content: flex-end; }
        .nav-btn {
          background: transparent; border: 1px solid #1a1715; padding: 0.4rem 0.7rem;
          font-family: inherit; font-size: 0.66rem; cursor: pointer; color: #1a1715;
          transition: all 0.12s; letter-spacing: 0.12em; text-transform: uppercase;
          display: inline-flex; align-items: center; gap: 0.3rem;
        }
        .nav-btn:hover { background: #1a1715; color: #ece3d0; }

        .stamp {
          position: absolute; top: -0.5rem; right: 38%; width: 92px; height: 92px;
          border: 2.5px solid #a13a2a; border-radius: 50%;
          display: flex; align-items: center; justify-content: center; flex-direction: column;
          color: #a13a2a; transform: rotate(-9deg); opacity: 0.55; pointer-events: none;
        }
        .stamp::before { content: ''; position: absolute; inset: 5px; border: 1px solid #a13a2a; border-radius: 50%; }
        .stamp .a { font-size: 0.55rem; letter-spacing: 0.2em; }
        .stamp .b { font-family: 'Instrument Serif', serif; font-style: italic; font-size: 1.1rem; line-height: 1; margin: 0.15em 0; }
        .stamp .c { font-size: 0.55rem; letter-spacing: 0.18em; }

        .toolbar {
          display: flex; gap: 0.5rem; flex-wrap: wrap; align-items: center;
          padding: 0.9rem 0; border-bottom: 1px dashed rgba(26,23,21,0.4); margin-bottom: 1.2rem;
        }
        .tool-btn {
          background: transparent; border: 1px solid #1a1715; padding: 0.55rem 0.95rem;
          font-family: inherit; font-size: 0.72rem; cursor: pointer; color: #1a1715;
          transition: all 0.12s; letter-spacing: 0.12em; text-transform: uppercase;
          display: inline-flex; align-items: center; gap: 0.5rem; font-weight: 500;
        }
        .tool-btn:hover { background: #1a1715; color: #ece3d0; }
        .tool-btn.primary { background: #a13a2a; border-color: #a13a2a; color: #f5ecd9; }
        .tool-btn.primary:hover { background: #1a1715; border-color: #1a1715; }
        .tool-btn.danger { border-color: #a13a2a; color: #a13a2a; }
        .tool-btn.danger:hover { background: #a13a2a; color: #f5ecd9; }
        .tool-btn:disabled { opacity: 0.35; cursor: not-allowed; }
        .tool-btn:disabled:hover { background: transparent; color: #1a1715; }
        .toolbar-sep { width: 1px; height: 28px; background: rgba(26,23,21,0.2); margin: 0 0.2rem; }
        .hour-config { display: flex; align-items: center; gap: 0.4rem; font-size: 0.7rem; letter-spacing: 0.1em; color: rgba(26,23,21,0.65); }
        .hour-config select {
          background: transparent; border: 1px solid rgba(26,23,21,0.35);
          font-family: inherit; font-size: 0.7rem; color: #1a1715; padding: 0.28rem 0.4rem; cursor: pointer;
        }
        .file-info { margin-left: auto; font-size: 0.7rem; opacity: 0.65; letter-spacing: 0.04em; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

        .view-toggle { display: flex; margin-bottom: 1.2rem; }
        .view-btn {
          background: transparent; border: 1px solid #1a1715; padding: 0.42rem 0.85rem;
          font-family: inherit; font-size: 0.7rem; cursor: pointer; color: #1a1715;
          letter-spacing: 0.15em; text-transform: uppercase;
          display: inline-flex; align-items: center; gap: 0.4rem; transition: all 0.12s;
        }
        .view-btn:first-child { border-right: none; }
        .view-btn.active { background: #1a1715; color: #ece3d0; }
        .view-btn:not(.active):hover { background: rgba(26,23,21,0.06); }

        .grid-wrap { overflow-x: auto; }
        .grid-table {
          border: 2px solid #1a1715; background: #faf3e0; width: 100%;
          border-collapse: collapse; box-shadow: 5px 5px 0 #1a1715; min-width: 720px;
        }
        .grid-table th, .grid-table td { border: 1px solid #1a1715; padding: 0; vertical-align: top; }
        .grid-table th {
          background: #1a1715; color: #ece3d0; font-family: inherit; font-weight: 500;
          font-size: 0.72rem; letter-spacing: 0.18em; padding: 0.7rem; text-align: left;
        }
        .grid-table th.today { background: #a13a2a; }
        .grid-table th .day-num { font-family: 'Instrument Serif', serif; font-style: italic; font-size: 1.5rem; margin-left: 0.4rem; opacity: 0.95; }
        .hour-cell {
          background: #1a1715 !important; color: #ece3d0; font-size: 0.72rem;
          padding: 0.5rem 0.7rem !important; text-align: center; letter-spacing: 0.1em;
          width: 92px; font-variant-numeric: tabular-nums;
        }
        .activity-cell { position: relative; height: 52px; }
        .activity-cell.today-col { background: rgba(161,58,42,0.04); }
        .activity-cell input {
          width: 100%; height: 100%; border: none; background: transparent;
          padding: 0.6rem 0.75rem; font-family: inherit; font-size: 0.84rem; color: #1a1715; outline: none;
        }
        .activity-cell input::placeholder { color: rgba(26,23,21,0.25); }
        .activity-cell input:focus { background: #fff8e0; box-shadow: inset 0 0 0 2px #a13a2a; }
        .activity-cell.filled input { background: rgba(26,23,21,0.05); }
        .activity-cell.filled input:focus { background: #fff8e0; }

        .month-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 1.2rem; }
        .month-title { font-family: 'Instrument Serif', serif; font-style: italic; font-size: 2.4rem; color: #a13a2a; letter-spacing: -0.02em; line-height: 1; }
        .month-title span { font-style: normal; color: #1a1715; font-size: 0.52em; margin-left: 0.35em; opacity: 0.7; }
        .month-nav { display: flex; gap: 0.35rem; }
        .month-grid { border: 2px solid #1a1715; background: #faf3e0; border-collapse: collapse; width: 100%; box-shadow: 5px 5px 0 #1a1715; }
        .month-grid th { background: #1a1715; color: #ece3d0; font-family: inherit; font-size: 0.7rem; letter-spacing: 0.2em; padding: 0.6rem 0.8rem; border: 1px solid #1a1715; text-align: left; }
        .month-day-cell { border: 1px solid #1a1715; padding: 0.65rem 0.8rem; vertical-align: top; width: 20%; cursor: pointer; transition: background 0.1s; height: 96px; }
        .month-day-cell:hover { background: #fff8e0; }
        .month-day-cell.outside { opacity: 0.28; cursor: default; pointer-events: none; }
        .month-day-cell.is-today { background: rgba(161,58,42,0.06); }
        .month-day-wn { font-size: 0.57rem; letter-spacing: 0.1em; color: rgba(26,23,21,0.35); margin-bottom: 0.25rem; }
        .month-day-num { font-family: 'Instrument Serif', serif; font-style: italic; font-size: 1.55rem; line-height: 1; color: #1a1715; }
        .month-day-cell.is-today .month-day-num { color: #a13a2a; }
        .month-fill-bar { height: 3px; background: rgba(26,23,21,0.1); margin: 0.35rem 0 0.2rem; overflow: hidden; }
        .month-fill-inner { height: 100%; background: #a13a2a; }
        .month-slot-count { font-size: 0.6rem; color: rgba(26,23,21,0.45); }
        .month-acts { font-size: 0.6rem; color: rgba(26,23,21,0.55); overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; margin-top: 0.15rem; line-height: 1.45; }

        .stats-bar {
          display: flex; align-items: center; gap: 1rem; margin-top: 1.6rem; padding-top: 1rem;
          border-top: 1px solid #1a1715; font-size: 0.72rem; letter-spacing: 0.08em; flex-wrap: wrap;
        }
        .progress-meta { display: flex; align-items: center; gap: 0.7rem; flex: 1; min-width: 220px; }
        .count-num { font-family: 'Instrument Serif', serif; font-style: italic; font-size: 1.7rem; color: #a13a2a; line-height: 1; }
        .count-total { font-size: 0.85rem; opacity: 0.6; }
        .progress-bar { flex: 1; max-width: 320px; height: 10px; border: 1px solid #1a1715; background: rgba(26,23,21,0.04); }
        .progress-fill { height: 100%; background: #a13a2a; transition: width 0.3s; }
        .status-msg { font-size: 0.7rem; min-height: 1em; font-style: italic; letter-spacing: 0.04em; text-align: right; }
        .status-msg.success { color: #2d6a3e; }
        .status-msg.error   { color: #a13a2a; }
        .status-msg.info    { color: rgba(26,23,21,0.5); }

        @media (max-width: 760px) {
          .pointage-app { padding: 1rem 0.6rem 2rem; }
          .top-banner { flex-direction: column; align-items: flex-start; gap: 1rem; }
          .week-block { text-align: left; width: 100%; }
          .week-num { font-size: 3.4rem; }
          .week-nav { justify-content: flex-start; }
          .stamp { display: none; }
          .grid-table { box-shadow: 3px 3px 0 #1a1715; }
          .file-info { width: 100%; margin: 0.4rem 0 0; }
          .tool-btn { font-size: 0.65rem; padding: 0.45rem 0.7rem; }
          .month-day-cell { padding: 0.4rem 0.45rem; height: 72px; }
          .month-day-num { font-size: 1.15rem; }
          .month-acts { display: none; }
          .toolbar-sep { display: none; }
        }
      `}</style>

      <div className="pointage-app">
        <div className="container">

          <div className="top-banner">
            <div className="title-block">
              <span className="title-eyebrow">Carnet de bord — Alternance</span>
              <h1 className="title">Pointage<em>.</em></h1>
            </div>
            <div className="stamp" aria-hidden="true">
              <div className="a">SEMAINE</div>
              <div className="b">№{String(weekNum).padStart(2, '0')}</div>
              <div className="c">{currentMonday.getFullYear()}</div>
            </div>
            <div className="week-block">
              <div className="week-num"><span>S</span>{String(weekNum).padStart(2, '0')}</div>
              <div className="week-dates">
                {weekDays[0].toLocaleDateString('fr-FR', { day: 'numeric', month: 'long' })}
                {' → '}
                {weekDays[4].toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })}
              </div>
              <div className="week-nav">
                <button className="nav-btn" onClick={() => { navigateWeek(-1); setView('week'); }}>
                  <ChevronLeft size={12} /> Préc.
                </button>
                <button className="nav-btn" onClick={() => { setCurrentMonday(getMonday(new Date())); setView('week'); }}>
                  Auj.
                </button>
                <button className="nav-btn" onClick={() => { navigateWeek(1); setView('week'); }}>
                  Suiv. <ChevronRight size={12} />
                </button>
              </div>
            </div>
          </div>
          <div className="double-rule" />

          <div className="toolbar">
            <button className="tool-btn" onClick={() => fileInputRef.current?.click()}>
              <Upload size={14} /> Importer .xlsx
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleUpload} style={{ display: 'none' }} />
            <button className="tool-btn" onClick={handleNew}>
              <FilePlus size={14} /> Nouveau
            </button>
            <button className="tool-btn primary" onClick={handleDownload}>
              <Download size={14} /> Télécharger
            </button>

            <div className="toolbar-sep" />

            <button className="tool-btn" onClick={handleUndo} disabled={!canUndo} title="Annuler la dernière action">
              <RotateCcw size={14} /> Undo
            </button>
            <button className="tool-btn" onClick={handleClearWeek}>
              <Trash2 size={14} /> Semaine
            </button>
            <button className="tool-btn danger" onClick={handleClearAll}>
              <Trash2 size={14} /> Tout
            </button>

            <div className="toolbar-sep" />

            <div className="hour-config">
              DE
              <select value={workStart} onChange={e => {
                const v = Number(e.target.value);
                setWorkStart(v);
                if (v >= workEnd) setWorkEnd(v + 1);
              }}>
                {Array.from({ length: 23 }, (_, i) => (
                  <option key={i} value={i}>{String(i).padStart(2, '0')}h</option>
                ))}
              </select>
              À
              <select value={workEnd} onChange={e => {
                const v = Number(e.target.value);
                setWorkEnd(v);
                if (v <= workStart) setWorkStart(v - 1);
              }}>
                {Array.from({ length: 23 }, (_, i) => i + 1).map(i => (
                  <option key={i} value={i}>{String(i).padStart(2, '0')}h</option>
                ))}
              </select>
            </div>

            <span className="file-info">→ {fileName}</span>
          </div>

          <div className="view-toggle">
            <button className={`view-btn ${view === 'week' ? 'active' : ''}`} onClick={() => setView('week')}>
              <LayoutGrid size={13} /> Semaine
            </button>
            <button className={`view-btn ${view === 'month' ? 'active' : ''}`} onClick={() => setView('month')}>
              <Calendar size={13} /> Mois
            </button>
          </div>

          <datalist id="activity-suggestions">
            {allActivities.map(act => <option key={act} value={act} />)}
          </datalist>

          {view === 'week' ? (
            <div className="grid-wrap">
              <table className="grid-table">
                <thead>
                  <tr>
                    <th className="hour-cell">Heure</th>
                    {weekDays.map((d, i) => {
                      const isToday = formatYMD(d) === todayKey;
                      return (
                        <th key={i} className={isToday ? 'today' : ''}>
                          {DAY_NAMES[i]}
                          <span className="day-num">{String(d.getDate()).padStart(2, '0')}</span>
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {workHours.map(h => (
                    <tr key={h}>
                      <td className="hour-cell">
                        {String(h).padStart(2, '0')}—{String(h + 1).padStart(2, '0')}
                      </td>
                      {weekDays.map((d, i) => {
                        const dayKey = formatYMD(d);
                        const value = data[dayKey]?.[h] || '';
                        const isToday = dayKey === todayKey;
                        return (
                          <td key={i} className={`activity-cell${value ? ' filled' : ''}${isToday ? ' today-col' : ''}`}>
                            <input
                              type="text"
                              list="activity-suggestions"
                              value={value}
                              onChange={e => setCell(dayKey, h, e.target.value)}
                              placeholder="—"
                            />
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <>
              <div className="month-header">
                <div className="month-title">
                  {MONTHS_FR[currentMonth.month]}
                  <span>{currentMonth.year}</span>
                </div>
                <div className="month-nav">
                  <button className="nav-btn" onClick={() => navigateMonth(-1)}><ChevronLeft size={12} /> Préc.</button>
                  <button className="nav-btn" onClick={() => setCurrentMonth({ year: new Date().getFullYear(), month: new Date().getMonth() })}>Auj.</button>
                  <button className="nav-btn" onClick={() => navigateMonth(1)}>Suiv. <ChevronRight size={12} /></button>
                </div>
              </div>
              <table className="month-grid">
                <thead>
                  <tr>{DAY_NAMES.map(d => <th key={d}>{d}</th>)}</tr>
                </thead>
                <tbody>
                  {monthWeeks.map(({ monday, days }, wi) => (
                    <tr key={wi}>
                      {days.map((d, di) => {
                        const dateKey = formatYMD(d);
                        const inMonth = d.getMonth() === currentMonth.month;
                        const isToday = dateKey === todayKey;
                        const { filled, acts, pct } = getDayStats(dateKey);
                        return (
                          <td
                            key={di}
                            className={`month-day-cell${!inMonth ? ' outside' : ''}${isToday ? ' is-today' : ''}`}
                            onClick={() => { setCurrentMonday(getMonday(d)); setView('week'); }}
                          >
                            {di === 0 && <div className="month-day-wn">S{String(getWeekNumber(monday)).padStart(2, '0')}</div>}
                            <div className="month-day-num">{d.getDate()}</div>
                            {inMonth && filled > 0 && (
                              <>
                                <div className="month-fill-bar">
                                  <div className="month-fill-inner" style={{ width: `${pct}%` }} />
                                </div>
                                <div className="month-slot-count">{filled}/{workHours.length}h</div>
                                {acts.length > 0 && <div className="month-acts">{acts.join(' · ')}</div>}
                              </>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          )}

          <div className="stats-bar">
            <div className="progress-meta">
              <span className="count-num">{filledCount}</span>
              <span className="count-total">/ {totalSlots} créneaux</span>
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: totalSlots ? `${(filledCount / totalSlots) * 100}%` : '0%' }} />
              </div>
            </div>
            <div className={`status-msg ${statusKind}`}>
              {statusMessage || `${allActivities.length} activité${allActivities.length !== 1 ? 's' : ''} en mémoire pour autocomplete`}
            </div>
          </div>

        </div>
      </div>
    </>
  );
}
