
import React, { useState, useEffect, useMemo } from 'react';
import { Download, FileText, Upload, Users, Calendar, CheckCircle, AlertCircle, Palette, Package, Library, Trash2 } from 'lucide-react';
import { DateRange, Person, TableConfig, TemplateField, AppMode } from './types';
import { DEFAULT_TABLE_CONFIG, REQUIRED_TEMPLATES } from './constants';
import { processNames, calculateDates, formatDate, extractKursId } from './utils/csvUtils';
import { generateFinalDocx } from './utils/wordUtils';

const DB_NAME = 'DocxMasterDB';
const STORE_NAME = 'templates';

const initDB = (): Promise<IDBDatabase> => new Promise((resolve, reject) => {
  const request = indexedDB.open(DB_NAME, 1);
  request.onupgradeneeded = () => request.result.createObjectStore(STORE_NAME);
  request.onsuccess = () => resolve(request.result);
  request.onerror = () => reject(request.error);
});

const saveTemplateToDB = async (id: string, buffer: ArrayBuffer) => {
  const db = await initDB();
  const tx = db.transaction(STORE_NAME, 'readwrite');
  tx.objectStore(STORE_NAME).put(buffer, id);
  return new Promise((resolve) => tx.oncomplete = resolve);
};

const getTemplateFromDB = async (id: string): Promise<ArrayBuffer | null> => {
  const db = await initDB();
  return new Promise((resolve) => {
    const request = db.transaction(STORE_NAME).objectStore(STORE_NAME).get(id);
    request.onsuccess = () => resolve(request.result || null);
  });
};

const App: React.FC = () => {
  const [appMode, setAppMode] = useState<AppMode>(AppMode.ATTENDANCE);
  const [csvData, setCsvData] = useState<Person[]>([]);
  const [storedTemplates, setStoredTemplates] = useState<Record<string, ArrayBuffer>>({});
  const [dateRanges, setDateRanges] = useState<DateRange[]>([{ id: crypto.randomUUID(), from: '', to: '' }]);
  const [tableConfig, setTableConfig] = useState<TableConfig>(DEFAULT_TABLE_CONFIG);
  const [templateFields, setTemplateFields] = useState<TemplateField[]>([]);
  const [status, setStatus] = useState<{ type: 'success' | 'error' | null, message: string }>({ type: null, message: '' });
  const [isGenerating, setIsGenerating] = useState(false);
  const [libStatus, setLibStatus] = useState({ jszip: false, pizzip: false, docxtemplater: false, papa: false });

  const allLibsReady = useMemo(() => Object.values(libStatus).every(v => v), [libStatus]);
  const allDates = useMemo(() => calculateDates(dateRanges), [dateRanges]);
  const kursId = useMemo(() => csvData.length > 0 ? extractKursId(csvData[0].kurs) : "", [csvData]);

  const getFieldsForBuffer = async (buffer: ArrayBuffer): Promise<TemplateField[]> => {
    if (!(window as any).JSZip) return [];
    try {
      const zip = new (window as any).JSZip();
      const content = await zip.loadAsync(buffer);
      const docXml = await content.file("word/document.xml")?.async("text");
      if (!docXml) return [];

      const regex = /\{\{([^}]+)\}\}/g;
      const matches = Array.from(docXml.matchAll(regex));
      const uniqueNames = Array.from(new Set(matches.map(m => m[1].replace(/<[^>]+>/g, '').trim()))).filter(n => n !== 'maintable' && n !== '');
      
      return uniqueNames.map(name => {
        let val = "";
        const ln = name.toLowerCase();
        if (csvData.length > 0) {
          const p = csvData[0];
          if (ln === 'vorname') val = p.firstName;
          else if (ln === 'nachname') val = p.lastName;
          else if (ln === 'titel') val = p.title;
          else if (ln.includes('komma')) val = p.titleSuffix ? `, ${p.titleSuffix}` : "";
          else if (ln.includes('nachgestellt')) val = p.titleSuffix;
          else if (ln.includes('kursbezeichnung')) val = p.kursbezeichnung;
          else if (ln.includes('kurs')) val = p.kurs;
          else if (ln.includes('einrichtung')) val = p.einrichtung;
        }
        if (allDates.length > 0) {
          if (ln === 'kursbeginn') val = formatDate(allDates[0]);
          if (ln === 'kursende') val = formatDate(allDates[allDates.length - 1]);
        }
        return { name, value: val };
      });
    } catch (e) { return []; }
  };

  const extractPlaceholders = async (buffer: ArrayBuffer) => {
    const fields = await getFieldsForBuffer(buffer);
    setTemplateFields(fields);
  };

  useEffect(() => {
    const check = setInterval(() => {
      setLibStatus({ jszip: !!(window as any).JSZip, pizzip: !!(window as any).PizZip, docxtemplater: !!(window as any).docxtemplater, papa: !!(window as any).Papa });
    }, 1000);
    const load = async () => {
      const loaded: Record<string, ArrayBuffer> = {};
      for (const t of REQUIRED_TEMPLATES) {
        const buf = await getTemplateFromDB(t.id);
        if (buf) loaded[t.id] = buf;
      }
      setStoredTemplates(loaded);
    };
    load();
    return () => clearInterval(check);
  }, []);

  useEffect(() => {
    const tId = appMode.toLowerCase();
    if (storedTemplates[tId]) extractPlaceholders(storedTemplates[tId]);
    else setTemplateFields([]);
  }, [appMode, storedTemplates, csvData, allDates]);

  const handleCsvUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !libStatus.papa) return;
    (window as any).Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: (results: any) => {
        const processed = processNames(results.data);
        setCsvData(processed);
        setStatus({ type: 'success', message: `${processed.length} Personen geladen.` });
      }
    });
  };

  const handleTemplateUpload = async (id: string, e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    await saveTemplateToDB(id, buf);
    setStoredTemplates(prev => ({ ...prev, [id]: buf }));
    setStatus({ type: 'success', message: `Template für ${id} gespeichert.` });
  };

  const generate = async (mode: AppMode) => {
    const tId = mode.toLowerCase();
    if (!storedTemplates[tId]) return setStatus({ type: 'error', message: `Template fehlt.` });
    setIsGenerating(true);
    try {
      const blob = await generateFinalDocx(storedTemplates[tId], csvData, dateRanges, templateFields, tableConfig, true, mode);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const suffix = kursId ? `_${kursId}` : "";
      link.href = url; link.download = `${mode}${suffix}.docx`; link.click();
      setStatus({ type: 'success', message: `Export erfolgreich!` });
    } catch (err: any) { setStatus({ type: 'error', message: err.message }); }
    finally { setIsGenerating(false); }
  };

  const generateAll = async () => {
    if (csvData.length === 0 || !libStatus.jszip) return;
    setIsGenerating(true);
    const JSZip = (window as any).JSZip;
    const zip = new JSZip();
    const suffix = kursId ? `_${kursId}` : "";

    try {
      const modesList = Object.keys(AppMode);
      for (const t of REQUIRED_TEMPLATES) {
        const modeKey = modesList.find(key => key.toLowerCase() === t.id.toLowerCase());
        if (!modeKey) continue;
        const mode = modeKey as AppMode;
        
        const buf = storedTemplates[t.id];
        if (buf) {
          const dynamicFields = await getFieldsForBuffer(buf);
          const blob = await generateFinalDocx(buf, csvData, dateRanges, dynamicFields, tableConfig, true, mode);
          zip.file(`${t.label}${suffix}.docx`, blob);
        }
      }
      const finalZip = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(finalZip);
      const link = document.createElement('a');
      link.href = url; link.download = `Gesamtpaket${suffix}.zip`; link.click();
      setStatus({ type: 'success', message: `Paket erfolgreich erstellt!` });
    } catch (err: any) { setStatus({ type: 'error', message: err.message }); }
    finally { setIsGenerating(false); }
  };

  return (
    <div className="max-w-7xl mx-auto p-4 md:p-8 space-y-8 pb-20">
      <header className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-2xl shadow-sm border border-gray-100">
        <div>
          <h1 className="text-2xl font-black text-gray-900 tracking-tight">Docx Master <span className="text-blue-600">Pro</span></h1>
          <p className="text-gray-500 text-sm font-medium italic">Status: {allLibsReady ? 'Betriebsbereit' : 'Bibliotheken werden geladen...'}</p>
        </div>
        <div className="flex flex-wrap bg-gray-100 p-1 rounded-xl gap-1">
          {Object.keys(AppMode).map((m) => (
            <button key={m} onClick={() => setAppMode(m as AppMode)} className={`px-3 py-2 rounded-lg transition-all text-[10px] font-black uppercase ${appMode === m ? 'bg-white shadow-sm text-blue-600' : 'text-gray-500 hover:text-gray-700'}`}>
              {m.replace('NAMETAGS', 'Schilder').replace('ATTENDANCE', 'Anwesenheit').replace('ABSENCES', 'Absenzen').replace('CONTACTS', 'Kontakte').replace('SOL', 'SOL')}
            </button>
          ))}
        </div>
      </header>

      {status.type && (
        <div className={`flex items-center gap-4 p-5 rounded-2xl border ${status.type === 'success' ? 'bg-green-50 border-green-200 text-green-800' : 'bg-red-50 border-red-200 text-red-800'}`}>
          {status.type === 'success' ? <CheckCircle className="w-5 h-5" /> : <AlertCircle className="w-5 h-5" />}
          <span className="text-sm font-bold">{status.message}</span>
        </div>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
        <div className="lg:col-span-3 space-y-8">
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 space-y-6">
            <h2 className="text-lg font-bold flex items-center gap-3"><Library className="w-5 h-5 text-purple-500" /> Vorlagen</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {REQUIRED_TEMPLATES.map(t => (
                <div key={t.id} className={`p-3 rounded-xl border-2 ${storedTemplates[t.id] ? 'border-green-100 bg-green-50/20' : 'border-dashed border-gray-200 bg-gray-50'}`}>
                   <div className="flex justify-between items-center gap-2">
                    <div className="overflow-hidden">
                      <p className="text-[9px] font-black uppercase text-gray-400 truncate">{t.label}</p>
                      <p className="text-xs font-bold text-gray-800">{storedTemplates[t.id] ? '✓ Bereit' : 'Fehlt'}</p>
                    </div>
                    <div className="relative w-8 h-8 bg-white rounded border flex items-center justify-center shadow-sm">
                      <Upload className="w-4 h-4 text-gray-400" />
                      <input type="file" accept=".docx" onChange={(e) => handleTemplateUpload(t.id, e)} className="absolute inset-0 opacity-0 cursor-pointer" />
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </section>

          <section className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 space-y-6">
            <h2 className="text-lg font-bold flex items-center gap-3"><Users className="w-5 h-5 text-blue-500" /> Teilnehmer & Felder</h2>
            <div className={`border-2 border-dashed rounded-2xl p-8 relative ${csvData.length > 0 ? 'border-blue-400 bg-blue-50/10' : 'border-gray-200'}`}>
              <input type="file" accept=".csv" onChange={handleCsvUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
              <div className="text-center">
                <FileText className="w-8 h-8 mx-auto mb-2 text-gray-300" />
                <p className="text-sm font-bold">{csvData.length > 0 ? `${csvData.length} Teilnehmer geladen` : 'CSV laden'}</p>
              </div>
            </div>
            {templateFields.length > 0 && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-gray-50 p-4 rounded-xl border">
                {templateFields.map(f => (
                  <div key={f.name}>
                    <label className="text-[9px] font-black text-gray-400 uppercase">{f.name}</label>
                    <input type="text" value={f.value} onChange={(e) => setTemplateFields(templateFields.map(tf => tf.name === f.name ? {...tf, value: e.target.value} : tf))} className="w-full p-2 border rounded-lg text-sm" />
                  </div>
                ))}
              </div>
            )}
          </section>

          <section className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-lg font-bold flex items-center gap-3"><Calendar className="w-5 h-5 text-green-500" /> Kurs-Tage</h2>
              <button onClick={() => setDateRanges([...dateRanges, {id: crypto.randomUUID(), from: '', to: ''}])} className="text-[10px] bg-blue-600 text-white px-3 py-1.5 rounded-lg font-black">+ NEUER BEREICH</button>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {dateRanges.map((range) => (
                <div key={range.id} className="flex gap-2 items-center bg-gray-50 p-3 rounded-xl border">
                  <input type="date" value={range.from} onChange={(e) => setDateRanges(dateRanges.map(r => r.id === range.id ? {...r, from: e.target.value} : r))} className="p-1 border rounded text-xs" />
                  <span className="text-gray-300">bis</span>
                  <input type="date" value={range.to} onChange={(e) => setDateRanges(dateRanges.map(r => r.id === range.id ? {...r, to: e.target.value} : r))} className="p-1 border rounded text-xs" />
                  <button onClick={() => setDateRanges(dateRanges.filter(r => r.id !== range.id))}><Trash2 className="w-4 h-4 text-red-400" /></button>
                </div>
              ))}
            </div>
          </section>
        </div>

        <div className="space-y-6">
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 space-y-5">
            <h2 className="text-sm font-black uppercase tracking-widest flex items-center gap-2"><Palette className="w-4 h-4 text-blue-500" /> Design</h2>
            <div className="grid grid-cols-2 gap-2">
              <input type="color" value={tableConfig.headerBgColor} onChange={(e) => setTableConfig({...tableConfig, headerBgColor: e.target.value})} className="w-full h-8 rounded" />
              <input type="color" value={tableConfig.headerTextColor} onChange={(e) => setTableConfig({...tableConfig, headerTextColor: e.target.value})} className="w-full h-8 rounded" />
            </div>
          </section>
          <button onClick={() => generate(appMode)} disabled={isGenerating || csvData.length === 0} className="w-full py-4 bg-blue-600 text-white rounded-2xl font-black shadow-lg hover:bg-blue-700 disabled:opacity-50">
            {appMode.replace('NAMETAGS','SCHILDER')} EXPORTIEREN
          </button>
          <button onClick={generateAll} disabled={isGenerating || csvData.length === 0} className="w-full py-6 bg-gradient-to-br from-indigo-600 to-purple-700 text-white rounded-2xl font-black shadow-xl hover:from-indigo-700 hover:to-purple-800 transition-all group">
            <Package className="w-8 h-8 mx-auto mb-1 group-hover:scale-110 transition-transform" /> GESAMTPAKET
          </button>
        </div>
      </div>
    </div>
  );
};

export default App;
