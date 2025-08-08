
import React from 'react';
import * as XLSX from 'xlsx';
import ReactDOM from 'react-dom';

const colorClasses = { green: 'text-green-600', yellow: 'text-yellow-600', blue: 'text-blue-600' };

const normalizeNumberInput = (value) => {
  if (value === null || value === undefined) return '';
  const stringValue = String(value);
  const valueWithDot = stringValue.replace(',', '.');
  if (valueWithDot === '') return '';
  if (!/^\d*\.?\d*$/.test(valueWithDot)) return null;
  return valueWithDot;
};

const calculateGrade = (score, maxScore, passingPercentage, settings) => {
  const { minGrade, passingGrade, maxGrade } = settings;
  if (isNaN(score) || isNaN(maxScore) || isNaN(passingPercentage) || maxScore <= 0 ||
      minGrade >= passingGrade || passingGrade >= maxGrade) return NaN;
  if (score > maxScore) score = maxScore;
  const passingScore = maxScore * (passingPercentage / 100);
  let finalGrade;
  if (score <= passingScore) {
    finalGrade = (passingScore === 0)
      ? passingGrade
      : minGrade + (score * ((passingGrade - minGrade) / passingScore));
  } else {
    const scoreRange = maxScore - passingScore;
    finalGrade = (scoreRange <= 0)
      ? maxGrade
      : passingGrade + ((score - passingScore) * ((maxGrade - passingGrade) / scoreRange));
  }
  return finalGrade;
};

const generateId = () => `id_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

const saveStateToLocalStorage = (state) => {
  try { localStorage.setItem('gradeCalculatorState', JSON.stringify(state)); }
  catch (e) { console.error("Error saving state to localStorage:", e); }
};

const loadStateFromLocalStorage = () => {
  try {
    const stateString = localStorage.getItem('gradeCalculatorState');
    if (stateString === null) return null;
    const parsed = JSON.parse(stateString);
    if (!parsed.sheets || !Array.isArray(parsed.sheets)) return null;
    return parsed;
  } catch (e) {
    console.error('Error loading state from localStorage:', e);
    return null;
  }
};

const AddEvalDropdown = React.memo(({ onAdd, onClose }) => {
  const [name, setName] = React.useState('Nueva Evaluación');
  const [maxScore, setMaxScore] = React.useState('10');
  const [exigency, setExigency] = React.useState('60');
  const dropdownRef = React.useRef(null);

  React.useEffect(() => {
    const handleClickOutside = (event) => { if (dropdownRef.current && !dropdownRef.current.contains(event.target)) onClose(); };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [onClose]);

  const handleSubmit = () => {
    const normMax = parseFloat(normalizeNumberInput(maxScore));
    const normExi = parseFloat(normalizeNumberInput(exigency));
    if (name && normMax > 0 && normExi >= 0 && normExi <= 100) { onAdd({ name, maxScore: normMax, exigency: normExi }); onClose(); }
  };

  return (
    <div ref={dropdownRef} className="absolute top-full right-0 mt-2 w-64 bg-white rounded-md shadow-lg z-30 p-4 border">
      <div className="space-y-3">
        <div>
          <label className="block text-sm font-medium text-gray-700">Nombre</label>
          <input type="text" value={name} onChange={e => setName(e.target.value)} className="mt-1 block w-full px-3 py-2 border rounded-md"/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Puntaje Máximo</label>
          <input type="text" value={maxScore} onChange={e => { const n = normalizeNumberInput(e.target.value); if (n !== null) setMaxScore(n); }} className="mt-1 block w-full px-3 py-2 border rounded-md"/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Exigencia (%)</label>
          <input type="text" value={exigency} onChange={e => { const n = normalizeNumberInput(e.target.value); if (n !== null) setExigency(n); }} className="mt-1 block w-full px-3 py-2 border rounded-md"/>
        </div>
        <div className="flex justify-end">
          <button onClick={handleSubmit} className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 text-sm">Añadir</button>
        </div>
      </div>
    </div>
  );
});

const DiffScoreDropdown = ({ evaluation, onSave, onClose }) => {
  const [evalData, setEvalData] = React.useState(evaluation);
  const dropdownRef = React.useRef(null);
  const [position, setPosition] = React.useState({ top: 0, left: 0 });

  React.useLayoutEffect(() => {
    const trigger = document.querySelector(`[data-dropdown-id="${evaluation.id}"]`);
    if (trigger) {
      const rect = trigger.getBoundingClientRect();
      setPosition({ top: rect.bottom + window.scrollY, left: rect.left + window.scrollX - 256 + rect.width });
    }
  }, [evaluation.id]);

  React.useEffect(() => {
    const handleClickOutside = (event) => { if (dropdownRef.current && !dropdownRef.current.contains(event.target)) onClose(); };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [onClose]);

  const handleToggle = () => setEvalData(prev => ({ ...prev, differentiatedScores: { ...prev.differentiatedScores, enabled: !prev.differentiatedScores.enabled } }));
  const handleChange = (color, value) => {
    const normalized = normalizeNumberInput(value);
    if (normalized === null) return;
    setEvalData(prev => ({ ...prev, differentiatedScores: { ...prev.differentiatedScores, [color]: normalized } }));
  };

  const dropdownContent = (
    <div ref={dropdownRef} className="absolute bg-white rounded-md shadow-lg z-[9999] p-4 border w-64" style={{ top: `${position.top}px`, left: `${position.left}px` }}>
      <div className="space-y-3">
        <div className="flex items-center">
          <input type="checkbox" checked={evalData.differentiatedScores.enabled} onChange={handleToggle} className="h-4 w-4 text-blue-600 border-gray-300 rounded"/>
          <label className="ml-2 block text-sm text-gray-900">Habilitar diferenciados</label>
        </div>
        {evalData.differentiatedScores.enabled && (
          <div className="space-y-2">
            {['green', 'yellow', 'blue'].map(color => (
              <div key={color}>
                <label className="block text-xs font-medium text-gray-700">P. Máx. <span className={`font-bold ${colorClasses[color]}`}>{color.charAt(0).toUpperCase() + color.slice(1)}</span></label>
                <input type="text" value={evalData.differentiatedScores[color]} onChange={e => handleChange(color, e.target.value)} className="mt-1 block w-full px-2 py-1 border rounded-md text-sm"/>
              </div>
            ))}
          </div>
        )}
        <div className="flex justify-end pt-2">
          <button onClick={() => { onSave(evalData); onClose(); }} className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 text-sm">Guardar</button>
        </div>
      </div>
    </div>
  );
  return ReactDOM.createPortal(dropdownContent, document.getElementById('portal-root'));
};

const ImportConfirmationModal = ({ onConfirmReplace, onConfirmAdd, onCancel }) => {
  const modalRef = React.useRef(null);
  React.useEffect(() => {
    const handleClickOutside = (event) => { if (modalRef.current && !modalRef.current.contains(event.target)) onCancel(); };
    const onEsc = (e) => { if (e.key === 'Escape') onCancel(); };
    document.addEventListener("mousedown", handleClickOutside);
    document.addEventListener("keydown", onEsc);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
      document.removeEventListener("keydown", onEsc);
    };
  }, [onCancel]);
  const modalContent = (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[10000]">
      <div ref={modalRef} className="bg-white p-6 rounded-lg shadow-xl">
        <h3 className="text-lg font-bold mb-4">Importar Hojas</h3>
        <p className="mb-6">¿Cómo quieres importar las hojas del archivo?</p>
        <div className="flex justify-end space-x-4">
          <button onClick={onCancel} className="px-4 py-2 bg-gray-300 rounded-md hover:bg-gray-400">Cancelar</button>
          <button onClick={onConfirmAdd} className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700">Añadir a las existentes</button>
          <button onClick={onConfirmReplace} className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700">Reemplazar todo</button>
        </div>
      </div>
    </div>
  );
  return ReactDOM.createPortal(modalContent, document.getElementById('portal-root'));
};

const SettingsPopover = ({ anchorId, open, onClose, children }) => {
  const popRef = React.useRef(null);
  const [pos, setPos] = React.useState({ top: 0, left: 0 });

  React.useLayoutEffect(() => {
    if (!open) return;
    const trigger = document.getElementById(anchorId);
    if (trigger) {
      const rect = trigger.getBoundingClientRect();
      setPos({ top: rect.bottom + window.scrollY + 8, left: rect.left + window.scrollX });
    }
  }, [anchorId, open]);

  React.useEffect(() => {
    if (!open) return;
    const onDoc = (e) => { if (popRef.current && !popRef.current.contains(e.target)) onClose(); };
    const onEsc = (e) => e.key === 'Escape' && onClose();
    document.addEventListener('mousedown', onDoc);
    document.addEventListener('keydown', onEsc);
    return () => {
      document.removeEventListener('mousedown', onDoc);
      document.removeEventListener('keydown', onEsc);
    };
  }, [open, onClose]);

  if (!open) return null;

  const content = (
    <div ref={popRef} className="absolute z-[10001] bg-white border rounded-lg shadow-xl w-[320px] p-4"
         style={{ top: pos.top, left: pos.left }} role="dialog" aria-modal="true">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-gray-700">Configuración de Escala de Notas</h3>
        <button onClick={onClose} aria-label="Cerrar" className="text-gray-500 hover:text-red-600">&times;</button>
      </div>
      {children}
    </div>
  );
  return ReactDOM.createPortal(content, document.getElementById('portal-root'));
};

const Header = React.memo(({ activeSheet, onAddEvaluation, onAddStudent, onFileImport, onExport, sheetsCount, isSaving, lastSaved, onToggleSettings }) => {
  const [isAddEvalDropdownOpen, setAddEvalDropdownOpen] = React.useState(false);
  const fileInputRef = React.useRef(null);
  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (file) onFileImport(file);
    event.target.value = '';
  };
  return (
    <header className="flex flex-wrap justify-between items-center mb-6 pb-4 border-b-2 border-gray-200">
      <div className="flex items-center gap-2">
        <h1 className="text-3xl font-bold text-gray-800">Calculadora de Notas</h1>
        <button id="settings-btn" onClick={onToggleSettings}
                className="ml-1 px-3 py-1.5 text-sm bg-gray-200 hover:bg-gray-300 rounded-md flex items-center gap-1"
                aria-haspopup="dialog" aria-expanded="false">
          <span>Configuración</span> <span>⚙️</span>
        </button>
        <span className={`ml-2 text-sm ${isSaving ? 'text-yellow-600 animate-pulse' : 'text-green-600'}`}>
          {isSaving ? 'Guardando…' : `✔ Guardado${lastSaved ? ' ' + lastSaved.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'}) : ''}`}
        </span>
      </div>
      <div className="flex items-center space-x-2 mt-4 sm:mt-0">
        <div className="relative">
          <button aria-label="Añadir nueva evaluación" onClick={() => setAddEvalDropdownOpen(prev => !prev)} disabled={!activeSheet} className="px-4 py-2 bg-green-500 text-white rounded-md shadow-sm hover:bg-green-600 disabled:bg-gray-300">Añadir Evaluación</button>
          {isAddEvalDropdownOpen && <AddEvalDropdown onAdd={onAddEvaluation} onClose={() => setAddEvalDropdownOpen(false)} />}
        </div>
        <button aria-label="Añadir nuevo alumno" onClick={onAddStudent} disabled={!activeSheet} className="px-4 py-2 bg-blue-500 text-white rounded-md shadow-sm hover:bg-blue-600 disabled:bg-gray-300">Añadir Alumno</button>
        <button aria-label="Importar desde archivo Excel" onClick={() => fileInputRef.current.click()} className="px-4 py-2 bg-gray-700 text-white rounded-md shadow-sm hover:bg-gray-800">Importar</button>
        <input type="file" ref={fileInputRef} onChange={handleFileChange} className="hidden" accept=".xlsx, .xls" />
        <button aria-label="Exportar a archivo Excel" onClick={onExport} disabled={sheetsCount === 0} className="px-4 py-2 bg-orange-500 text-white rounded-md shadow-sm hover:bg-orange-600 disabled:bg-gray-300">Exportar</button>
      </div>
    </header>
  );
});

const GlobalSettings = React.memo(({ settings, setSettings }) => {
  const [displayValues, setDisplayValues] = React.useState({
    minGrade: settings.minGrade.toFixed(1),
    passingGrade: settings.passingGrade.toFixed(1),
    maxGrade: settings.maxGrade.toFixed(1)
  });
  React.useEffect(() => {
    setDisplayValues({
      minGrade: settings.minGrade.toFixed(1),
      passingGrade: settings.passingGrade.toFixed(1),
      maxGrade: settings.maxGrade.toFixed(1)
    });
  }, [settings]);
  const handleDisplayChange = (key, value) => {
    const normalized = normalizeNumberInput(value);
    if (normalized !== null) setDisplayValues(p => ({ ...p, [key]: normalized }));
  };
  const handleBlur = (key) => {
    let numValue = parseFloat(displayValues[key]);
    if (isNaN(numValue)) numValue = settings[key];
    setSettings(p => ({ ...p, [key]: numValue }));
  };
  const labels = { minGrade: 'Nota Mínima', passingGrade: 'Nota Aprobación', maxGrade: 'Nota Máxima' };
  return (
    <div className="space-y-3">
      {['minGrade', 'passingGrade', 'maxGrade'].map(k => (
        <div key={k}>
          <label className="block text-sm font-medium text-gray-500">{labels[k]}</label>
          <input type="text" value={displayValues[k]} onChange={e => handleDisplayChange(k, e.target.value)} onBlur={() => handleBlur(k)} className="mt-1 w-full p-2 border rounded-md"/>
        </div>
      ))}
    </div>
  );
});

const Tabs = React.memo(({ sheets, activeSheetId, onSelectTab, onDeleteTab, onRenameTab, onAddTab, onReorderSheets }) => {
  const [editingId, setEditingId] = React.useState(null);
  const [name, setName] = React.useState('');
  const [draggedIndex, setDraggedIndex] = React.useState(null);
  const [dragOverIndex, setDragOverIndex] = React.useState(null);

  const handleRename = (id) => { if (name.trim()) onRenameTab(id, name.trim()); setEditingId(null); };
  const handleDragStart = (e, index) => { setDraggedIndex(index); e.dataTransfer.effectAllowed = 'move'; };
  const handleDragOver = (e, index) => { e.preventDefault(); if (index !== dragOverIndex) setDragOverIndex(index); };
  const handleDrop = (e, dropIndex) => { e.preventDefault(); if (draggedIndex === null || draggedIndex === dropIndex) return; onReorderSheets(draggedIndex, dropIndex); };
  const handleDragEnd = () => { setDraggedIndex(null); setDragOverIndex(null); };

  return (
    <div className="flex border-b border-gray-200">
      {sheets.map((s, i) => (
        <div key={s.id}
             className={`tab flex items-center px-4 py-2 border-t border-l border-r rounded-t-lg cursor-pointer -mb-px draggable-tab ${activeSheetId === s.id ? 'active bg-white font-semibold' : 'bg-gray-200'} ${dragOverIndex === i ? 'dragging-over-tab' : ''}`}
             onClick={() => onSelectTab(s.id)}
             onDoubleClick={() => { setEditingId(s.id); setName(s.name); }}
             draggable="true"
             onDragStart={(e) => handleDragStart(e, i)}
             onDragOver={(e) => handleDragOver(e, i)}
             onDrop={(e) => handleDrop(e, i)}
             onDragEnd={handleDragEnd}
             onDragLeave={() => setDragOverIndex(null)}>
          {editingId === s.id
            ? <input type="text" value={name} onChange={e => setName(e.target.value)} onBlur={() => handleRename(s.id)} onKeyDown={e => e.key === 'Enter' && handleRename(s.id)} autoFocus className="bg-transparent"/>
            : <span>{s.name}</span>}
          {activeSheetId === s.id && (
            <button aria-label={`Eliminar hoja ${s.name}`} onClick={e => { e.stopPropagation(); onDeleteTab(s.id); }} className="ml-3 text-gray-500 hover:text-red-600">&times;</button>
          )}
        </div>
      ))}
      <button aria-label="Añadir nueva hoja" onClick={onAddTab} className="px-4 py-2 bg-gray-200 rounded-t-lg hover:bg-gray-300">+</button>
    </div>
  );
});

const StudentRow = React.memo(({ student, rowIndex, evaluations, globalSettings, handleScoreChange, handleScoreBlur, handleKeyDown, inputRefs, setEditingRow }) => {
  const rowData = React.useMemo(() => {
    let total = 0, count = 0;
    const grades = evaluations.map((ev, e) => {
      let maxScore = ev.maxScore;
      if (ev.differentiatedScores.enabled && student.highlight) {
        maxScore = parseFloat(ev.differentiatedScores[student.highlight]) || ev.maxScore;
      }
      const g = calculateGrade(parseFloat(student.scores[e]), maxScore, ev.exigency, globalSettings);
      if (!isNaN(g)) { total += g; count++; }
      return g;
    });
    const avg = count > 0 ? (total / count) : NaN;
    return { grades, avg };
  }, [student.scores, student.highlight, evaluations, globalSettings]);

  return (
    <tr className={`group h-14 ${student.highlight ? `highlight-${student.highlight}`:''}`}>
      {evaluations.map((ev, e) => (
        <React.Fragment key={ev.id}>
          <td className={`p-1 text-sm text-center ${!student.highlight && e % 2 !== 0 ? 'bg-gray-50' : ''}`}>
            <input
              type="text"
              value={student.scores[e] ?? ''}
              onChange={evt => handleScoreChange(rowIndex, e, evt.target.value)}
              onBlur={() => { handleScoreBlur(rowIndex, e); setEditingRow(null); }}
              onFocus={() => setEditingRow(rowIndex)}
              onKeyDown={evt => handleKeyDown(evt, rowIndex, e)}
              ref={el => inputRefs.current[`${rowIndex}-${e}`] = el}
              className="w-20 text-center bg-transparent focus:bg-yellow-100 p-2 rounded"
            />
          </td>
          <td className={`p-1 text-sm text-center font-bold ${!isNaN(rowData.grades[e]) && (rowData.grades[e] >= globalSettings.passingGrade ? 'grade-pass' : 'grade-fail')} ${!student.highlight && e % 2 !== 0 ? 'bg-gray-50' : ''}`}>
            {isNaN(rowData.grades[e]) ? 'N/A' : rowData.grades[e].toFixed(1)}
          </td>
        </React.Fragment>
      ))}
      <td className={`p-4 text-sm text-center font-bold bg-green-50 ${!isNaN(rowData.avg) && (rowData.avg >= globalSettings.passingGrade ? 'grade-pass' : 'grade-fail')}`}>
        {isNaN(rowData.avg) ? 'N/A' : rowData.avg.toFixed(1)}
      </td>
    </tr>
  );
});

const GradesTable = ({ sheet, onUpdateSheet, globalSettings }) => {
  const [studentRoot, setStudentRoot] = React.useState(null);
  const [gradesRoot, setGradesRoot] = React.useState(null);
  const [openDropdownEvalId, setOpenDropdownEvalId] = React.useState(null);
  const [draggedIndex, setDraggedIndex] = React.useState(null);
  const [dragOverIndex, setDragOverIndex] = React.useState(null);
  const inputRefs = React.useRef({});
  const [editingRow, setEditingRow] = React.useState(null);

  const handleSaveDiffScores = React.useCallback((updatedEval) => {
    const updatedEvals = sheet.evaluations.map(ev => ev.id === updatedEval.id ? updatedEval : ev);
    onUpdateSheet({ ...sheet, evaluations: updatedEvals });
  }, [sheet, onUpdateSheet]);

  const update = React.useCallback((data) => onUpdateSheet({ ...sheet, ...data }), [sheet, onUpdateSheet]);

  const handleScoreChange = React.useCallback((r, e, v) => {
    const normalized = normalizeNumberInput(v);
    if (normalized === null) return;
    const d = sheet.studentData.map(s => ({ ...s, scores: [...s.scores] }));
    d[r].scores[e] = normalized;
    update({ ...sheet, studentData: d });
  }, [sheet, update]);

  const handleScoreBlur = React.useCallback((r, e) => {
    const student = sheet.studentData[r];
    const evaluation = sheet.evaluations[e];
    let numericScore = parseFloat(student.scores[e]);
    if (isNaN(numericScore)) numericScore = '';
    let maxScore = evaluation.maxScore;
    if (evaluation.differentiatedScores.enabled && student.highlight) {
      maxScore = parseFloat(evaluation.differentiatedScores[student.highlight]) || evaluation.maxScore;
    }
    if (numericScore > maxScore) numericScore = maxScore;
    const d = [...sheet.studentData];
    d[r].scores[e] = numericScore;
    update({ studentData: d });
  }, [sheet.studentData, sheet.evaluations, update]);

  const handleDragStart = React.useCallback((e, index) => { setDraggedIndex(index); e.dataTransfer.effectAllowed = 'move'; }, []);
  const handleDragOver = React.useCallback((e, index) => { e.preventDefault(); if (index !== dragOverIndex) setDragOverIndex(index); }, [dragOverIndex]);
  const handleDrop = React.useCallback((e, dropIndex) => {
    e.preventDefault();
    if (draggedIndex === null || draggedIndex === dropIndex) return;
    const newEvals = [...sheet.evaluations];
    const [draggedItem] = newEvals.splice(draggedIndex, 1);
    newEvals.splice(dropIndex, 0, draggedItem);
    const newStudentData = sheet.studentData.map(student => {
      const newScores = [...student.scores];
      const [draggedScore] = newScores.splice(draggedIndex, 1);
      newScores.splice(dropIndex, 0, draggedScore);
      return { ...student, scores: newScores };
    });
    update({ evaluations: newEvals, studentData: newStudentData });
    setDraggedIndex(null);
    setDragOverIndex(null);
  }, [draggedIndex, sheet.evaluations, sheet.studentData, update]);
  const handleDragEnd = React.useCallback(() => { setDraggedIndex(null); setDragOverIndex(null); }, []);

  const handleNameChange = React.useCallback((r, v) => { const d = [...sheet.studentData]; d[r].name = v; update({ studentData: d }); }, [sheet.studentData, update]);
  const handleHighlight = React.useCallback((r) => { const h = ['green','yellow','blue',null]; const d = [...sheet.studentData]; d[r].highlight = h[(h.indexOf(d[r].highlight) + 1) % 4]; update({ studentData: d }); }, [sheet.studentData, update]);
  const handleEvalHeaderChange = React.useCallback((index, field, value) => {
    const normalized = normalizeNumberInput(value);
    if (normalized === null) return;
    const numValue = normalized === '' ? 0 : parseFloat(normalized);
    const d = [...sheet.evaluations];
    d[index][field] = numValue;
    update({ evaluations: d });
  }, [sheet.evaluations, update]);
  const handleEvalNameChange = React.useCallback((index, value) => { const d = [...sheet.evaluations]; d[index].name = value; update({ evaluations: d }); }, [sheet.evaluations, update]);
  const delEval = React.useCallback((e) => update({ evaluations: sheet.evaluations.filter((_,i)=>i!==e), studentData: sheet.studentData.map(s=>({...s, scores: s.scores.filter((_,i)=>i!==e)})) }), [sheet, update]);
  const delStudent = React.useCallback((r) => update({ studentData: sheet.studentData.filter((_,i)=>i!==r) }), [sheet.studentData, update]);

  const handleKeyDown = React.useCallback((event, rowIndex, colIndex) => {
    const { key } = event;
    let nextRow = rowIndex, nextCol = colIndex;
    switch (key) {
      case 'ArrowUp': nextRow--; break;
      case 'ArrowDown': nextRow++; break;
      case 'Enter': event.preventDefault(); nextRow++; break;
      case 'ArrowLeft': nextCol--; break;
      case 'ArrowRight': nextCol++; break;
      default: return;
    }
    event.preventDefault();
    let targetRef;
    if (['ArrowUp','ArrowDown','Enter'].includes(key) && nextRow >= 0 && nextRow < sheet.studentData.length)
      targetRef = inputRefs.current[`${nextRow}-${colIndex}`];
    else if ((key === 'ArrowLeft' || key === 'ArrowRight') && nextCol >= 0 && nextCol < sheet.evaluations.length)
      targetRef = inputRefs.current[`${rowIndex}-${nextCol}`];
    if (targetRef) { targetRef.focus(); targetRef.select(); }
  }, [sheet.studentData.length, sheet.evaluations.length]);

  const handleColumnResize = React.useCallback((e) => {
    e.preventDefault();
    const startX = e.clientX, startWidth = sheet.studentColumnWidth;
    const handleMouseMove = (moveEvent) => {
      const newWidth = startWidth + (moveEvent.clientX - startX);
      if (newWidth > 100) onUpdateSheet({ ...sheet, studentColumnWidth: newWidth });
    };
    const handleMouseUp = () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  }, [sheet, onUpdateSheet]);

  const StudentColumnPortal = studentRoot ? ReactDOM.createPortal(
    <table className="table-fixed border-separate" style={{ borderSpacing: 0 }}>
      <thead className="bg-gray-50">
      <tr className="h-16">
        <th rowSpan="3" className="p-3 w-12 text-left text-xs font-medium text-gray-500 border-b-2">#</th>
        <th style={{ width: `${sheet.studentColumnWidth}px`}} rowSpan="3" className="p-3 text-left text-xs font-medium text-gray-500 border-b-2 relative">
          {sheet.studentHeader}
          <div className="resize-handle" onMouseDown={handleColumnResize}></div>
        </th>
      </tr>
      <tr className="h-10"></tr>
      <tr className="h-10"></tr>
      </thead>
      <tbody className="bg-white divide-y divide-gray-200">
      {sheet.studentData.map((s, r) => (
        <tr key={s.id} className={`group h-14 ${s.highlight ? `highlight-${s.highlight}`:''}`}>
          <td className="p-4 text-sm relative text-center">{r+1}
            <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100 transition-opacity">
              <button aria-label={`Eliminar estudiante ${s.name}`} onClick={()=>delStudent(r)} className="p-1 bg-white rounded shadow hover:bg-red-100"><span className="text-red-500 font-bold">×</span></button>
            </div>
          </td>
          <td className="p-1 text-sm">
            <div className="flex items-center">
              <input
                type="text"
                value={s.name}
                onChange={e => handleNameChange(r, e.target.value)}
                className={`bg-transparent w-full focus:bg-yellow-100 p-2 rounded ${editingRow === r ? 'font-bold text-blue-700' : ''}`}
              />
              <button aria-label="Cambiar color de resaltado" onClick={() => handleHighlight(r)} className="ml-2 text-xl text-gray-500 hover:text-blue-600" title="Cambiar color">🎨</button>
            </div>
          </td>
        </tr>
      ))}
      </tbody>
    </table>, studentRoot) : null;

  const GradesColumnPortal = gradesRoot ? ReactDOM.createPortal(
    <table className="min-w-full border-separate" style={{ borderSpacing: 0 }}>
      <thead className="bg-gray-50">
      <tr className="h-16">
        {sheet.evaluations.map((ev, i) => (
          <th key={ev.id} colSpan="2"
              className={`p-3 text-center text-xs font-medium text-gray-500 relative group draggable-header ${i % 2 !== 0 ? 'bg-gray-100' : ''} ${dragOverIndex === i ? 'dragging-over-col' : ''}`}
              draggable="true"
              onDragStart={(e) => handleDragStart(e, i)} onDragOver={(e) => handleDragOver(e, i)} onDrop={(e) => handleDrop(e, i)} onDragEnd={handleDragEnd} onDragLeave={() => setDragOverIndex(null)}>
            <div className="flex justify-center items-center">
              <input type="text" value={ev.name} onChange={e => handleEvalNameChange(i, e.target.value)} className="font-bold text-base bg-transparent w-full text-center focus:bg-yellow-100 px-1 rounded"/>
              <div className="relative">
                <button aria-label={`Configuración para ${ev.name}`} data-dropdown-id={ev.id} onClick={() => setOpenDropdownEvalId(openDropdownEvalId === ev.id ? null : ev.id)} className="ml-2 text-lg text-gray-500 hover:text-blue-600">⚙️</button>
                {openDropdownEvalId === ev.id && <DiffScoreDropdown evaluation={ev} onSave={handleSaveDiffScores} onClose={() => setOpenDropdownEvalId(null)} />}
              </div>
              <div className="absolute top-0 right-0 opacity-0 group-hover:opacity-100 transition-opacity">
                <button aria-label={`Eliminar evaluación ${ev.name}`} onClick={() => delEval(i)} className="p-1 bg-white rounded shadow hover:bg-red-100"><span className="text-red-500 font-bold">×</span></button>
              </div>
            </div>
          </th>
        ))}
        <th rowSpan="3" className="p-3 w-28 text-center text-xs font-medium text-gray-500 bg-green-50 border-b-2">Promedio</th>
      </tr>
      <tr className="h-10">
        {sheet.evaluations.map((ev, i) =>
          <th key={ev.id} colSpan="2" className={`p-1 ${i % 2 !== 0 ? 'bg-gray-100' : ''}`}>
            <div className="flex justify-center items-center space-x-2">
              <label className="text-xs">P.Máx:</label>
              <input type="text" value={ev.maxScore} onChange={e => handleEvalHeaderChange(i, 'maxScore', e.target.value)} className="w-16 p-1 text-xs border rounded" />
              <label className="text-xs">Exig:</label>
              <input type="text" value={ev.exigency} onChange={e => handleEvalHeaderChange(i, 'exigency', e.target.value)} className="w-16 p-1 text-xs border rounded"/>
            </div>
          </th>)}
      </tr>
      <tr className="h-10">
        {sheet.evaluations.map((ev, i) => <React.Fragment key={ev.id}>
          <th className={`border-t border-b-2 p-1 w-20 text-xs font-medium text-gray-500 ${i % 2 !== 0 ? 'bg-gray-100' : ''}`}>Puntaje</th>
          <th className={`border-t border-b-2 p-1 w-20 bg-blue-50 text-xs font-medium ${i % 2 !== 0 ? 'bg-gray-100' : ''}`}>Nota</th>
        </React.Fragment>)}
      </tr>
      </thead>
      <tbody className="bg-white divide-y divide-gray-200">
      {sheet.studentData.map((s, r) => (
        <StudentRow key={s.id} student={s} rowIndex={r} evaluations={sheet.evaluations}
                    globalSettings={globalSettings} handleScoreChange={handleScoreChange}
                    handleScoreBlur={handleScoreBlur} handleKeyDown={handleKeyDown} inputRefs={inputRefs} setEditingRow={setEditingRow}/>
      ))}
      </tbody>
    </table>, gradesRoot) : null;

  return (
    <div className="bg-white p-2 sm:p-4 rounded-b-lg shadow-md flex">
      <div ref={setStudentRoot} className="relative flex-shrink-0"></div>
      <div ref={setGradesRoot} className="overflow-x-auto flex-grow"></div>
      {StudentColumnPortal}
      {GradesColumnPortal}
    </div>
  );
};

const App = () => {
  const defaultState = {
    sheets: [],
    activeSheetId: null,
    globalSettings: { minGrade: 1.0, passingGrade: 4.0, maxGrade: 7.0 }
  };
  const getInitialState = () => {
    const loaded = loadStateFromLocalStorage();
    if (loaded) {
      const sheetsWithWidth = (loaded.sheets || []).map(sheet => ({ ...sheet, studentColumnWidth: sheet.studentColumnWidth || 200 }));
      return { ...defaultState, ...loaded, sheets: sheetsWithWidth, globalSettings: { ...defaultState.globalSettings, ...(loaded.globalSettings || {}) } };
    }
    return defaultState;
  };
  const initial = getInitialState();
  const [sheets, setSheets] = React.useState(initial.sheets);
  const [activeSheetId, setActiveSheetId] = React.useState(initial.activeSheetId);
  const [globalSettings, setGlobalSettings] = React.useState(initial.globalSettings);
  const [pendingImport, setPendingImport] = React.useState(null);

  const [isSaving, setIsSaving] = React.useState(false);
  const [lastSaved, setLastSaved] = React.useState(null);
  const [settingsOpen, setSettingsOpen] = React.useState(false);

  React.useEffect(() => {
    setIsSaving(true);
    const handler = setTimeout(() => {
      saveStateToLocalStorage({ sheets, activeSheetId, globalSettings });
      setIsSaving(false);
      setLastSaved(new Date());
    }, 500);
    return () => clearTimeout(handler);
  }, [sheets, activeSheetId, globalSettings]);

  const handleUpdateSheet = React.useCallback(u => setSheets(p => p.map(s => s.id === u.id ? u : s)), []);
  const handleCreateSheet = React.useCallback(() => {
    const name = `Hoja ${sheets.length + 1}`;
    const newSheet = { id: generateId(), name, studentHeader: 'Alumnos', evaluations: [], studentData: [], studentColumnWidth: 200 };
    setSheets(p => [...p, newSheet]);
    setActiveSheetId(newSheet.id);
  }, [sheets.length]);
  const handleDeleteSheet = React.useCallback(id => {
    setSheets(p => {
      const n = p.filter(s => s.id !== id);
      if (activeSheetId === id) setActiveSheetId(n[0]?.id || null);
      return n;
    });
  }, [activeSheetId]);
  const handleRenameSheet = React.useCallback((id, name) => setSheets(p => p.map(s => s.id === id ? { ...s, name } : s)), []);
  const handleReorderSheets = React.useCallback((dragIndex, dropIndex) => {
    setSheets(prevSheets => {
      const newSheets = [...prevSheets];
      const [draggedItem] = newSheets.splice(dragIndex, 1);
      newSheets.splice(dropIndex, 0, draggedItem);
      return newSheets;
    });
  }, []);

  const activeSheet = React.useMemo(() => sheets.find(s => s.id === activeSheetId), [sheets, activeSheetId]);

  const handleAddEvaluation = React.useCallback((evalData) => {
    if (!activeSheet) return;
    const newEvaluation = {
      ...evalData,
      id: generateId(),
      differentiatedScores: { enabled: false, green: evalData.maxScore, yellow: evalData.maxScore, blue: evalData.maxScore }
    };
    const updatedSheet = {
      ...activeSheet,
      evaluations: [...activeSheet.evaluations, newEvaluation],
      studentData: activeSheet.studentData.map(s => ({ ...s, scores: [...s.scores, ''] }))
    };
    handleUpdateSheet(updatedSheet);
  }, [activeSheet, handleUpdateSheet]);

  const handleAddStudent = React.useCallback(() => {
    if (!activeSheet) return;
    const newStudent = {
      id: generateId(),
      name: `Estudiante ${activeSheet.studentData.length + 1}`,
      scores: Array(activeSheet.evaluations.length).fill(''),
      highlight: null
    };
    handleUpdateSheet({ ...activeSheet, studentData: [...activeSheet.studentData, newStudent] });
  }, [activeSheet, handleUpdateSheet]);

  const handleFileImport = React.useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const importedSheets = workbook.SheetNames.map(name => {
          const ws = workbook.Sheets[name];
          const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
          if (json.length < 5) return null;
          const studentHeader = json[4][1] || 'Alumnos';
          const highlightColIndex = json[4].findIndex(h => h === 'Highlight');
          const evaluations = [];
          for (let i = 2; i < json[0].length; i += 2) {
            if (!json[0][i]) continue;
            const maxScore = parseFloat(json[1][i] || 0);
            const exigency = parseFloat(json[2][i] || 0);
            if (isNaN(maxScore) || isNaN(exigency)) continue;
            const diffScoresRaw = String(json[3][i + 1] || '').split(',');
            evaluations.push({
              id: generateId(), name: String(json[0][i]), maxScore, exigency,
              differentiatedScores: {
                enabled: json[3][i] === 'SI',
                green: parseFloat(diffScoresRaw[0]) || maxScore,
                yellow: parseFloat(diffScoresRaw[1]) || maxScore,
                blue: parseFloat(diffScoresRaw[2]) || maxScore,
              }
            });
          }
          if (evaluations.length === 0) return null;
          const studentData = json.slice(5).map(r => {
            if (!r[1]) return null;
            return {
              id: generateId(), name: String(r[1]),
              scores: evaluations.map((_, j) => String(r[2 + j * 2] || '')),
              highlight: highlightColIndex !== -1 ? (r[highlightColIndex] || null) : null
            };
          }).filter(Boolean);
          if (studentData.length === 0) return null;
          return { id: generateId(), name, studentHeader, evaluations, studentData, studentColumnWidth: 200 };
        }).filter(Boolean);

        if (importedSheets.length === 0) {
          alert('No se encontraron datos válidos en el archivo Excel.');
          return;
        }
        setPendingImport(importedSheets);
      } catch (err) {
        console.error("Error importing file:", err);
        alert("Hubo un error al procesar el archivo. Asegúrese de que tenga el formato correcto.");
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleConfirmReplace = () => {
    setSheets(pendingImport);
    setActiveSheetId(pendingImport[0]?.id || null);
    setPendingImport(null);
  };
  const handleConfirmAdd = () => {
    setSheets(prev => [...prev, ...pendingImport]);
    if (!activeSheetId) setActiveSheetId(pendingImport[0]?.id || null);
    setPendingImport(null);
  };

  const handleExport = React.useCallback(() => {
    const wb = XLSX.utils.book_new();
    sheets.forEach(sheet => {
      const processedData = sheet.studentData.map((student, index) => {
        let total = 0, count = 0;
        const row = [index + 1, student.name];
        sheet.evaluations.forEach((ev, i) => {
          let maxScore = ev.maxScore;
          if (ev.differentiatedScores.enabled && student.highlight) {
            maxScore = ev.differentiatedScores[student.highlight] || ev.maxScore;
          }
          const numericScore = parseFloat(student.scores[i]);
          const grade = calculateGrade(numericScore, maxScore, ev.exigency, globalSettings);
          row.push(student.scores[i] ?? '', isNaN(grade) ? '' : grade.toFixed(1));
          if (!isNaN(grade)) { total += grade; count++; }
        });
        const avg = count > 0 ? (total / count) : NaN;
        row.push(isNaN(avg) ? '' : avg.toFixed(1));
        row.push(student.highlight || '');
        return row;
      });
      const aoa = [
        ['', '', ...sheet.evaluations.flatMap(ev => [ev.name, ''])],
        ['', 'P. Máx:', ...sheet.evaluations.flatMap(ev => [ev.maxScore, ''])],
        ['', 'Exig:', ...sheet.evaluations.flatMap(ev => [ev.exigency, ''])],
        ['', 'Diferenciados', ...sheet.evaluations.flatMap(ev => [ev.differentiatedScores.enabled ? 'SI' : 'NO', `${ev.differentiatedScores.green},${ev.differentiatedScores.yellow},${ev.differentiatedScores.blue}`])],
        ['#', sheet.studentHeader, ...sheet.evaluations.flatMap(() => ['Puntaje', 'Nota']), 'Promedio Final', 'Highlight'],
        ...processedData
      ];
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
    XLSX.writeFile(wb, 'calculadora_notas.xlsx');
  }, [sheets, globalSettings]);

  React.useEffect(() => {
    const onKey = (e) => {
      if (e.ctrlKey && e.key === 'Enter') { e.preventDefault(); handleAddStudent(); }
      if (e.ctrlKey && e.shiftKey && e.key === 'Enter') { e.preventDefault(); handleAddEvaluation({ name: 'Nueva Evaluación', maxScore: 10, exigency: 60 }); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [handleAddStudent, handleAddEvaluation]);

  return (
    <div className="p-4 sm:p-6 lg:p-8 bg-gray-50 min-h-screen">
      {pendingImport && (
        <ImportConfirmationModal
          onConfirmReplace={handleConfirmReplace}
          onConfirmAdd={handleConfirmAdd}
          onCancel={() => setPendingImport(null)}
        />
      )}

      <Header
        activeSheet={activeSheet}
        onAddEvaluation={handleAddEvaluation}
        onAddStudent={handleAddStudent}
        onFileImport={handleFileImport}
        onExport={handleExport}
        sheetsCount={sheets.length}
        isSaving={isSaving}
        lastSaved={lastSaved}
        onToggleSettings={() => setSettingsOpen(v => !v)}
      />

      <SettingsPopover anchorId="settings-btn" open={settingsOpen} onClose={() => setSettingsOpen(false)}>
        <GlobalSettings settings={globalSettings} setSettings={setGlobalSettings} />
      </SettingsPopover>

      {sheets.length > 0 && activeSheetId && activeSheet ? (
        <div>
          <Tabs
            sheets={sheets}
            activeSheetId={activeSheetId}
            onSelectTab={setActiveSheetId}
            onDeleteTab={handleDeleteSheet}
            onRenameTab={handleRenameSheet}
            onAddTab={handleCreateSheet}
            onReorderSheets={handleReorderSheets}
          />
          <GradesTable key={activeSheetId} sheet={activeSheet} onUpdateSheet={handleUpdateSheet} globalSettings={globalSettings} />
        </div>
      ) : (
        <div className="text-center bg-white p-10 rounded-lg shadow-md">
          <h2 className="text-2xl font-bold text-gray-700 mb-4">Comienza a Calcular Notas</h2>
          <p className="text-gray-500 mb-6">Puedes importar un archivo de Excel o crear una tabla manualmente.</p>
          <button onClick={handleCreateSheet} className="px-6 py-3 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700">
            Crear Tabla Manualmente
          </button>
        </div>
      )}
    </div>
  );
};

export default App;