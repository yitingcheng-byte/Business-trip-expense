/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, type ReactNode, type FormEvent } from 'react';
import { 
  Plus, 
  Trash2, 
  Download, 
  CheckCircle2, 
  ChevronLeft, 
  History, 
  Calculator,
  Building2,
  User,
  Calendar,
  MapPin,
  CircleDollarSign,
  Briefcase,
  ExternalLink,
  ChevronRight,
  X,
  Edit2,
  Save
} from 'lucide-react';
import { format } from 'date-fns';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

/** Utility for Tailwind class merging */
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---

type ExpenseCategory = '交通費' | '住宿費' | '膳雜費' | '交際費' | '其他費用';
type TransportMode = '飛機' | '高鐵' | '火車' | '捷運' | '計程車' | '租車' | '其他';

interface ExpenseItem {
  id: string;
  date: string;
  location: string;
  description: string;
  currency: string;
  category: ExpenseCategory;
  amount: number;
  projectCode: string;
  transportMode: TransportMode;
}

interface PrepaidItem {
  id: string;
  currency: string;
  amount: number;
}

interface ExpenseReport {
  id: string;
  submitDate: string;
  employeeName: string;
  employeeId: string;
  unit: string;
  department: string;
  startDate: string;
  endDate: string;
  items: ExpenseItem[];
  prepaidItems: PrepaidItem[];
  // Legacy fields for backward compatibility
  prepaidCurrency?: string;
  prepaidAmount?: number;
}

// --- Constants ---

const CATEGORIES: ExpenseCategory[] = ['交通費', '住宿費', '膳雜費', '交際費', '其他費用'];
const TRANSPORT_MODES: TransportMode[] = ['飛機', '高鐵', '火車', '捷運', '計程車', '租車', '其他'];
const CURRENCIES = ['TWD', 'USD', 'EUR', 'JPY', 'CNY', 'HKD', 'GBP'];

// --- Components ---

export default function App() {
  const [view, setView] = useState<'dashboard' | 'form'>('dashboard');
  const [reports, setReports] = useState<ExpenseReport[]>([]);
  const [editingReport, setEditingReport] = useState<ExpenseReport | null>(null);

  // Load from localStorage on mount
  useEffect(() => {
    const saved = localStorage.getItem('trip_expenses');
    if (saved) {
      try {
        setReports(JSON.parse(saved));
      } catch (e) {
        console.error('Failed to parse saved reports', e);
      }
    }
  }, []);

  // Save to localStorage
  const saveReports = (newReports: ExpenseReport[]) => {
    setReports(newReports);
    localStorage.setItem('trip_expenses', JSON.stringify(newReports));
  };

  const handleSubmitReport = (report: ExpenseReport) => {
    if (editingReport) {
      // Update existing
      saveReports(reports.map(r => r.id === report.id ? report : r));
    } else {
      // Create new
      saveReports([report, ...reports]);
    }
    setView('dashboard');
    setEditingReport(null);
  };

  const handleEditReport = (report: ExpenseReport) => {
    setEditingReport(report);
    setView('form');
  };

  const handleDeleteReport = (id: string) => {
    if (window.confirm('確定要刪除這筆報銷單嗎？此操作無法復原。')) {
      saveReports(reports.filter(r => r.id !== id));
    }
  };

  const handleNewReport = () => {
    setEditingReport(null);
    setView('form');
  };

  return (
    <div className="min-h-screen bg-[#FDFBF7] text-[#3D3D33] font-sans selection:bg-[#7C8A71]/20 selection:text-[#3D3D33]">
      <header className="bg-[#7C8A71] text-white sticky top-0 z-10 shadow-md">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-white/20 rounded-xl flex items-center justify-center text-white backdrop-blur-sm">
              <Calculator size={22} />
            </div>
            <div>
              <h1 className="font-bold text-lg leading-tight tracking-wide text-white">出差費用報支系統</h1>
              <p className="text-[10px] opacity-70 font-medium uppercase tracking-[0.2em]">Business Trip Expense Tracker</p>
            </div>
          </div>
          
          {view === 'dashboard' && (
            <button 
              onClick={handleNewReport}
              className="bg-white/10 hover:bg-white/20 border border-white/30 text-white px-4 py-2 rounded-lg font-medium transition-all flex items-center gap-2"
            >
              <Plus size={18} />
              <span>建立新報銷單</span>
            </button>
          )}
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-4 md:p-8">
        <AnimatePresence mode="wait">
          {view === 'dashboard' ? (
            <Dashboard 
              key="dashboard" 
              reports={reports} 
              onNew={handleNewReport} 
              onEdit={handleEditReport}
              onDelete={handleDeleteReport}
            />
          ) : (
            <ReportForm 
              key="form" 
              initialData={editingReport}
              onCancel={() => {
                setView('dashboard');
                setEditingReport(null);
              }} 
              onSubmit={handleSubmitReport} 
            />
          )}
        </AnimatePresence>
      </main>

      <footer className="max-w-7xl mx-auto px-4 py-8 border-t border-[#DCD7CC] mt-12 flex items-center justify-between text-[10px] text-[#A5A58D] uppercase tracking-widest">
        <p>© 2026 出差費用報銷系統 · 財務暨投資管理中心</p>
        <p>系統版本 2.4.0-Natural</p>
      </footer>
    </div>
  );
}

function Dashboard({ reports, onNew, onEdit, onDelete }: { 
  reports: ExpenseReport[], 
  onNew: () => void, 
  onEdit: (r: ExpenseReport) => void,
  onDelete: (id: string) => void,
  key?: string 
}) {
  const exportToExcel = async (report: ExpenseReport) => {
    try {
      const templateUrl = `${import.meta.env.BASE_URL}templates/expense_template.xlsx`;
      const response = await fetch(templateUrl);
      
      const contentType = response.headers.get('content-type');
      if (!response.ok || (contentType && contentType.includes('text/html'))) {
        throw new Error(`無法載入範本檔 (${templateUrl})，請確認系統是否存在該檔案。`);
      }
      
      const arrayBuffer = await response.arrayBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      const ws = workbook.worksheets[0];
      if (!ws) {
        const sheetInfo = workbook.worksheets ? workbook.worksheets.map(s => s.name).join(', ') : '無';
        throw new Error(`範本中找不到任何工作表。目前共有 ${workbook.worksheets?.length || 0} 張工作表：[${sheetInfo}]`);
      }

      const currencyMap: Record<string, string> = {
        'TWD': '台幣',
        'USD': '美金',
        'JPY': '日幣',
        'EUR': '歐元',
        'CNY': '人民幣',
        'HKD': '港幣',
        'GBP': '英鎊',
      };
      const getCurrencyName = (c: string) => currencyMap[c] || c;

      // 1. Fill base data (Row 13, per original layout logic mapping)
      ws.getCell('C13').value = report.employeeName;
      ws.getCell('G13').value = report.employeeId;
      ws.getCell('K13').value = report.unit;
      ws.getCell('R13').value = report.department;
      ws.getCell('Y13').value = `${report.startDate} ~ ${report.endDate}`;

      // 2. Data Rows (Start at Row 16 based on layout where 14-15 are table headers)
      const dataStartRow = 16;
      const numItems = report.items.length;

      // Pre-calculate merged arrays for items
      const itemMerges = [
        ['A', 'B'], ['C', 'F'], ['G', 'L'], ['M', 'N'],
        ['O', 'P'], ['Q', 'R'], ['S', 'T'], ['U', 'V'],
        ['W', 'X'], ['Y', 'AA'], ['AB', 'AD']
      ];

      if (numItems > 1) {
        // Insert empty rows for the extra items, pushing footer down
        ws.spliceRows(dataStartRow + 1, 0, ...Array(numItems - 1).fill([]));
      }

      report.items.forEach((item, index) => {
        const r = dataStartRow + index;
        const row = ws.getRow(r);

        // If it's a duplicated row, copy styles
        if (index > 0) {
          for (let c = 1; c <= 30; c++) {
            row.getCell(c).style = ws.getCell(dataStartRow, c).style;
          }
        }

        // Apply column merges
        itemMerges.forEach(([startCol, endCol]) => {
          try { ws.mergeCells(`${startCol}${r}:${endCol}${r}`); } catch {}
        });

        // Populate values
        row.getCell('A').value = item.date;
        row.getCell('C').value = item.location;
        row.getCell('G').value = item.description.trim() || item.category;
        row.getCell('M').value = getCurrencyName(item.currency);
        
        row.getCell('O').value = item.category === '交通費' ? item.amount : '';
        row.getCell('Q').value = item.category === '住宿費' ? item.amount : '';
        row.getCell('S').value = item.category === '膳雜費' ? item.amount : '';
        row.getCell('U').value = item.category === '交際費' ? item.amount : '';
        row.getCell('W').value = item.category === '其他費用' ? item.amount : '';
        
        row.getCell('Y').value = item.projectCode;
        row.getCell('AB').value = item.category === '交通費' ? item.transportMode : '';

        // Dynamic Row Height Calculation
        const descCell = row.getCell('G');
        const descLength = String(descCell.value).length;
        // Estimate ~15 full width chars per line. Base row height ~20.
        const lineCount = Math.max(1, Math.ceil(descLength / 15));
        row.height = Math.max(20, Math.ceil(lineCount * 18));
        
        // Enforce wrap text just in case template cell missed it
        if (!descCell.alignment) descCell.alignment = {};
        descCell.alignment = { ...descCell.alignment, wrapText: true, vertical: 'top' };
        
        const locCell = row.getCell('C');
        if (!locCell.alignment) locCell.alignment = {};
        locCell.alignment = { ...locCell.alignment, wrapText: true, vertical: 'top' };
      });

      // 3. Totals Area
      // The spacer originally at row 17, footer header at 18, so dynamic footer data row is at 19 internally (if 1 item).
      const footerDataRow = 18 + numItems; 

      const expenseTotals: Record<string, number> = {};
      report.items.forEach(item => {
        expenseTotals[item.currency] = (expenseTotals[item.currency] || 0) + item.amount;
      });

      const prepaidTotals: Record<string, number> = {};
      (report.prepaidItems || []).forEach(p => {
        prepaidTotals[p.currency] = (prepaidTotals[p.currency] || 0) + p.amount;
      });

      const allCurrencies = Array.from(new Set([
        ...Object.keys(expenseTotals),
        ...Object.keys(prepaidTotals)
      ]));

      const numCurrencies = Math.max(1, allCurrencies.length);

      if (numCurrencies > 1) {
        ws.spliceRows(footerDataRow + 1, 0, ...Array(numCurrencies - 1).fill([]));
      }

      const footerMerges = [
        ['A', 'E'], ['F', 'G'], ['H', 'J'],
        ['K', 'O'], ['P', 'Q'], ['R', 'T'],
        ['U', 'Y'], ['Z', 'AA'], ['AB', 'AD']
      ];

      (allCurrencies.length > 0 ? allCurrencies : ['TWD']).forEach((curr, i) => {
        const r = footerDataRow + i;
        const row = ws.getRow(r);

        if (i > 0) {
          for (let c = 1; c <= 30; c++) {
            row.getCell(c).style = ws.getCell(footerDataRow, c).style;
          }
        }

        footerMerges.forEach(([startCol, endCol]) => {
          try { ws.mergeCells(`${startCol}${r}:${endCol}${r}`); } catch {}
        });

        const exp = expenseTotals[curr] || 0;
        const pre = prepaidTotals[curr] || 0;
        const diff = exp - pre;

        if (expenseTotals[curr] !== undefined) {
          row.getCell('F').value = getCurrencyName(curr);
          row.getCell('H').value = exp;
        } else {
          row.getCell('F').value = '';
          row.getCell('H').value = '';
        }

        if (prepaidTotals[curr] !== undefined) {
          row.getCell('P').value = getCurrencyName(curr);
          row.getCell('R').value = pre;
        } else {
          row.getCell('P').value = '';
          row.getCell('R').value = '';
        }

        row.getCell('Z').value = getCurrencyName(curr);
        row.getCell('AB').value = diff;
      });

      // 4. Print & Page Setup overrides (Ensuring A4 formatting)
      if (ws.pageSetup) {
        ws.pageSetup.paperSize = 9; // A4
        ws.pageSetup.fitToPage = true;
        ws.pageSetup.fitToWidth = 1;
        ws.pageSetup.fitToHeight = 0; // Allows growing downward across multiple pages
        ws.pageSetup.printArea = undefined; // Clear any hardcoded print areas that might cut off extra rows
      }

      // Generate the export buffer and trigger download
      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), `出差報銷_${report.employeeName}_${report.startDate}.xlsx`);

    } catch (e) {
      console.error(e);
      alert(e instanceof Error ? e.message : '匯出失敗，請確認是否已有準備好制式範本檔。');
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="space-y-8"
    >
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <StatCard 
          title="累計報銷單" 
          value={reports.length.toString()} 
          icon={<History className="text-[#7C8A71]" />} 
          color="bg-[#F0F2EE]"
        />
        <StatCard 
          title="本月提交" 
          value={reports.filter(r => r.submitDate.startsWith(format(new Date(), 'yyyy-MM'))).length.toString()} 
          icon={<CheckCircle2 className="text-[#4F5946]" />} 
          color="bg-[#E8EDE4]"
        />
        <StatCard 
          title="最近更新" 
          value={reports.length > 0 ? reports[0].submitDate : '無'} 
          icon={<Calendar className="text-[#A5A58D]" />} 
          color="bg-[#FDFBF7]"
        />
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-[#DCD7CC] overflow-hidden">
        <div className="p-6 border-b border-[#E5E1D8] bg-[#F8F7F2] flex items-center justify-between">
          <h2 className="font-bold text-lg flex items-center gap-2 text-[#3D3D33]">
            <Briefcase size={20} className="text-[#7C8A71]" />
            報銷歷史記錄
          </h2>
        </div>
        
        {reports.length === 0 ? (
          <div className="p-16 text-center bg-white">
            <div className="w-20 h-20 bg-[#FDFBF7] rounded-full flex items-center justify-center mx-auto mb-6 text-[#DCD7CC] border border-[#E5E1D8]">
              <History size={36} />
            </div>
            <h3 className="text-[#3D3D33] font-bold text-xl mb-2">尚未有報銷記錄</h3>
            <p className="text-[#A5A58D] mb-8 max-w-sm mx-auto text-sm leading-relaxed">
              提交您的第一筆出差費用報銷單，我們將為您自動計算並生成對應的 Excel 報表。
            </p>
            <button 
              onClick={onNew}
              className="bg-[#7C8A71] text-white px-8 py-3 rounded-lg font-bold hover:bg-[#6A7661] transition-all shadow-lg shadow-[#7C8A71]/20 flex items-center gap-2 mx-auto"
            >
              <Plus size={20} /> 建立首張單據
            </button>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-[#FDFBF7] text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest border-b border-[#E5E1D8]">
                  <th className="px-6 py-4">提交日期</th>
                  <th className="px-6 py-4">出差人 / 部門</th>
                  <th className="px-6 py-4">期間</th>
                  <th className="px-6 py-4 text-right">費用累計 (幣別)</th>
                  <th className="px-6 py-4 text-center">操作</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#F0EFEC]">
                {reports.map((report) => {
                  const totalsByCurrency = report.items.reduce((acc, item) => {
                    acc[item.currency] = (acc[item.currency] || 0) + item.amount;
                    return acc;
                  }, {} as Record<string, number>);

                  return (
                    <tr key={report.id} className="hover:bg-[#FDFCF8] transition-colors group">
                      <td className="px-6 py-4 whitespace-nowrap">
                        <span className="text-sm font-medium text-[#3D3D33]">{report.submitDate}</span>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex flex-col">
                          <span className="text-sm font-bold text-[#3D3D33]">{report.employeeName}</span>
                          <span className="text-xs text-[#A5A58D]">{report.unit} / {report.department}</span>
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-2 text-xs text-[#3D3D33] bg-[#F0F2EE] px-2 py-1 rounded w-fit border border-[#DCD7CC]/50">
                          <span>{report.startDate}</span>
                          <span className="text-[#A5A58D]">~</span>
                          <span>{report.endDate}</span>
                        </div>
                      </td>
                      <td className="px-6 py-4 text-right">
                        <div className="flex flex-col gap-1">
                          {Object.entries(totalsByCurrency).map(([curr, amt]) => (
                            <span key={curr} className="text-sm font-mono font-bold text-[#4F5946]">
                              {curr} {amt.toLocaleString()}
                            </span>
                          ))}
                        </div>
                      </td>
                      <td className="px-6 py-4 text-center">
                        <div className="flex items-center justify-center gap-1">
                          <button 
                            onClick={() => onEdit(report)}
                            className="p-2 text-[#A5A58D] hover:text-[#7C8A71] hover:bg-[#F0F2EE] rounded-lg transition-all"
                            title="查看 / 編輯"
                          >
                            <ExternalLink size={18} />
                          </button>
                          <button 
                            onClick={() => exportToExcel(report)}
                            className="p-2 text-[#A5A58D] hover:text-[#7C8A71] hover:bg-[#F0F2EE] rounded-lg transition-all"
                            title="匯出 Excel"
                          >
                            <Download size={18} />
                          </button>
                          <button 
                            onClick={() => onDelete(report.id)}
                            className="p-2 text-[#DCD7CC] hover:text-red-500 hover:bg-red-50 rounded-lg transition-all"
                            title="刪除記錄"
                          >
                            <Trash2 size={18} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </motion.div>
  );
}

function StatCard({ title, value, icon, color }: { title: string, value: string, icon: ReactNode, color: string }) {
  return (
    <div className="bg-white p-6 rounded-xl border border-[#DCD7CC] shadow-sm flex items-start justify-between group hover:border-[#7C8A71] transition-all">
      <div className="space-y-2">
        <p className="text-[10px] font-bold text-[#A5A58D] uppercase tracking-widest">{title}</p>
        <p className="text-3xl font-light text-[#3D3D33] tracking-tight">{value}</p>
      </div>
      <div className={cn("w-12 h-12 rounded-xl flex items-center justify-center transition-transform group-hover:rotate-6", color)}>
        {icon}
      </div>
    </div>
  );
}

function ExpenseItemModal({ isOpen, item, onClose, onSave }: { isOpen: boolean, item: ExpenseItem | null, onClose: () => void, onSave: (item: ExpenseItem, keepOpen: boolean) => void }) {
  const [localItem, setLocalItem] = useState<ExpenseItem | null>(null);

  useEffect(() => {
    if (item) setLocalItem(item);
  }, [item]);

  if (!isOpen || !localItem) return null;

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-[#3D3D33]/40 backdrop-blur-sm"
      />
      <motion.div 
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-full max-w-lg bg-white rounded-2xl shadow-2xl border border-[#DCD7CC] overflow-hidden"
      >
        <div className="bg-[#F8F7F2] p-6 border-b border-[#E5E1D8] flex items-center justify-between">
          <h3 className="font-bold text-[#3D3D33] flex items-center gap-2">
            <Plus size={18} className="text-[#7C8A71]" />
            新增支出明細
          </h3>
          <button onClick={onClose} className="p-1 text-[#A5A58D] hover:text-[#7C8A71] transition-colors">
            <X size={20} />
          </button>
        </div>

        <div className="p-8 space-y-6 max-h-[70vh] overflow-y-auto custom-scrollbar">
          <div className="grid grid-cols-2 gap-6">
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">日期</label>
              <input 
                type="date"
                required
                className="w-full bg-[#FDFBF7] border border-[#E5E1D8] rounded-lg px-3 py-2 text-sm outline-none focus:border-[#7C8A71] focus:ring-1 focus:ring-[#7C8A71]/20 transition-all font-mono"
                value={localItem.date}
                onChange={e => setLocalItem({...localItem, date: e.target.value})}
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">地點</label>
              <input 
                className="w-full bg-[#FDFBF7] border border-[#E5E1D8] rounded-lg px-3 py-2 text-sm outline-none focus:border-[#7C8A71] transition-all placeholder:text-[#DCD7CC]"
                placeholder="請輸入地點"
                value={localItem.location}
                onChange={e => setLocalItem({...localItem, location: e.target.value})}
              />
            </div>
          </div>

          <div className="grid grid-cols-2 gap-6">
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">支出類別</label>
              <select 
                className="w-full bg-[#FDFBF7] border border-[#E5E1D8] rounded-lg px-3 py-2 text-sm outline-none focus:border-[#7C8A71] cursor-pointer transition-all"
                value={localItem.category}
                onChange={e => setLocalItem({...localItem, category: e.target.value as ExpenseCategory})}
              >
                {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            
            <AnimatePresence mode="wait">
              {localItem.category === '交通費' && (
                <motion.div 
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="space-y-1.5 overflow-hidden"
                >
                  <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">交通工具</label>
                  <select 
                    className="w-full bg-[#FDFBF7] border border-[#E5E1D8] rounded-lg px-3 py-2 text-sm outline-none focus:border-[#7C8A71] cursor-pointer transition-all"
                    value={localItem.transportMode}
                    onChange={e => setLocalItem({...localItem, transportMode: e.target.value as TransportMode})}
                  >
                    {TRANSPORT_MODES.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          <div className="grid grid-cols-2 gap-6">
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">金額</label>
              <input 
                type="number"
                className="w-full bg-[#FDFBF7] border border-[#E5E1D8] rounded-lg px-3 py-2 text-sm outline-none focus:border-[#7C8A71] text-right font-bold transition-all"
                value={localItem.amount}
                onChange={e => setLocalItem({...localItem, amount: Number(e.target.value)})}
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">幣別</label>
              <select 
                className="w-full bg-[#FDFBF7] border border-[#E5E1D8] rounded-lg px-3 py-2 text-sm outline-none focus:border-[#7C8A71] cursor-pointer transition-all"
                value={localItem.currency}
                onChange={e => setLocalItem({...localItem, currency: e.target.value})}
              >
                {CURRENCIES.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">專案代號</label>
            <input 
              className="w-full bg-[#FDFBF7] border border-[#E5E1D8] rounded-lg px-3 py-2 text-sm outline-none focus:border-[#7C8A71] font-mono transition-all placeholder:text-[#DCD7CC]"
              placeholder="請輸入專案代號"
              value={localItem.projectCode}
              onChange={e => setLocalItem({...localItem, projectCode: e.target.value})}
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">費用說明 (備註)</label>
            <textarea 
              rows={3}
              className="w-full bg-[#FDFBF7] border border-[#E5E1D8] rounded-lg px-3 py-2 text-sm outline-none focus:border-[#7C8A71] resize-none transition-all placeholder:text-[#DCD7CC] italic"
              placeholder="相關費用說明... (交際費請列對象/若刷公司卡請註記)"
              value={localItem.description}
              onChange={e => setLocalItem({...localItem, description: e.target.value})}
            />
          </div>
        </div>

        <div className="p-6 bg-[#FDFBF7] border-t border-[#E5E1D8] flex gap-3">
          <button 
            type="button"
            onClick={() => onSave(localItem, true)}
            className="flex-1 bg-white border border-[#7C8A71] text-[#7C8A71] font-bold py-2.5 px-4 rounded-lg text-xs uppercase tracking-widest hover:bg-[#F8F7F2] transition-all flex items-center justify-center gap-2"
          >
            <Save size={14} /> 保存並連續新增
          </button>
          <button 
            type="button"
            onClick={() => onSave(localItem, false)}
            className="flex-1 bg-[#7C8A71] text-white font-bold py-2.5 px-4 rounded-lg text-xs uppercase tracking-widest hover:bg-[#6A7661] transition-all shadow-md shadow-[#7C8A71]/20 flex items-center justify-center gap-2"
          >
            完成新增 <ChevronRight size={14} />
          </button>
        </div>
      </motion.div>
    </div>
  );
}


function ReportForm({ initialData, onCancel, onSubmit }: { 
  initialData: ExpenseReport | null,
  onCancel: () => void, 
  onSubmit: (r: ExpenseReport) => void, 
  key?: string 
}) {
  const [formData, setFormData] = useState({
    employeeName: initialData?.employeeName || '',
    employeeId: initialData?.employeeId || '',
    unit: initialData?.unit || '',
    department: initialData?.department || '',
    startDate: initialData?.startDate || '',
    endDate: initialData?.endDate || '',
  });

  const [items, setItems] = useState<ExpenseItem[]>(initialData?.items || []);
  const [prepaidItems, setPrepaidItems] = useState<PrepaidItem[]>(initialData?.prepaidItems || []);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingItem, setEditingItem] = useState<ExpenseItem | null>(null);

  // Migration for old reports with flat prepaidCurrency/prepaidAmount
  useEffect(() => {
    if (initialData && !initialData.prepaidItems && initialData.prepaidCurrency) {
      setPrepaidItems([{
        id: Math.random().toString(36).substr(2, 9),
        currency: initialData.prepaidCurrency,
        amount: initialData.prepaidAmount || 0
      }]);
    }
  }, [initialData]);

  const openAddItemModal = () => {
    setEditingItem({
      id: Math.random().toString(36).substr(2, 9),
      date: items.length > 0 ? items[items.length - 1].date : format(new Date(), 'yyyy-MM-dd'),
      location: items.length > 0 ? items[items.length - 1].location : '',
      description: '',
      currency: items.length > 0 ? items[items.length - 1].currency : 'TWD',
      category: '交通費',
      amount: 0,
      projectCode: items.length > 0 ? items[items.length - 1].projectCode : '',
      transportMode: '火車'
    });
    setIsModalOpen(true);
  };

  const handleModalSave = (item: ExpenseItem, keepOpen: boolean) => {
    const existingIdx = items.findIndex(i => i.id === item.id);
    if (existingIdx >= 0) {
      const newItems = [...items];
      newItems[existingIdx] = item;
      setItems(newItems);
    } else {
      setItems([...items, item]);
    }

    if (keepOpen) {
      setEditingItem({
        id: Math.random().toString(36).substr(2, 9),
        date: item.date,
        location: item.location,
        description: '',
        currency: item.currency,
        category: item.category,
        amount: 0,
        projectCode: item.projectCode,
        transportMode: item.transportMode
      });
    } else {
      setIsModalOpen(false);
      setEditingItem(null);
    }
  };

  const removeItem = (id: string) => {
    setItems(items.filter(item => item.id !== id));
  };

  const addPrepaidItem = () => {
    setPrepaidItems([...prepaidItems, {
      id: Math.random().toString(36).substr(2, 9),
      currency: 'TWD',
      amount: 0
    }]);
  };

  const removePrepaidItem = (id: string) => {
    setPrepaidItems(prepaidItems.filter(p => p.id !== id));
  };

  const updatePrepaidItem = (id: string, field: keyof PrepaidItem, value: any) => {
    setPrepaidItems(prepaidItems.map(p => p.id === id ? { ...p, [field]: value } : p));
  };

  const handleSubmit = (e: FormEvent) => {
    e.preventDefault();
    if (!formData.employeeName || !formData.employeeId || items.length === 0) {
      alert('請填寫完整資訊並至少新增一筆明細');
      return;
    }
    
    onSubmit({
      id: initialData?.id || Math.random().toString(36).substr(2, 9),
      submitDate: initialData?.submitDate || format(new Date(), 'yyyy-MM-dd'),
      ...formData,
      items,
      prepaidItems
    });
  };

  const expenseTotals = useMemo(() => {
    const res: Record<string, number> = {};
    items.forEach(item => {
      res[item.currency] = (res[item.currency] || 0) + (Number(item.amount) || 0);
    });
    return res;
  }, [items]);

  const prepaidTotals = useMemo(() => {
    const res: Record<string, number> = {};
    prepaidItems.forEach(p => {
      res[p.currency] = (res[p.currency] || 0) + (Number(p.amount) || 0);
    });
    return res;
  }, [prepaidItems]);

  const allCurrencies = useMemo(() => {
    return Array.from(new Set([
      ...Object.keys(expenseTotals),
      ...Object.keys(prepaidTotals)
    ]));
  }, [expenseTotals, prepaidTotals]);

  return (
    <motion.div 
      initial={{ opacity: 0, scale: 0.98 }}
      animate={{ opacity: 1, scale: 1 }}
      exit={{ opacity: 0, scale: 0.98 }}
      className="space-y-6 max-w-6xl mx-auto pb-20"
    >
      <div className="flex items-center justify-between mb-2">
        <button 
          onClick={onCancel}
          className="flex items-center gap-1 text-xs font-bold text-[#A5A58D] hover:text-[#7C8A71] transition-colors uppercase tracking-widest"
        >
          <ChevronLeft size={16} /> 返回儀表板
        </button>
        <div className="px-3 py-1 bg-[#F0F2EE] text-[#7C8A71] text-[10px] font-bold uppercase tracking-widest rounded">
          Draft Review
        </div>
      </div>

      <form onSubmit={handleSubmit} className="space-y-8">
        {/* Step 1: Base Information */}
        <section className="bg-white rounded-xl shadow-sm border border-[#DCD7CC] overflow-hidden">
          <div className="p-6 border-b border-[#E5E1D8] bg-[#F8F7F2]">
            <h2 className="font-bold text-lg flex items-center gap-2 text-[#3D3D33]">
              <User size={18} className="text-[#7C8A71]" />
              1. 基本出差資訊
            </h2>
          </div>
          <div className="p-8 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">出差人</label>
              <input 
                required
                className="w-full bg-white border-b border-[#E5E1D8] px-0 py-2 outline-none focus:border-[#7C8A71] transition-all text-sm placeholder:text-[#DCD7CC]"
                placeholder="請輸入姓名"
                value={formData.employeeName}
                onChange={e => setFormData({...formData, employeeName: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">工號</label>
              <input 
                required
                className="w-full bg-white border-b border-[#E5E1D8] px-0 py-2 outline-none focus:border-[#7C8A71] transition-all text-sm placeholder:text-[#DCD7CC]"
                placeholder="請輸入工號"
                value={formData.employeeId}
                onChange={e => setFormData({...formData, employeeId: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">
                單位 / 部門
              </label>
              <div className="flex gap-4">
                <input 
                  required
                  className="w-1/2 bg-white border-b border-[#E5E1D8] px-0 py-2 outline-none focus:border-[#7C8A71] transition-all text-sm placeholder:text-[#DCD7CC]"
                  placeholder="單位"
                  value={formData.unit}
                  onChange={e => setFormData({...formData, unit: e.target.value})}
                />
                <input 
                  className="w-1/2 bg-white border-b border-[#E5E1D8] px-0 py-2 outline-none focus:border-[#7C8A71] transition-all text-sm placeholder:text-[#DCD7CC]"
                  placeholder="部門 (選填)"
                  value={formData.department}
                  onChange={e => setFormData({...formData, department: e.target.value})}
                />
              </div>
            </div>
            <div className="space-y-2 lg:col-span-2">
              <label className="text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest block">
                出差期間
              </label>
              <div className="flex items-center gap-4">
                <input 
                  required
                  type="date"
                  className="flex-1 bg-white border-b border-[#E5E1D8] px-0 py-2 outline-none focus:border-[#7C8A71] transition-all text-sm"
                  value={formData.startDate}
                  onChange={e => setFormData({...formData, startDate: e.target.value})}
                />
                <input 
                  required
                  type="date"
                  className="flex-1 bg-white border-b border-[#E5E1D8] px-0 py-2 outline-none focus:border-[#7C8A71] transition-all text-sm"
                  value={formData.endDate}
                  onChange={e => setFormData({...formData, endDate: e.target.value})}
                />
              </div>
            </div>
          </div>
        </section>

        {/* Step 2: Expense Items */}
        <section className="bg-white rounded-xl shadow-sm border border-[#DCD7CC] overflow-hidden">
          <div className="p-6 border-b border-[#E5E1D8] bg-[#F8F7F2] flex items-center justify-between">
            <h2 className="font-bold text-lg flex items-center gap-2 text-[#3D3D33]">
              <MapPin size={18} className="text-[#7C8A71]" />
              2. 支出明細填寫
            </h2>
            <div className="flex items-center gap-4">
              <span className="text-[10px] text-[#A5A58D] italic hidden sm:block">專案代號查詢：雲端專案編號清單</span>
              <button 
                type="button"
                onClick={openAddItemModal}
                className="flex items-center gap-1.5 text-[10px] font-bold bg-[#7C8A71] text-white px-4 py-2 rounded uppercase tracking-widest hover:bg-[#6A7661] transition-all shadow-md shadow-[#7C8A71]/20"
              >
                <Plus size={14} /> 新增明細
              </button>
            </div>
          </div>
          
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="border-b border-[#E5E1D8] bg-[#FDFBF7] text-[10px] font-bold text-[#7C8A71] uppercase tracking-widest">
                  <th className="px-6 py-4 w-32">日期</th>
                  <th className="px-6 py-4 w-40">地點 / 專案</th>
                  <th className="px-6 py-4 w-40">類別 / 工具</th>
                  <th className="px-6 py-4 w-32">金額 / 幣別</th>
                  <th className="px-6 py-4">費用說明 (備註)</th>
                  <th className="px-6 py-4 text-center w-24">操作</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#F0EFEC]">
                {items.length === 0 && (
                  <tr>
                    <td colSpan={6} className="px-6 py-12 text-center text-[#A5A58D] italic text-sm bg-white">
                      尚未新增任何費用明細，請點擊右上角新增。
                    </td>
                  </tr>
                )}
                <AnimatePresence>
                  {items.map((item) => (
                    <motion.tr 
                      key={item.id}
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, x: 20 }}
                      className="hover:bg-[#FDFCF8] transition-colors group"
                    >
                      <td className="px-6 py-4 align-top whitespace-nowrap">
                        <span className="text-sm font-medium text-[#3D3D33]">{item.date}</span>
                      </td>
                      <td className="px-6 py-4 align-top">
                        <div className="flex flex-col">
                          <span className="text-sm text-[#3D3D33]">{item.location}</span>
                          <span className="text-[10px] font-mono text-[#A5A58D]">{item.projectCode}</span>
                        </div>
                      </td>
                      <td className="px-6 py-4 align-top">
                        <div className="flex flex-col gap-1">
                          <span className="inline-block px-1.5 py-0.5 bg-[#F5F5F0] rounded text-[10px] w-fit">{item.category}</span>
                          {item.category === '交通費' && (
                            <span className="text-[10px] text-[#A5A58D]">{item.transportMode}</span>
                          )}
                        </div>
                      </td>
                      <td className="px-6 py-4 align-top">
                        <div className="flex flex-col">
                          <span className="text-sm font-bold text-[#3D3D33]">{item.amount.toLocaleString()}</span>
                          <span className="text-[10px] text-[#A5A58D]">{item.currency}</span>
                        </div>
                      </td>
                      <td className="px-6 py-4 align-top">
                        <p className="text-xs text-[#A5A58D] italic line-clamp-2">{item.description}</p>
                      </td>
                      <td className="px-6 py-4 align-top text-center">
                        <div className="flex items-center justify-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                          <button 
                            type="button"
                            onClick={() => {
                              setEditingItem(item);
                              setIsModalOpen(true);
                            }}
                            className="p-1.5 text-[#A5A58D] hover:text-[#7C8A71] hover:bg-[#F0F2EE] rounded transition-all"
                          >
                            <Edit2 size={14} />
                          </button>
                          <button 
                            type="button"
                            onClick={() => removeItem(item.id)}
                            className="p-1.5 text-[#DCD7CC] hover:text-red-500 hover:bg-red-50 rounded transition-all"
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </td>
                    </motion.tr>
                  ))}
                </AnimatePresence>
              </tbody>
            </table>
          </div>
        </section>

        <ExpenseItemModal 
          isOpen={isModalOpen}
          item={editingItem}
          onClose={() => {
            setIsModalOpen(false);
            setEditingItem(null);
          }}
          onSave={handleModalSave}
        />

        {/* Step 3: Summary & Prepayment */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 items-start">
          <section className="bg-white rounded-xl shadow-sm border border-[#DCD7CC] overflow-hidden">
            <div className="p-6 border-b border-[#E5E1D8] bg-[#F8F7F2] flex items-center justify-between">
              <h2 className="font-bold text-lg flex items-center gap-2 text-[#3D3D33]">
                <CircleDollarSign size={18} className="text-[#7C8A71]" />
                3. 已先預支費用
              </h2>
              <button 
                type="button"
                onClick={addPrepaidItem}
                className="flex items-center gap-1.5 text-[10px] font-bold bg-[#A5A58D] text-white px-3 py-1.5 rounded uppercase tracking-widest hover:bg-[#8B8B75] transition-all"
              >
                <Plus size={12} /> 新增預支
              </button>
            </div>
            <div className="p-0">
              {prepaidItems.length === 0 ? (
                <div className="p-8 text-center text-[#A5A58D] italic text-xs">
                  無預支明細，如無預支則跳過此項。
                </div>
              ) : (
                <div className="divide-y divide-[#F0EFEC]">
                  {prepaidItems.map((p) => (
                    <div key={p.id} className="p-6 flex items-center gap-6 group">
                      <div className="flex-1 grid grid-cols-2 gap-4">
                        <div className="space-y-1.5">
                          <label className="text-[9px] font-bold text-[#A5A58D] uppercase tracking-widest">預支幣別</label>
                          <select 
                            className="w-full bg-white border-b border-[#E5E1D8] px-0 py-1 outline-none focus:border-[#7C8A71] transition-all text-sm"
                            value={p.currency}
                            onChange={e => updatePrepaidItem(p.id, 'currency', e.target.value)}
                          >
                            {CURRENCIES.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>
                        <div className="space-y-1.5 text-right">
                          <label className="text-[9px] font-bold text-[#A5A58D] uppercase tracking-widest">金額</label>
                          <input 
                            type="number"
                            className="w-full bg-white border-b border-[#E5E1D8] px-0 py-1 outline-none focus:border-[#7C8A71] transition-all text-sm text-right font-medium"
                            value={p.amount}
                            onChange={e => updatePrepaidItem(p.id, 'amount', Number(e.target.value))}
                          />
                        </div>
                      </div>
                      <button 
                        type="button"
                        onClick={() => removePrepaidItem(p.id)}
                        className="p-1.5 text-[#DCD7CC] hover:text-red-500 hover:bg-red-50 rounded transition-all opacity-0 group-hover:opacity-100"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </section>

          <section className="bg-[#F8F7F2] rounded-xl border-2 border-dashed border-[#DCD7CC] p-8 space-y-8">
            <div className="space-y-6">
              <div className="flex items-center justify-between">
                <h2 className="font-bold text-[#7C8A71] uppercase tracking-[0.2em] text-xs">費用報支結算</h2>
                <div className="flex items-center gap-2">
                   <div className="w-2 h-2 rounded-full bg-[#7C8A71] animate-pulse" />
                   <span className="text-[10px] font-mono text-[#A5A58D]">Auto Calc</span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-6">
                <div className="space-y-4">
                  <p className="text-[10px] text-[#A5A58D] font-bold uppercase tracking-widest">總計支出項目</p>
                  <div className="space-y-3">
                    {Object.keys(expenseTotals).length === 0 ? (
                      <p className="text-[#DCD7CC] italic text-xs">無數據</p>
                    ) : (
                      Object.entries(expenseTotals).map(([curr, amt]) => (
                        <div key={curr} className="flex items-center justify-between">
                          <span className="text-[10px] font-bold text-[#A5A58D]">{curr}</span>
                          <span className="text-lg font-light text-[#3D3D33]">{amt.toLocaleString()}</span>
                        </div>
                      ))
                    )}
                  </div>
                </div>

                <div className="space-y-6 border-l border-[#DCD7CC] pl-6 flex flex-col justify-between">
                   <div className="space-y-4">
                      <p className="text-[10px] text-[#7C8A71] font-bold uppercase tracking-widest">應付員工 / 繳回</p>
                      <div className="space-y-4">
                        {allCurrencies.length === 0 ? (
                          <span className="text-[#DCD7CC] italic text-[11px]">無計算數據</span>
                        ) : (
                          allCurrencies.map(curr => {
                            const exp = expenseTotals[curr] || 0;
                            const pre = prepaidTotals[curr] || 0;
                            const diff = exp - pre;
                            return (
                              <div key={curr} className="flex flex-col">
                                <span className="text-[9px] font-bold text-[#A5A58D] uppercase">{curr} {diff >= 0 ? "公司支付" : "員工繳回"}</span>
                                <span className={cn(
                                  "text-2xl font-bold tracking-tight",
                                  diff >= 0 ? "text-[#4F5946]" : "text-[#A66B56]"
                                )}>
                                  {diff.toLocaleString()}
                                </span>
                              </div>
                            );
                          })
                        )}
                      </div>
                   </div>

                   <div className="flex gap-2">
                      <button 
                        type="submit"
                        className="flex-1 bg-[#7C8A71] text-white font-bold py-3 px-6 rounded-lg text-[11px] uppercase tracking-widest hover:bg-[#6A7661] transition-all shadow-lg shadow-[#7C8A71]/20 flex items-center justify-center gap-2"
                      >
                        提交報支單 <ChevronRight size={14} />
                      </button>
                   </div>
                </div>
              </div>
            </div>
          </section>
        </div>
      </form>
    </motion.div>
  );
}
