import React, { useState, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, Printer, AlertCircle, CheckCircle2, Calendar, FileText, ChevronUp, ChevronDown, Package, PackageX } from 'lucide-react';

// ==========================================
// 核心系統設定與常數
// ==========================================
const REQUIRED_HEADERS = {
  sales: ['客戶編號', '客戶姓名', '銷貨單編號', '商品名稱', '折抵後單價', '出貨數量', '是否為贈品', '金額小計'],
  returns: ['客戶編號', '客戶姓名', '退換貨單號', '商品名稱', '折抵後單價', '申請退換貨數量', '申請退換貨金額小計']
};

export default function App() {
  // --- 狀態管理 ---
  const [dates, setDates] = useState({ start: '', end: '' });
  const [fileStatus, setFileStatus] = useState({ sales: null, returns: null });
  const [parsedData, setParsedData] = useState({ sales: [], returns: [] });
  const [errors, setErrors] = useState([]);
  const [reportData, setReportData] = useState(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isPanelOpen, setIsPanelOpen] = useState(true);

  // --- 動態載入 SheetJS (xlsx) ---
  useEffect(() => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.async = true;
    document.body.appendChild(script);
    return () => {
      document.body.removeChild(script);
    };
  }, []);

  // ==========================================
  // 工具函式庫 (Utility Functions)
  // ==========================================

  const parseAmount = (val) => Number(String(val).replace(/[^0-9.-]+/g, "")) || 0;
  const formatCurrency = (num) => Math.round(num).toLocaleString('en-US');

  // 提取數字用作排序 (例如 C001 -> 1)
  const extractNumber = (str) => parseInt(String(str).replace(/\D/g, ''), 10) || 0;

  // 判斷是否為贈品 (彈性字串比對)
  const isGift = (val) => {
    if (val === undefined || val === null) return false;
    const str = String(val).trim().toLowerCase();
    return ['是', 'y', 'true', '1'].includes(str);
  };

  // 計算 RowSpan 陣列 (用於報表畫面視覺合併)
  const computeRowSpans = (items, key) => {
    let spans = new Array(items.length).fill(0);
    let i = 0;
    while (i < items.length) {
      let count = 1;
      for (let j = i + 1; j < items.length; j++) {
        if (items[i][key] === items[j][key]) {
          count++;
        } else {
          break;
        }
      }
      spans[i] = count;
      i += count;
    }
    return spans;
  };

  // ==========================================
  // 業務邏輯：檔案上傳與驗證
  // ==========================================
  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!window.XLSX) {
      setErrors(prev => [...prev, "Excel 處理套件尚未載入完成，請稍候再試。"]);
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = window.XLSX.read(bstr, { type: 'binary' });
        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        
        const data = window.XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (data.length === 0) throw new Error("上傳的檔案為空。");

        const headers = data[0];
        const requiredHeaders = REQUIRED_HEADERS[type];
        
        const missingHeaders = requiredHeaders.filter(h => !headers.includes(h));
        if (missingHeaders.length > 0) {
          const typeName = type === 'sales' ? '銷貨商品明細' : '銷退單商品明細';
          setErrors(prev => [...prev.filter(err => !err.includes(typeName)), 
            `【${typeName}】缺少必要欄位：${missingHeaders.join(', ')}`]);
          setFileStatus(prev => ({ ...prev, [type]: 'error' }));
          setParsedData(prev => ({ ...prev, [type]: [] }));
          return;
        }

        const jsonData = window.XLSX.utils.sheet_to_json(ws, { raw: false, defval: "" });
        
        setParsedData(prev => ({ ...prev, [type]: jsonData }));
        setFileStatus(prev => ({ ...prev, [type]: 'success' }));
        const typeName = type === 'sales' ? '銷貨商品明細' : '銷退單商品明細';
        setErrors(prev => prev.filter(err => !err.includes(typeName)));

      } catch (error) {
        setErrors(prev => [...prev, `解析檔案發生錯誤: ${error.message}`]);
        setFileStatus(prev => ({ ...prev, [type]: 'error' }));
      }
    };
    reader.readAsBinaryString(file);
  };

  // ==========================================
  // 業務邏輯：資料彙整與報表生成
  // ==========================================
  const generateReport = () => {
    if (parsedData.sales.length === 0 && parsedData.returns.length === 0) {
      setErrors(["請至少上傳一份有效且包含數據的 Excel 檔案再生成報表。"]);
      return;
    }

    setIsGenerating(true);

    setTimeout(() => {
      const customerMap = new Map();

      const ensureCustomer = (id, name) => {
        if (!customerMap.has(id)) {
          customerMap.set(id, { customerID: id, customerName: name, sales: [], returns: [], totalSales: 0, totalReturns: 0 });
        }
      };

      // 匯入銷貨資料
      parsedData.sales.forEach(row => {
        const cID = row['客戶編號'];
        ensureCustomer(cID, row['客戶姓名']);
        const amount = parseAmount(row['金額小計']);
        customerMap.get(cID).sales.push({
          ...row,
          '折抵後單價': parseAmount(row['折抵後單價']),
          '出貨數量': parseAmount(row['出貨數量']),
          '金額小計': amount
        });
        customerMap.get(cID).totalSales += amount;
      });

      // 匯入退貨資料
      parsedData.returns.forEach(row => {
        const cID = row['客戶編號'];
        ensureCustomer(cID, row['客戶姓名']);
        const amount = parseAmount(row['申請退換貨金額小計']);
        customerMap.get(cID).returns.push({
          ...row,
          '折抵後單價': parseAmount(row['折抵後單價']),
          '申請退換貨數量': parseAmount(row['申請退換貨數量']),
          '申請退換貨金額小計': amount
        });
        customerMap.get(cID).totalReturns += amount;
      });

      let globalTotalSales = 0;
      let globalTotalReturns = 0;
      
      const customersList = Array.from(customerMap.values()).map(c => {
        c.netTotal = c.totalSales - c.totalReturns;
        globalTotalSales += c.totalSales;
        globalTotalReturns += c.totalReturns;

        // 內部明細排序 (依單號字串排序)
        c.sales.sort((a, b) => String(a['銷貨單編號']).localeCompare(String(b['銷貨單編號'])));
        c.returns.sort((a, b) => String(a['退換貨單號']).localeCompare(String(b['退換貨單號'])));
        
        return c;
      });

      // 客戶層級排序 (依客戶編號數字排序)
      customersList.sort((a, b) => extractNumber(a.customerID) - extractNumber(b.customerID));

      setReportData({
        summary: {
          grandTotal: globalTotalSales - globalTotalReturns,
          printTime: new Date().toLocaleString('zh-TW', { hour12: false })
        },
        details: customersList
      });

      setIsGenerating(false);
      setErrors([]);
      setIsPanelOpen(false); // 自動收合面板
    }, 500);
  };

  const triggerPrint = () => {
    window.print();
  };

  // ==========================================
  // UI 渲染區塊
  // ==========================================
  return (
    <div className="min-h-screen bg-gray-50 text-gray-800 font-sans">
      
      {/* 頂部操作區塊 */}
      <div className="print:hidden bg-white shadow-sm border-b border-gray-200 p-6 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <h1 className="text-2xl font-bold text-blue-900 flex items-center gap-2">
              <FileSpreadsheet className="w-6 h-6" />
              月結請款單明細表生成器
            </h1>
            <p className="text-sm text-gray-500 mt-1">上傳包含商品資訊的銷貨與退貨明細，自動彙整為 A4 排版對帳單。</p>
          </div>
          <div className="flex gap-3">
            <button 
              onClick={() => setIsPanelOpen(!isPanelOpen)}
              className="bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors shadow-sm"
            >
              {isPanelOpen ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
              {isPanelOpen ? '收合面板' : '展開面板'}
            </button>
            {reportData && (
              <button 
                onClick={triggerPrint}
                className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-md font-medium flex items-center gap-2 transition-colors shadow-sm"
              >
                <Printer className="w-5 h-5" />
                列印 / 存為 PDF
              </button>
            )}
          </div>
        </div>

        {/* 設定面板 */}
        <div className={`transition-all duration-300 ${isPanelOpen ? 'block' : 'hidden'}`}>
          <div className="max-w-6xl mx-auto mt-6 grid grid-cols-1 md:grid-cols-3 gap-6">
            
            <div className="bg-gray-50 p-4 rounded-lg border border-gray-200">
              <h3 className="font-semibold text-gray-700 flex items-center gap-2 mb-3">
                <Calendar className="w-4 h-4" /> 報表日期範圍
              </h3>
              <div className="flex gap-2 items-center">
                <input 
                  type="date" 
                  value={dates.start} 
                  onChange={e => setDates({...dates, start: e.target.value})}
                  className="flex-1 border border-gray-300 rounded p-2 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none" 
                />
                <span className="text-gray-500">至</span>
                <input 
                  type="date" 
                  value={dates.end} 
                  onChange={e => setDates({...dates, end: e.target.value})}
                  className="flex-1 border border-gray-300 rounded p-2 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none" 
                />
              </div>
              <p className="text-xs text-gray-400 mt-2">*僅顯示於報表表頭，不進行資料過濾。</p>
            </div>

            <div className="bg-gray-50 p-4 rounded-lg border border-gray-200">
              <h3 className="font-semibold text-gray-700 flex items-center gap-2 mb-3">
                <UploadCloud className="w-4 h-4" /> 銷貨商品明細
              </h3>
              <label className="block w-full cursor-pointer bg-white border border-dashed border-gray-300 hover:border-blue-500 p-3 rounded text-center transition-colors">
                <span className="text-sm text-gray-600">點擊選擇或拖曳 Excel 檔案</span>
                <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={(e) => handleFileUpload(e, 'sales')} />
              </label>
              {fileStatus.sales === 'success' && <div className="mt-2 text-xs text-green-600 flex items-center gap-1"><CheckCircle2 className="w-3 h-3"/> 上傳並解析成功</div>}
              {fileStatus.sales === 'error' && <div className="mt-2 text-xs text-red-600 flex items-center gap-1"><AlertCircle className="w-3 h-3"/> 上傳失敗</div>}
            </div>

            <div className="bg-gray-50 p-4 rounded-lg border border-gray-200">
              <h3 className="font-semibold text-gray-700 flex items-center gap-2 mb-3">
                <UploadCloud className="w-4 h-4" /> 銷退單商品明細
              </h3>
              <label className="block w-full cursor-pointer bg-white border border-dashed border-gray-300 hover:border-blue-500 p-3 rounded text-center transition-colors">
                <span className="text-sm text-gray-600">點擊選擇或拖曳 Excel 檔案</span>
                <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={(e) => handleFileUpload(e, 'returns')} />
              </label>
              {fileStatus.returns === 'success' && <div className="mt-2 text-xs text-green-600 flex items-center gap-1"><CheckCircle2 className="w-3 h-3"/> 上傳並解析成功</div>}
              {fileStatus.returns === 'error' && <div className="mt-2 text-xs text-red-600 flex items-center gap-1"><AlertCircle className="w-3 h-3"/> 上傳失敗</div>}
            </div>
          </div>

          {errors.length > 0 && (
            <div className="max-w-6xl mx-auto mt-4 bg-red-50 border-l-4 border-red-500 p-4 rounded-md">
              <div className="flex items-start">
                <AlertCircle className="w-5 h-5 text-red-500 mr-2 mt-0.5" />
                <div>
                  <h4 className="text-red-800 font-medium text-sm">請注意以下錯誤：</h4>
                  <ul className="list-disc list-inside text-sm text-red-700 mt-1">
                    {errors.map((err, i) => <li key={i}>{err}</li>)}
                  </ul>
                </div>
              </div>
            </div>
          )}

          <div className="max-w-6xl mx-auto mt-6 text-center">
            <button 
              onClick={generateReport}
              disabled={isGenerating}
              className={`px-8 py-3 rounded-md font-bold text-white shadow-md transition-all ${isGenerating ? 'bg-gray-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 hover:shadow-lg'}`}
            >
              {isGenerating ? '資料彙整中...' : '生成報表預覽'}
            </button>
          </div>
        </div>
      </div>

      {/* 報表預覽區塊 */}
      {reportData && (
        <div className="p-6 print:p-0">
          <div className="max-w-[210mm] mx-auto bg-white shadow-lg print:shadow-none print:max-w-none text-sm leading-relaxed">
            
            {/* ================================== */}
            {/* 第一頁：總表摘要 (Page 1)            */}
            {/* ================================== */}
            <div className="p-16 print:p-10 bg-white min-h-[297mm] flex flex-col justify-center items-center border-b border-gray-200 print:border-none">
              <div className="text-center w-full max-w-2xl">
                <h2 className="text-4xl font-extrabold text-gray-900 tracking-widest mb-6">月結請款單 - 明細表</h2>
                
                <div className="bg-gray-50 p-8 rounded-lg border border-gray-200 mb-8 w-full shadow-sm">
                  <div className="grid grid-cols-2 gap-6 text-left mb-6">
                    <div>
                      <span className="block text-gray-500 text-sm mb-1">報表期間</span>
                      <span className="font-semibold text-gray-800 text-lg">
                        {dates.start ? dates.start.replace(/-/g, '/') : '未指定'} ~ {dates.end ? dates.end.replace(/-/g, '/') : '未指定'}
                      </span>
                    </div>
                    <div>
                      <span className="block text-gray-500 text-sm mb-1">製表時間</span>
                      <span className="font-semibold text-gray-800 text-lg">{reportData.summary.printTime}</span>
                    </div>
                  </div>
                  
                  <div className="pt-6 border-t border-gray-300">
                    <span className="block text-gray-600 text-base mb-2 text-center">本期請款總金額 (Grand Total)</span>
                    <span className="block text-center font-bold text-blue-800 text-5xl">
                      NT$ {formatCurrency(reportData.summary.grandTotal)}
                    </span>
                  </div>
                </div>
                <p className="text-gray-400 text-sm">請翻閱次頁查看各客戶明細帳單。</p>
              </div>
            </div>

            {/* ================================== */}
            {/* 第二頁起：客戶明細 (Customer Details) */}
            {/* ================================== */}
            {reportData.details.map((customer) => {
              // 分別計算銷貨單與退貨單的 RowSpan
              const salesRowSpans = computeRowSpans(customer.sales, '銷貨單編號');
              const returnsRowSpans = computeRowSpans(customer.returns, '退換貨單號');

              return (
                <div key={customer.customerID} className="p-10 print:p-10 bg-white" style={{ pageBreakBefore: 'always' }}>
                  
                  {/* 客戶標題列 */}
                  <div className="mb-6 pb-3 border-b-2 border-gray-800 flex justify-between items-end">
                    <div>
                      <h3 className="text-2xl font-bold text-gray-900 tracking-wide">
                        {customer.customerName} <span className="text-lg text-gray-500 font-normal ml-2">{customer.customerID}</span>
                      </h3>
                      <p className="text-sm text-gray-600 mt-1">月結請款單明細</p>
                    </div>
                    <div className="text-right text-sm text-gray-600">
                      期間：{dates.start ? dates.start.replace(/-/g, '/') : '未指定'} ~ {dates.end ? dates.end.replace(/-/g, '/') : '未指定'}
                    </div>
                  </div>

                  {/* 表格 1：銷貨單商品明細 */}
                  <div className="mb-8">
                    <h4 className="font-bold text-lg text-gray-800 mb-2 flex items-center gap-2">
                      <Package className="w-5 h-5 text-blue-700" /> 銷貨單商品明細
                    </h4>
                    <table className="w-full border-collapse border border-gray-800">
                      <thead className="table-header-group">
                        <tr className="bg-gray-100 text-gray-800">
                          <th className="border border-gray-800 py-2 px-3 text-center whitespace-nowrap w-36">銷貨單編號</th>
                          <th className="border border-gray-800 py-2 px-3 text-left w-auto">商品名稱</th>
                          <th className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap w-20">折抵後單價</th>
                          <th className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap w-16">數量</th>
                          <th className="border border-gray-800 py-2 px-3 text-center whitespace-nowrap w-16">贈品</th>
                          <th className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap w-24">金額小計</th>
                        </tr>
                      </thead>
                      <tbody>
                        {customer.sales.length > 0 ? (
                          customer.sales.map((sale, i) => {
                            const span = salesRowSpans[i];
                            return (
                              <tr key={`s-${i}`} className="hover:bg-gray-50" style={{ pageBreakInside: 'avoid' }}>
                                {span > 0 && (
                                  <td rowSpan={span} className="border border-gray-800 py-2 px-3 text-center bg-white align-top font-medium text-blue-900 whitespace-nowrap">{sale['銷貨單編號']}</td>
                                )}
                                <td className="border border-gray-800 py-2 px-3 text-left text-gray-700 break-words">{sale['商品名稱']}</td>
                                <td className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap">${formatCurrency(sale['折抵後單價'])}</td>
                                <td className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap">{formatCurrency(sale['出貨數量'])}</td>
                                <td className="border border-gray-800 py-2 px-3 text-center whitespace-nowrap">
                                  {isGift(sale['是否為贈品']) && (
                                    <span className="inline-block bg-orange-100 text-orange-800 text-xs px-2 py-0.5 rounded-full font-medium border border-orange-200">
                                      贈品
                                    </span>
                                  )}
                                </td>
                                <td className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap font-medium">${formatCurrency(sale['金額小計'])}</td>
                              </tr>
                            );
                          })
                        ) : (
                          <tr>
                            <td colSpan="6" className="border border-gray-800 py-6 text-center text-gray-500 italic">
                              本期無銷貨紀錄
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>

                  {/* 表格 2：退換貨商品明細 */}
                  <div className="mb-8">
                    <h4 className="font-bold text-lg text-gray-800 mb-2 flex items-center gap-2">
                      <PackageX className="w-5 h-5 text-red-700" /> 退換貨商品明細
                    </h4>
                    <table className="w-full border-collapse border border-gray-800">
                      <thead className="table-header-group">
                        <tr className="bg-gray-100 text-gray-800">
                          <th className="border border-gray-800 py-2 px-3 text-center whitespace-nowrap w-36">退換貨單號</th>
                          <th className="border border-gray-800 py-2 px-3 text-left w-auto">商品名稱</th>
                          <th className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap w-20">折抵後單價</th>
                          <th className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap w-20">退換貨數量</th>
                          <th className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap w-32">金額小計</th>
                        </tr>
                      </thead>
                      <tbody>
                        {customer.returns.length > 0 ? (
                          customer.returns.map((ret, i) => {
                            const span = returnsRowSpans[i];
                            return (
                              <tr key={`r-${i}`} className="hover:bg-gray-50" style={{ pageBreakInside: 'avoid' }}>
                                {span > 0 && (
                                  <td rowSpan={span} className="border border-gray-800 py-2 px-3 text-center bg-white align-top font-medium text-red-900 whitespace-nowrap">{ret['退換貨單號']}</td>
                                )}
                                <td className="border border-gray-800 py-2 px-3 text-left text-gray-700 break-words">{ret['商品名稱']}</td>
                                <td className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap">${formatCurrency(ret['折抵後單價'])}</td>
                                <td className="border border-gray-800 py-2 px-3 text-right whitespace-nowrap">{formatCurrency(ret['申請退換貨數量'])}</td>
                                <td className="border border-gray-800 py-2 px-3 text-right text-red-700 whitespace-nowrap font-medium">
                                  -${formatCurrency(ret['申請退換貨金額小計'])}
                                </td>
                              </tr>
                            );
                          })
                        ) : (
                          <tr>
                            <td colSpan="5" className="border border-gray-800 py-6 text-center text-gray-500 italic">
                              本期無退換貨紀錄
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>

                  {/* 客戶結算列 */}
                  <div className="mt-6 border-t border-gray-300 pt-4 flex justify-end">
                    <div className="bg-blue-50/50 px-6 py-4 rounded border border-blue-100 shadow-sm flex items-center gap-4">
                      <span className="text-gray-700 font-medium">
                        {customer.customerName} 本期請款金額：
                      </span>
                      <span className="text-2xl font-bold text-blue-900 tracking-tight">
                        NT$ {formatCurrency(customer.netTotal)}
                      </span>
                    </div>
                  </div>

                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}