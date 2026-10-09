const { useState, useMemo, useEffect, useRef } = React;

function ERPApp() {
  // 1. 데이터 상태
  const [supplierData, setSupplierData] = useState([]);
  const [customerData, setCustomerData] = useState([]);
  const [itemMasterData, setItemMasterData] = useState([]);

  // 2. 메인 검색 & 캘린더 상태
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedSearchItem, setSelectedSearchItem] = useState(null);
  const [showDropdown, setShowDropdown] = useState(false);
  const searchContainerRef = useRef(null);

  const today = new Date();
  const [currentYear, setCurrentYear] = useState(today.getFullYear());
  const [currentMonth, setCurrentMonth] = useState(today.getMonth());
  const [selectedCalendarDate, setSelectedCalendarDate] = useState(null);

  // 3. 부족예정 경고 전용 상태 (필터 & 검색)
  const [alertFilter, setAlertFilter] = useState('ALL'); // ALL, YELLOW, RED
  const [alertCustomerFilter, setAlertCustomerFilter] = useState('ALL'); // 고객사 필터
  const [alertSearchTerm, setAlertSearchTerm] = useState(''); // 경고 목록 전용 품목 검색
  const [calendarRedHighlightDate, setCalendarRedHighlightDate] = useState(null); 
  const [expandedAlertId, setExpandedAlertId] = useState(null); 

  const getTodayString = () => {
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  };

  // 엑셀 날짜 정제
  const parseExcelDate = (val) => {
    if (val === null || val === undefined || val === '') return '-';

    if (typeof val === 'number') {
      const date = XLSX.SSF.parse_date_code(val);
      if (date) {
        const y = date.y;
        const m = String(date.m).padStart(2, '0');
        const d = String(date.d).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
    }

    let str = String(val).trim();

    if (str.includes('월')) {
      const match = str.match(/(\d+)\s*월\s*(\d+)\s*일?/);
      if (match) {
        const m = String(match[1]).padStart(2, '0');
        const d = String(match[2]).padStart(2, '0');
        return `${currentYear}-${m}-${d}`;
      }
    }

    str = str.replace(/[\.\/]/g, '-');
    const parts = str.split('-').filter(Boolean);

    if (parts.length === 3) {
      const y = parts[0].length === 2 ? `20${parts[0]}` : parts[0];
      const m = parts[1].padStart(2, '0');
      const d = parts[2].padStart(2, '0');
      return `${y}-${m}-${d}`;
    } else if (parts.length === 2) {
      const m = parts[0].padStart(2, '0');
      const d = parts[1].padStart(2, '0');
      return `${currentYear}-${m}-${d}`;
    }

    return str;
  };

  const getCellValue = (worksheet, colLetter, rowNum) => {
    const cellAddress = `${colLetter}${rowNum}`;
    const cell = worksheet[cellAddress];
    return cell ? cell.v : undefined;
  };

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (searchContainerRef.current && !searchContainerRef.current.contains(e.target)) {
        setShowDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // 마스터 품목 자동 통합 동기화
  const syncMasterListWithAllItems = (newSupplierList, newCustomerList, currentMasterList) => {
    const masterMap = new Map();

    currentMasterList.forEach(item => {
      masterMap.set(item.itemCode, item);
    });

    const registerAutoItem = (itemCode) => {
      if (!itemCode || itemCode === 'undefined' || itemCode === '품목코드') return;
      if (!masterMap.has(itemCode)) {
        masterMap.set(itemCode, {
          id: `auto-master-${itemCode}`,
          itemCode: itemCode,
          futuresCode: '미등록',
          currentStock: 0,
          customerCodes: {}
        });
      }
    };

    newSupplierList.forEach(sup => registerAutoItem(sup.itemCode));
    newCustomerList.forEach(cust => registerAutoItem(cust.itemCode));

    return Array.from(masterMap.values());
  };

  // --------------------------------------------------------------------------
  // 엑셀 업로드 파일 핸들러
  // --------------------------------------------------------------------------
  const handleSupplierFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });
        let parsedResults = [];

        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          if (!worksheet) return;
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
          const trimmedSheetName = sheetName.trim();

          for (let R = 2; R <= range.e.r; R++) {
            const excelRowNum = R + 1;
            let itemCode = '', dueDate = '', qty = 0, customer = '';

            if (trimmedSheetName === '유니락') {
              const rawCode = getCellValue(worksheet, 'D', excelRowNum);
              const rawDate = getCellValue(worksheet, 'I', excelRowNum);
              const rawQty = getCellValue(worksheet, 'P', excelRowNum);

              itemCode = rawCode !== undefined ? String(rawCode).trim() : '';
              dueDate = parseExcelDate(rawDate);
              qty = rawQty !== undefined ? Number(rawQty) || 0 : 0;
              customer = '-';
            } else {
              const rawCode = getCellValue(worksheet, 'B', excelRowNum);
              const rawDate = getCellValue(worksheet, 'C', excelRowNum);
              const rawQty = getCellValue(worksheet, 'D', excelRowNum);
              const rawCust = getCellValue(worksheet, 'E', excelRowNum);

              itemCode = rawCode !== undefined ? String(rawCode).trim() : '';
              dueDate = parseExcelDate(rawDate);
              qty = rawQty !== undefined ? Number(rawQty) || 0 : 0;
              customer = rawCust !== undefined ? String(rawCust).trim() : '-';
            }

            if (itemCode && itemCode !== 'undefined' && itemCode !== '품목코드') {
              parsedResults.push({
                id: `sup-${trimmedSheetName}-${excelRowNum}`,
                supplier: trimmedSheetName,
                itemCode,
                dueDate,
                qty,
                customer
              });
            }
          }
        });

        setSupplierData(parsedResults);
        setItemMasterData(prev => syncMasterListWithAllItems(parsedResults, customerData, prev));
        alert(`공급사.xlsx 로드 완료! (${parsedResults.length}건)`);
      } catch (err) {
        console.error(err);
        alert('공급사 파싱 에러');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleCustomerFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });
        let parsedResults = [];

        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          if (!worksheet) return;
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
          const trimmedSheetName = sheetName.trim();

          for (let R = 2; R <= range.e.r; R++) {
            const excelRowNum = R + 1;
            let itemCode = '', outDate = '', qty = 0, customer = '';

            if (trimmedSheetName === '그외') {
              const rawCode = getCellValue(worksheet, 'B', excelRowNum);
              const rawDate = getCellValue(worksheet, 'C', excelRowNum);
              const rawQty = getCellValue(worksheet, 'D', excelRowNum);
              const rawCust = getCellValue(worksheet, 'E', excelRowNum);

              itemCode = rawCode !== undefined ? String(rawCode).trim() : '';
              outDate = parseExcelDate(rawDate);
              qty = rawQty !== undefined ? Number(rawQty) || 0 : 0;
              customer = rawCust !== undefined ? String(rawCust).trim() : '그외';
            } else {
              const rawCode = getCellValue(worksheet, 'B', excelRowNum);
              const rawDate = getCellValue(worksheet, 'C', excelRowNum);
              const rawQty = getCellValue(worksheet, 'D', excelRowNum);

              itemCode = rawCode !== undefined ? String(rawCode).trim() : '';
              outDate = parseExcelDate(rawDate);
              qty = rawQty !== undefined ? Number(rawQty) || 0 : 0;
              customer = trimmedSheetName;
            }

            if (itemCode && itemCode !== 'undefined' && itemCode !== '품목코드') {
              parsedResults.push({
                id: `cust-${trimmedSheetName}-${excelRowNum}`,
                sheetName: trimmedSheetName,
                itemCode,
                outDate,
                qty,
                customer
              });
            }
          }
        });

        setCustomerData(parsedResults);
        setItemMasterData(prev => syncMasterListWithAllItems(supplierData, parsedResults, prev));
        alert(`고객사.xlsx 로드 완료! (${parsedResults.length}건)`);
      } catch (err) {
        console.error(err);
        alert('고객사 파싱 에러');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleCodeMasterFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });

        const codeSheet = workbook.Sheets['품목코드'];
        const stockSheet = workbook.Sheets['전체개수'];

        if (!codeSheet) {
          alert("'품목코드' 시트를 찾을 수 없습니다.");
          return;
        }

        const stockMap = new Map();
        if (stockSheet) {
          const stockRange = XLSX.utils.decode_range(stockSheet['!ref'] || 'A1:Z1000');
          for (let R = 2; R <= stockRange.e.r; R++) {
            const excelRowNum = R + 1;
            const code = getCellValue(stockSheet, 'B', excelRowNum);
            const qty = getCellValue(stockSheet, 'C', excelRowNum);
            if (code && code !== '품목코드') {
              stockMap.set(String(code).trim(), Number(qty) || 0);
            }
          }
        }

        const codeRange = XLSX.utils.decode_range(codeSheet['!ref'] || 'A1:Z1000');
        const customerHeaderMap = [];
        for (let C = 3; C <= codeRange.e.c; C++) {
          const colLetter = XLSX.utils.encode_col(C);
          const headerName = getCellValue(codeSheet, colLetter, 2);
          if (headerName) {
            customerHeaderMap.push({
              colLetter,
              customerName: String(headerName).trim()
            });
          }
        }

        let parsedMaster = [];
        for (let R = 2; R <= codeRange.e.r; R++) {
          const excelRowNum = R + 1;
          const itemCode = getCellValue(codeSheet, 'B', excelRowNum);
          const futuresCode = getCellValue(codeSheet, 'C', excelRowNum);

          if (itemCode && itemCode !== '품목코드') {
            const cleanCode = String(itemCode).trim();
            const customerCodes = {};
            customerHeaderMap.forEach(h => {
              const codeVal = getCellValue(codeSheet, h.colLetter, excelRowNum);
              customerCodes[h.customerName] = codeVal ? String(codeVal).trim() : '-';
            });

            parsedMaster.push({
              id: `master-${excelRowNum}`,
              itemCode: cleanCode,
              futuresCode: futuresCode ? String(futuresCode).trim() : '-',
              currentStock: stockMap.get(cleanCode) ?? 0,
              customerCodes
            });
          }
        }

        const fullMasterList = syncMasterListWithAllItems(supplierData, customerData, parsedMaster);
        setItemMasterData(fullMasterList);
        alert(`코드/재고.xlsx 로드 완료! (총 ${fullMasterList.length}품목 동기화됨)`);
      } catch (err) {
        console.error(err);
        alert('코드 파싱 에러');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // --------------------------------------------------------------------------
  // 검색 로직
  // --------------------------------------------------------------------------
  const autoSuggestions = useMemo(() => {
    const query = searchTerm.trim().toLowerCase();
    if (!query) return [];
    return itemMasterData
      .filter(item => 
        item.itemCode.toLowerCase().includes(query) || 
        (item.futuresCode && item.futuresCode.toLowerCase().includes(query))
      )
      .slice(0, 5);
  }, [searchTerm, itemMasterData]);

  const getDeliveryCustomers = (item) => {
    if (!item) return [];
    const customersSet = new Set();
    Object.entries(item.customerCodes || {}).forEach(([custName, codeVal]) => {
      if (codeVal && codeVal !== '-' && codeVal !== 'undefined') customersSet.add(custName);
    });
    customerData.forEach(c => {
      if (c.itemCode === item.itemCode && c.customer) customersSet.add(c.customer);
    });
    return Array.from(customersSet);
  };

  const handleSelectSuggestion = (item) => {
    setSelectedSearchItem(item);
    setSearchTerm(item.itemCode);
    setShowDropdown(false);
  };

  const handleSearchInputChange = (e) => {
    const value = e.target.value;
    setSearchTerm(value);
    setShowDropdown(true);

    if (!value.trim()) {
      setSelectedSearchItem(null);
    } else {
      const exactMatch = itemMasterData.find(i => 
        i.itemCode.toLowerCase() === value.trim().toLowerCase() ||
        (i.futuresCode && i.futuresCode.toLowerCase() === value.trim().toLowerCase())
      );
      if (exactMatch) setSelectedSearchItem(exactMatch);
    }
  };

  // --------------------------------------------------------------------------
  // 부족예정 경고 계산 및 필터링
  // --------------------------------------------------------------------------
  const deficitAlerts = useMemo(() => {
    const alerts = [];
    const todayStr = getTodayString();

    itemMasterData.forEach(master => {
      const itemCode = master.itemCode;
      
      const itemOutbounds = customerData
        .filter(c => c.itemCode === itemCode)
        .sort((a, b) => a.outDate.localeCompare(b.outDate));

      const itemInbounds = supplierData
        .filter(s => s.itemCode === itemCode)
        .sort((a, b) => a.dueDate.localeCompare(b.dueDate));

      if (itemOutbounds.length === 0) return;

      let runningStock = master.currentStock;
      
      itemInbounds.forEach(inb => {
        if (inb.dueDate < todayStr) {
          runningStock += inb.qty;
        }
      });

      const processedInboundIds = new Set();

      itemOutbounds.forEach(out => {
        itemInbounds.forEach(inb => {
          if (!processedInboundIds.has(inb.id) && inb.dueDate >= todayStr && inb.dueDate <= out.outDate) {
            runningStock += inb.qty;
            processedInboundIds.add(inb.id);
          }
        });

        runningStock -= out.qty;

        let status = 'NORMAL';
        if (runningStock < 0) {
          status = 'RED';
        } else if (runningStock >= 0 && runningStock <= 20) {
          status = 'YELLOW';
        }

        if (status !== 'NORMAL' && out.outDate >= todayStr) {
          const futureInbounds = itemInbounds.filter(i => i.dueDate >= out.outDate && i.qty > 0);
          const delayedInbounds = itemInbounds.filter(i => i.dueDate < todayStr && i.qty > 0);
          const recentInbound = futureInbounds[0];

          alerts.push({
            id: `alert-${itemCode}-${out.outDate}-${out.customer}`,
            itemCode,
            customer: out.customer,
            currentStock: master.currentStock,
            deliveryDate: out.outDate,
            outQty: out.qty,
            status,
            expectedStock: runningStock,
            deficitQty: runningStock < 0 ? Math.abs(runningStock) : 0, 
            recentInboundText: recentInbound 
              ? `${recentInbound.dueDate} (+${recentInbound.qty}EA)` 
              : (delayedInbounds.length > 0 ? `지연건 ${delayedInbounds.length}건 존재` : '입고예정 없음'),
            futureInbounds,
            delayedInbounds
          });
        }
      });
    });

    return alerts;
  }, [itemMasterData, customerData, supplierData]);

  // 부족예정 경고가 발생한 고객사 목록 추출 (드롭다운 옵션용)
  const alertCustomers = useMemo(() => {
    const set = new Set();
    deficitAlerts.forEach(a => {
      if (a.customer && a.customer !== '-') set.add(a.customer);
    });
    return Array.from(set).sort();
  }, [deficitAlerts]);

  // 다중 필터링 적용 (상태 + 고객사 + 경고 목록 전용 품목 검색)
  const filteredAlerts = useMemo(() => {
    return deficitAlerts.filter(alert => {
      if (alertFilter === 'YELLOW' && alert.status !== 'YELLOW') return false;
      if (alertFilter === 'RED' && alert.status !== 'RED') return false;

      if (alertCustomerFilter !== 'ALL' && alert.customer !== alertCustomerFilter) return false;

      if (alertSearchTerm.trim()) {
        const query = alertSearchTerm.trim().toLowerCase();
        const master = itemMasterData.find(m => m.itemCode === alert.itemCode);
        const futuresCode = master ? master.futuresCode : '';
        
        const matchItemCode = alert.itemCode.toLowerCase().includes(query);
        const matchFuturesCode = futuresCode.toLowerCase().includes(query);

        if (!matchItemCode && !matchFuturesCode) return false;
      }

      return true;
    });
  }, [deficitAlerts, alertFilter, alertCustomerFilter, alertSearchTerm, itemMasterData]);

  const toggleCalendarRedHighlight = (targetDate) => {
    if (calendarRedHighlightDate === targetDate) {
      setCalendarRedHighlightDate(null);
    } else {
      setCalendarRedHighlightDate(targetDate);
    }
  };

  // --------------------------------------------------------------------------
  // 캘린더 계산
  // --------------------------------------------------------------------------
  const highlightedOutboundDatesSet = useMemo(() => {
    if (!selectedSearchItem) return new Set();
    const dates = customerData
      .filter(c => c.itemCode === selectedSearchItem.itemCode)
      .map(c => c.outDate);
    return new Set(dates);
  }, [selectedSearchItem, customerData]);

  const outboundsByDateMap = useMemo(() => {
    const map = new Map();
    customerData.forEach(c => {
      if (!map.has(c.outDate)) map.set(c.outDate, []);
      map.get(c.outDate).push(c);
    });
    return map;
  }, [customerData]);

  const calendarDays = useMemo(() => {
    const firstDayIndex = new Date(currentYear, currentMonth, 1).getDay();
    const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();

    const days = [];
    for (let i = 0; i < firstDayIndex; i++) {
      days.push({ day: null, dateStr: null });
    }
    for (let d = 1; d <= daysInMonth; d++) {
      const mStr = String(currentMonth + 1).padStart(2, '0');
      const dStr = String(d).padStart(2, '0');
      const dateStr = `${currentYear}-${mStr}-${dStr}`;

      const dateOutboundList = outboundsByDateMap.get(dateStr) || [];
      const isHighlighted = highlightedOutboundDatesSet.has(dateStr);
      const isRedAlertHighlighted = calendarRedHighlightDate === dateStr;

      days.push({
        day: d,
        dateStr,
        outboundList: dateOutboundList,
        isHighlighted,
        isRedAlertHighlighted
      });
    }
    return days;
  }, [currentYear, currentMonth, outboundsByDateMap, highlightedOutboundDatesSet, calendarRedHighlightDate]);

  const selectedDayOutbounds = useMemo(() => {
    if (!selectedCalendarDate) return [];
    return outboundsByDateMap.get(selectedCalendarDate) || [];
  }, [selectedCalendarDate, outboundsByDateMap]);

  const prevMonth = () => {
    if (currentMonth === 0) { setCurrentYear(p => p - 1); setCurrentMonth(11); }
    else { setCurrentMonth(p => p - 1); }
  };
  const nextMonth = () => {
    if (currentMonth === 11) { setCurrentYear(p => p + 1); setCurrentMonth(0); }
    else { setCurrentMonth(p => p + 1); }
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50">
      {/* Header */}
      <header className="bg-slate-900 text-white px-8 py-4 flex justify-between items-center shadow-md">
        <div className="flex items-center space-x-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center font-bold text-lg">F</div>
          <h1 className="text-xl font-bold tracking-wide">
            FUTURES ERP <span className="text-xs font-normal text-slate-400 pl-2">출고 & 부족예정 모니터링</span>
          </h1>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 p-8 max-w-[1600px] w-full mx-auto space-y-6">
        
        {/* Upload Bar */}
        <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center space-x-3">
            <span className="font-bold text-xs text-slate-500">엑셀 데이터 업로드:</span>
            <label className="cursor-pointer bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs px-3 py-1.5 rounded-lg shadow transition">
              <span>코드/재고.xlsx</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleCodeMasterFileUpload} className="hidden" />
            </label>

            <label className="cursor-pointer bg-slate-800 hover:bg-slate-700 text-white font-bold text-xs px-3 py-1.5 rounded-lg shadow transition">
              <span>공급사.xlsx</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleSupplierFileUpload} className="hidden" />
            </label>
            <span className="text-xs text-slate-400">({supplierData.length}건)</span>

            <label className="cursor-pointer bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs px-3 py-1.5 rounded-lg shadow transition">
              <span>고객사.xlsx</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleCustomerFileUpload} className="hidden" />
            </label>
            <span className="text-xs text-slate-400">({customerData.length}건)</span>

            <span className="text-xs font-bold text-emerald-600 bg-emerald-50 px-2.5 py-1 rounded-md ml-2 border border-emerald-200">
              총 {itemMasterData.length}개 품목 분석 가능
            </span>
          </div>
        </div>

        {/* TOP: SEARCH & RESULT CARD */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div ref={searchContainerRef} className="lg:col-span-7 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-3 relative">
            <label className="block text-sm font-bold text-slate-700">🔍 품목검색 (품목코드 또는 퓨처스코드)</label>
            <div className="relative">
              <input 
                type="text" 
                value={searchTerm}
                onChange={handleSearchInputChange}
                onFocus={() => setShowDropdown(true)}
                placeholder="품목명/코드를 입력하면 출고 날짜가 캘린더에 노란색으로 표시됩니다..." 
                className="w-full pl-4 pr-10 py-3 bg-slate-50 border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 font-bold text-slate-800"
              />
              {searchTerm && (
                <button onClick={() => { setSearchTerm(''); setSelectedSearchItem(null); }} className="absolute right-3 top-3 text-slate-400 hover:text-slate-600 font-bold">✕</button>
              )}
              {showDropdown && autoSuggestions.length > 0 && (
                <div className="absolute left-0 right-0 top-full mt-2 bg-white border border-slate-200 rounded-xl shadow-xl z-50 overflow-hidden divide-y divide-slate-100">
                  {autoSuggestions.map((item) => (
                    <div key={item.id} onClick={() => handleSelectSuggestion(item)} className="p-3 hover:bg-blue-50 cursor-pointer transition flex justify-between items-center">
                      <div>
                        <p className="font-bold text-xs text-slate-800">{item.itemCode}</p>
                        <p className="text-[11px] text-purple-600 font-mono">퓨처스코드: {item.futuresCode}</p>
                      </div>
                      <span className="text-[11px] font-semibold text-blue-600 bg-blue-50 px-2 py-0.5 rounded">선택</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>

          <div className="lg:col-span-5 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
            <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">검색 품목 정보</h3>
            {selectedSearchItem ? (
              <div className="space-y-3">
                <p className="text-xs font-bold text-purple-600">코드: {selectedSearchItem.futuresCode}</p>
                <h2 className="text-base font-extrabold text-slate-800">{selectedSearchItem.itemCode}</h2>
                <div className="grid grid-cols-2 gap-4 pt-2 border-t border-slate-100">
                  <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-100">
                    <span className="text-[11px] font-semibold text-slate-400">현재고</span>
                    <p className="text-lg font-black text-blue-600 mt-0.5">{selectedSearchItem.currentStock.toLocaleString()} EA</p>
                  </div>
                  <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-100">
                    <span className="text-[11px] font-semibold text-slate-400">납품 고객사</span>
                    <div className="flex flex-wrap gap-1 mt-1">
                      {getDeliveryCustomers(selectedSearchItem).length > 0 ? (
                        getDeliveryCustomers(selectedSearchItem).map(cust => (
                          <span key={cust} className="bg-white border text-slate-700 font-bold text-[10px] px-1.5 py-0.5 rounded">{cust}</span>
                        ))
                      ) : <span className="text-xs text-slate-400">-</span>}
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <div className="py-4 text-center text-slate-400 text-xs">검색 시 품목 정보 및 재고가 표시됩니다.</div>
            )}
          </div>
        </div>

        {/* CALENDAR & DEFICIT ALERT DASHBOARD */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* LEFT: CALENDAR (Col 7) */}
          <div className="lg:col-span-7 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex justify-between items-center">
              <h2 className="text-base font-black text-slate-800 flex items-center gap-2">
                <span>📅</span> 출고 예정 캘린더 <span className="text-sm font-bold text-blue-600">({currentYear}년 {currentMonth + 1}월)</span>
              </h2>
              <div className="flex space-x-2">
                <button onClick={prevMonth} className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-lg transition">◀ 이전달</button>
                <button onClick={nextMonth} className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-lg transition">다음달 ▶</button>
              </div>
            </div>

            <div className="grid grid-cols-7 gap-2 text-center text-xs font-bold text-slate-400 py-1 border-b border-slate-100">
              <div className="text-red-500">일</div><div>월</div><div>화</div><div>수</div><div>목</div><div>금</div><div>토</div>
            </div>

            <div className="grid grid-cols-7 gap-2">
              {calendarDays.map((cell, idx) => {
                if (!cell.day) return <div key={idx} className="h-20 bg-slate-50/50 rounded-xl border border-transparent"></div>;

                const hasOutbound = cell.outboundList.length > 0;
                const dayOfWeek = (idx % 7);

                let bgStyle = 'bg-white border-slate-200';
                if (cell.isRedAlertHighlighted) {
                  bgStyle = 'bg-red-500/20 border-red-500 ring-2 ring-red-400 font-bold shadow-sm';
                } else if (cell.isHighlighted) {
                  bgStyle = 'bg-amber-200 border-amber-400 ring-2 ring-amber-300 font-bold shadow-sm';
                }

                return (
                  <div
                    key={cell.dateStr}
                    onClick={() => setSelectedCalendarDate(cell.dateStr)}
                    className={`h-20 p-2 rounded-xl border transition cursor-pointer flex flex-col justify-between hover:shadow-md ${bgStyle}`}
                  >
                    <div className="flex justify-between items-center">
                      <span className={`text-xs font-bold ${dayOfWeek === 0 ? 'text-red-500' : 'text-slate-700'}`}>{cell.day}</span>
                      {hasOutbound && <span className="w-1.5 h-1.5 rounded-full bg-blue-600"></span>}
                    </div>

                    {hasOutbound ? (
                      <span className="bg-blue-50 text-blue-700 font-extrabold text-[9px] px-1.5 py-0.5 rounded block truncate border border-blue-100">
                        출고 {cell.outboundList.length}건
                      </span>
                    ) : <span className="text-[9px] text-slate-300">-</span>}
                  </div>
                );
              })}
            </div>
          </div>

          {/* RIGHT: DEFICIT ALERT PANEL (Col 5) */}
          <div className="lg:col-span-5 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col space-y-4">
            
            {/* Header & Status Filter Tabs */}
            <div className="flex flex-wrap justify-between items-center gap-2 border-b border-slate-100 pb-3">
              <h2 className="text-base font-black text-slate-800 flex items-center gap-2">
                <span>⚠️</span> 부족예정 경고 
                <span className="text-xs font-bold text-slate-400">({filteredAlerts.length} / 전체 {deficitAlerts.length}건)</span>
              </h2>

              <div className="flex bg-slate-100 p-1 rounded-xl text-xs font-semibold">
                <button
                  onClick={() => setAlertFilter('ALL')}
                  className={`px-2.5 py-1 rounded-lg transition ${alertFilter === 'ALL' ? 'bg-white shadow text-slate-800 font-bold' : 'text-slate-500'}`}
                >
                  전체
                </button>
                <button
                  onClick={() => setAlertFilter('YELLOW')}
                  className={`px-2.5 py-1 rounded-lg transition ${alertFilter === 'YELLOW' ? 'bg-amber-400 text-slate-900 font-bold' : 'text-slate-500'}`}
                >
                  주의(0~20)
                </button>
                <button
                  onClick={() => setAlertFilter('RED')}
                  className={`px-2.5 py-1 rounded-lg transition ${alertFilter === 'RED' ? 'bg-red-500 text-white font-bold' : 'text-slate-500'}`}
                >
                  위험(&lt;0)
                </button>
              </div>
            </div>

            {/* Customer Dropdown & Item Search Bar */}
            <div className="grid grid-cols-1 sm:grid-cols-12 gap-2 bg-slate-50 p-2.5 rounded-xl border border-slate-200/80">
              {/* 고객사 선택 드롭다운 */}
              <div className="sm:col-span-5">
                <select
                  value={alertCustomerFilter}
                  onChange={(e) => setAlertCustomerFilter(e.target.value)}
                  className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-bold text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="ALL">🏢 모든 고객사 ({alertCustomers.length}곳)</option>
                  {alertCustomers.map(cust => (
                    <option key={cust} value={cust}>{cust}</option>
                  ))}
                </select>
              </div>

              {/* 경고 전용 품목코드 검색창 */}
              <div className="sm:col-span-7 relative">
                <input
                  type="text"
                  value={alertSearchTerm}
                  onChange={(e) => setAlertSearchTerm(e.target.value)}
                  placeholder="경고 항목 내 품목코드 검색..."
                  className="w-full bg-white border border-slate-300 rounded-lg pl-2.5 pr-7 py-1.5 text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 placeholder:font-normal placeholder:text-slate-400"
                />
                {alertSearchTerm && (
                  <button
                    onClick={() => setAlertSearchTerm('')}
                    className="absolute right-2 top-1.5 text-slate-400 hover:text-slate-600 text-xs font-bold"
                  >
                    ✕
                  </button>
                )}
              </div>
            </div>

            {/* Alert List */}
            <div className="space-y-3 max-h-[480px] overflow-y-auto pr-1">
              {filteredAlerts.length === 0 ? (
                <div className="py-16 text-center text-slate-400 text-xs">
                  {deficitAlerts.length === 0 
                    ? '오늘 이후 예정된 부족 항목이 없습니다.' 
                    : '검색 및 필터 조건에 일치하는 부족 항목이 없습니다.'}
                </div>
              ) : (
                filteredAlerts.map(alert => {
                  const isRed = alert.status === 'RED';
                  const isExpanded = expandedAlertId === alert.id;
                  const isCalendarActive = calendarRedHighlightDate === alert.deliveryDate;

                  return (
                    <div key={alert.id} className="border border-slate-200/80 rounded-2xl overflow-hidden shadow-sm bg-slate-50/50">
                      <div className="flex items-stretch">
                        <div className={`w-3 flex-shrink-0 ${isRed ? 'bg-red-500' : 'bg-amber-400'}`}></div>

                        <div className="p-3.5 flex-1 min-w-0 space-y-2">
                          <div className="flex justify-between items-start gap-2">
                            <h4 
                              onClick={() => setExpandedAlertId(isExpanded ? null : alert.id)}
                              className="font-bold text-xs text-slate-800 truncate cursor-pointer hover:text-blue-600"
                            >
                              {alert.itemCode}
                            </h4>
                            <span className="text-[10px] font-bold bg-white border text-slate-600 px-2 py-0.5 rounded-full whitespace-nowrap">
                              {alert.customer}
                            </span>
                          </div>

                          <div className="text-xs text-slate-600 grid grid-cols-2 gap-x-2 gap-y-1">
                            <div>
                              현재고: <span className="font-bold text-slate-800">{alert.currentStock}EA</span>
                            </div>
                            <div>납품일자: <span className="font-bold text-red-600">{alert.deliveryDate}</span></div>
                            <div className="col-span-2 text-[11px]">
                              출고시 예상재고: {' '}
                              <span className={`font-extrabold ${isRed ? 'text-red-600' : 'text-amber-600'}`}>
                                {alert.expectedStock}EA
                              </span>
                              {alert.deficitQty > 0 && (
                                <span className="text-red-600 font-extrabold ml-1">
                                  ({alert.deficitQty}EA 부족)
                                </span>
                              )}
                            </div>
                            <div className="col-span-2 text-[11px] text-slate-500">
                              최근 입고예정: <span className="font-semibold text-slate-700">{alert.recentInboundText}</span>
                            </div>
                          </div>

                          <div className="pt-2 border-t border-slate-200/60 flex justify-between items-center text-xs">
                            <button
                              onClick={() => toggleCalendarRedHighlight(alert.deliveryDate)}
                              className={`text-[11px] font-bold px-2 py-1 rounded-lg transition flex items-center gap-1 ${
                                isCalendarActive 
                                  ? 'bg-red-600 text-white shadow-sm' 
                                  : 'text-red-600 hover:text-red-800 bg-red-50'
                              }`}
                            >
                              <span>{isCalendarActive ? '📍 캘린더 표시 끄기' : '📍 캘린더로 보기'}</span>
                            </button>

                            <button
                              onClick={() => setExpandedAlertId(isExpanded ? null : alert.id)}
                              className="text-[11px] font-bold text-blue-600 hover:text-blue-800 flex items-center gap-0.5"
                            >
                              <span>{isExpanded ? '▲ 닫기' : '▼ 입고일정 상세'}</span>
                            </button>
                          </div>

                          {/* 입고일정 상세 영역 */}
                          {isExpanded && (
                            <div className="mt-2 pt-2 border-t border-slate-200 text-xs bg-white p-3 rounded-xl space-y-2.5 shadow-inner">
                              {alert.delayedInbounds && alert.delayedInbounds.length > 0 && (
                                <div className="space-y-1 bg-red-50/60 p-2 rounded-lg border border-red-100">
                                  <p className="font-bold text-red-600 text-[11px] flex items-center gap-1">
                                    <span>⚠️</span> 미입고 지연건 (납기일 경과):
                                  </p>
                                  {alert.delayedInbounds.map((inb, i) => (
                                    <div key={`del-${i}`} className="flex justify-between text-[11px] text-red-700 border-b border-red-100/50 pb-0.5 last:border-none">
                                      <span>{inb.dueDate} ({inb.supplier})</span>
                                      <span className="font-bold">+{inb.qty}EA (미입고)</span>
                                    </div>
                                  ))}
                                </div>
                              )}

                              <div className="space-y-1">
                                <p className="font-bold text-slate-700 text-[11px]">
                                  📦 납품일({alert.deliveryDate}) 기준 입고 예정 목록:
                                </p>
                                {alert.futureInbounds.length > 0 ? (
                                  alert.futureInbounds.map((inb, i) => (
                                    <div key={`fut-${i}`} className="flex justify-between text-[11px] text-slate-600 border-b border-slate-50 pb-1">
                                      <span>{inb.dueDate} ({inb.supplier})</span>
                                      <span className="font-bold text-emerald-600">+{inb.qty}EA</span>
                                    </div>
                                  ))
                                ) : (
                                  <p className="text-slate-400 text-[11px]">납품일 이후 예정된 입고 일정이 없습니다.</p>
                                )}
                              </div>
                            </div>
                          )}

                        </div>
                      </div>
                    </div>
                  );
                })
              )}
            </div>

          </div>

        </div>

      </main>

      {/* DAY OUTBOUND DETAIL MODAL */}
      {selectedCalendarDate && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full p-6 relative">
            <button onClick={() => setSelectedCalendarDate(null)} className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 font-bold text-xl">✕</button>
            <h3 className="text-base font-extrabold text-slate-800 mb-1 flex items-center gap-2"><span>📅</span> {selectedCalendarDate} 출고 상세</h3>
            <p className="text-xs text-slate-400 mb-4">해당 일자에 예정된 출고 목록입니다.</p>
            <div className="space-y-3 max-h-[350px] overflow-y-auto pr-1">
              {selectedDayOutbounds.length === 0 ? (
                <div className="py-8 text-center text-slate-400 text-xs">출고 내역이 없습니다.</div>
              ) : (
                selectedDayOutbounds.map((item, idx) => (
                  <div key={idx} className="p-3.5 bg-slate-50 border border-slate-200/80 rounded-xl flex justify-between items-center">
                    <div>
                      <span className="bg-blue-100 text-blue-800 text-[10px] font-bold px-2 py-0.5 rounded">고객사: {item.customer}</span>
                      <h4 className="font-bold text-xs text-slate-800 mt-1">{item.itemCode}</h4>
                    </div>
                    <div className="text-right">
                      <span className="text-base font-black text-blue-600 font-mono">{item.qty.toLocaleString()}</span>
                      <span className="text-xs font-semibold text-slate-500 pl-0.5">EA</span>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const rootElement = document.getElementById('root');
if (rootElement) {
  ReactDOM.createRoot(rootElement).render(<ERPApp />);
}