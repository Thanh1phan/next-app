'use client'

import React, { useState, useRef, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, CheckCircle, X, Settings, Check, Download } from 'lucide-react';
import { ExcelViewer } from '@/components/excel-viewer';
import { Input } from '@heroui/input';
import { Button } from '@heroui/button';
import { Card, NumberInput, Select, SelectItem } from '@heroui/react';

interface HeaderMapping {
    col: number;
    originalValue: string;
    displayName: string;
    rowIndex: number;
    sheet: string;
}

interface DataStartCell {
    row: number;
    col: number;
    sheet: string;
    field?: string;
}

interface Table {
    tableName: string;
    fields: Field[];
}

interface Field {
    fieldName: string;
    type: 'number' | 'string' | 'date' | 'bool';
}

interface FieldSelected {
    field: string;
    isSelected: boolean;
}

type Step = 'select_mode' | 'select_headers' | 'select_data_start' | 'configure';

export default function ExcelImporter() {
    const [file, setFile] = useState<File | null>(null);
    const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
    const [selectedSheet, setSelectedSheet] = useState('');

    const [hasHeader, setHasHeader] = useState<boolean | null>(null);
    const [step, setStep] = useState<Step>('select_mode');
    const [headerMappings, setHeaderMappings] = useState<HeaderMapping[]>([]);
    const [selectedHeaderCells, setSelectedHeaderCells] = useState<Set<string>>(new Set());
    const [dataStartCells, setDataStartCells] = useState<DataStartCell[]>([]);
    const [sheetsConfigured, setsheetsConfigured] = useState<Set<string>>(new Set());
    const [table, setTable] = useState<string>();
    const [fields, setFileds] = useState<FieldSelected[]>([]);

    const tables: Table[] = ([
        {
            tableName: 'Asset',
            fields: [{ fieldName: 'Name', type: 'string' },
            { fieldName: 'Code', type: 'string' },
            { fieldName: 'OriginalPrice', type: 'number' }]
        },
        {
            tableName: 'User',
            fields: [{ fieldName: 'Name', type: 'string' },
            { fieldName: 'DateOfBirth', type: 'date' }]
        },
    ])

    const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const uploadedFile = e.target.files?.[0];
        if (!uploadedFile) return;

        setFile(uploadedFile);
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                const wb = XLSX.read(event.target?.result, { type: 'binary' });
                setWorkbook(wb);
                setSelectedSheet(wb.SheetNames[0]);
            } catch (error) {
                alert('Lỗi khi đọc file Excel: ' + (error as Error).message);
            }
        };

        reader.readAsBinaryString(uploadedFile);
    };

    const resetConfiguration = () => {
        setHasHeader(null);
        setStep('select_mode');
        setHeaderMappings([]);
        setSelectedHeaderCells(new Set());
        setDataStartCells([]);
        setsheetsConfigured(new Set());
        // setTable(undefined);
        // setFileds([]);
    };

    const handleSelectMode = (withHeader: boolean) => {
        setHasHeader(withHeader);
        setStep(withHeader ? 'select_headers' : 'select_data_start');
    };

    const handleCellClick = (rowIdx: number, colIdx: number, sheet: string) => {

        if (step === 'select_headers') {
            const cellKey = `${rowIdx}-${colIdx}-${sheet}`;
            const newSelected = new Set(selectedHeaderCells);

            if (newSelected.has(cellKey)) {
                newSelected.delete(cellKey);
                setHeaderMappings(prev => prev.filter(m => !(m.col === colIdx && m.rowIndex === rowIdx && m.sheet === sheet)));
                setDataStartCells(prev => prev.filter(m => !(m.col === colIdx && m.row === rowIdx + 1 && m.sheet === sheet)));
            } else {
                if (dataStartCells.length >= fields.length) {
                    return
                }
                newSelected.add(cellKey);
                const worksheet = workbook?.Sheets[sheet];
                const data = XLSX.utils.sheet_to_json(worksheet!, { header: 1, defval: '' }) as any[][];
                const cellValue = data[rowIdx]?.[colIdx] || '';

                setHeaderMappings(prev => [...prev, {
                    col: colIdx,
                    originalValue: cellValue.toString(),
                    displayName: cellValue.toString(),
                    rowIndex: rowIdx,
                    sheet: sheet
                }]);
                setDataStartCells(prev => [...prev, { row: rowIdx + 1, col: colIdx, sheet: sheet }]);
            }

            setSelectedHeaderCells(newSelected);
        } else if (step === 'select_data_start') {
            const existingIndex = dataStartCells.findIndex(cell => cell.row === rowIdx && cell.col === colIdx && cell.sheet === sheet);

            if (existingIndex !== -1) {
                setDataStartCells(prev => prev.filter((_, idx) => idx !== existingIndex));
            } else {
                if (dataStartCells.length >= fields.length) {
                    return
                }
                setDataStartCells(prev => [...prev, { row: rowIdx, col: colIdx, sheet: sheet }]);
            }
        }
        setsheetsConfigured(prev => new Set([...prev, sheet]));
    };

    const getCellStyle = (rowIdx: number, colIdx: number, sheet: string) => {
        if (step === 'select_headers' && selectedHeaderCells.has(`${rowIdx}-${colIdx}-${sheet}`)) {
            return 'bg-green-200 font-bold border-2 border-green-500 cursor-pointer';
        }
        if (step === 'select_data_start' && dataStartCells.some(cell => cell.row === rowIdx && cell.col === colIdx && cell.sheet === sheet)) {
            return 'bg-blue-200 font-bold border-2 border-blue-500 cursor-pointer';
        }
        if (step === 'configure' && selectedHeaderCells.has(`${rowIdx}-${colIdx}-${sheet}`)) {
            return 'bg-green-200 font-bold border-2 border-green-500';
        }
        if (step === 'configure' && dataStartCells.some(cell => cell.row === rowIdx && cell.col === colIdx && cell.sheet === sheet)) {
            return 'bg-blue-200 font-semibold border-2 border-blue-500';
        }
        if (step === 'select_headers' || step === 'select_data_start') {
            return 'bg-white hover:bg-gray-100 cursor-pointer';
        }
        return 'bg-white';
    };

    const confirmHeaderSelection = () => {
        if (headerMappings.length === 0) {
            alert('Vui lòng chọn ít nhất một header!');
            return;
        }
        setStep('configure');
    };

    const confirmDataStartSelection = () => {
        if (dataStartCells.length === 0) {
            alert('Vui lòng chọn ít nhất một ô để bắt đầu lấy dữ liệu!');
            return;
        }
        setStep('configure');
    };

    const updateHeaderName = (col: number, rowIndex: number, newName: string, sheet: string) => {
        setHeaderMappings(prev => prev.map(m =>
            m.col === col && m.rowIndex === rowIndex && m.sheet === sheet ? { ...m, displayName: newName } : m
        ));
    };

    const updateDataStartRow = (col: number, newRow: number, sheet: string) => {
        setDataStartCells(prev => prev.map(m =>
            m.col === col && m.sheet === sheet ? { ...m, row: newRow - 1 } : m
        ));
    };

    const updateDataField = (col: number, sheet: string, newField?: string, oldField?: string) => {
        setDataStartCells(prev => prev.map(m =>
            m.col === col && m.sheet === sheet ? { ...m, field: newField } : m
        ));

        if (oldField) {
            setFileds(prev => prev.map(f => ({
                ...f, isSelected: f.field === oldField ? false : f.isSelected
            })));
        }

        if (newField) {
            setFileds(prev => prev.map(f => ({
                ...f, isSelected: f.field === newField ? true : f.isSelected
            })));
        }
    };

    const extractAndDownload = () => {
        if (!workbook) return;

        const wb = XLSX.utils.book_new();
        const extractedData: (string | number | boolean | null)[][] = [];

        // Chuyển Set<string> thành array
        const sheetList = [...sheetsConfigured]; // Array of sheet names

        if (sheetList.length === 0) {
            alert("Bạn chưa chọn sheet nào!");
            return;
        }

        // Lấy dữ liệu của từng sheet dạng array of arrays
        const sheetRowsData = sheetList.map(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) return [];

            const arr = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: ""
            }) as any[][];
            return arr;
        });

        // Tạo header => mỗi sheet 1 cột
        extractedData.push(sheetList);

        // Tìm số dòng lớn nhất trong tất cả sheet
        const maxRows = Math.max(...sheetRowsData.map(r => r.length));

        for (let i = 0; i < maxRows; i++) {
            const row: (string | number | boolean | null)[] = [];

            sheetRowsData.forEach(sheetData => {
                // Lấy dòng i của sheet, nếu không tồn tại thì trả [] để fill ''
                const values = sheetData[i] ?? [];
                // Lấy giá trị toàn row từ sheet đó
                // Nếu bạn chỉ cần 1 ô đặc thù, sẽ custom ở phần dưới
                row.push(...values);
            });

            extractedData.push(row);
        }

        // Tạo sheet tổng hợp
        const ws = XLSX.utils.aoa_to_sheet(extractedData);
        XLSX.utils.book_append_sheet(wb, ws, "Extracted Data");
        XLSX.writeFile(wb, `extracted_data_${Date.now()}.xlsx`);

        alert(`Extract thành công ${maxRows} dòng dữ liệu!`);
    };


    const resetFile = () => {
        setFile(null);
        setWorkbook(null);
        setSelectedSheet('');
        resetConfiguration();
        setsheetsConfigured(new Set());
    };

    return (
        <div className="min-h-screen p-6">
            <div className="max-w-7xl mx-auto">
                <div className="rounded-lg shadow-xl p-6">
                    <h1 className="text-3xl font-bold mb-6 flex items-center gap-3">
                        <FileSpreadsheet size={32} className="text-blue-600" />
                        Excel Data Extractor
                    </h1>

                    {!file ? (
                        <div className="border-2 border-dashed border-gray-300 rounded-lg p-12 text-center hover:border-blue-400 transition-colors">
                            <Upload className="mx-auto  mb-4" size={48} />
                            <label className="cursor-pointer">
                                <span className="text-lg hover:text-blue-600">
                                    Nhấp để chọn file Excel
                                </span>
                                <Input
                                    type="file"
                                    accept=".xlsx,.xls"
                                    onChange={handleFileChange}
                                    className="hidden"
                                />
                            </label>
                            <p className="text-sm text-gray-500 mt-2">Hỗ trợ định dạng .xlsx và .xls</p>
                        </div>
                    ) : (
                        <>
                            <div className="mb-6 flex items-center justify-between border border-blue-200 p-4 rounded-lg">
                                <div className="flex items-center gap-3">
                                    <CheckCircle className="text-green-600" />
                                    <div>
                                        <p className="font-semibold">{file.name}</p>
                                        <p className="text-sm">
                                            {(file.size / 1024).toFixed(2)} KB
                                        </p>
                                    </div>
                                </div>
                                <Button
                                    onClick={resetFile}
                                    color='danger'
                                    startContent={<X size={16} />}
                                >
                                    Xóa file
                                </Button>
                            </div>

                            {workbook && (
                                <div className="grid grid-cols-1">
                                    <div className="border-2 border-blue-200 rounded-lg p-5 shadow-md">
                                        <div className="flex items-center gap-2 mb-4">
                                            <Settings size={20} className="text-blue-600" />
                                            <h3 className="font-bold text-lg">Cấu hình</h3>
                                        </div>

                                        <div className="mb-4">
                                            <Select
                                                className="max-w-xs"
                                                label="Chọn table mapping"
                                                placeholder="Chọn table"
                                                defaultSelectedKeys={table}
                                                variant="bordered"

                                                onChange={(e) => {
                                                    setTable(e.target.value);
                                                    const tab = tables.find(tab => tab.tableName === e.target.value);
                                                    setFileds(tab?.fields.map<FieldSelected>(f => ({ field: f.fieldName, isSelected: false })) ?? []);
                                                    resetConfiguration();
                                                }}
                                            >
                                                {tables.map((tab) => (
                                                    <SelectItem key={tab.tableName}>{tab.tableName}</SelectItem>
                                                ))}
                                            </Select>
                                        </div>

                                        {step === 'select_mode' && table && (
                                            <div className="space-y-4">
                                                <p className="text-sm font-semibold mb-3">
                                                    Dữ liệu của bạn có header không?
                                                </p>
                                                <div className="grid grid-cols-2 gap-2">
                                                    <Button
                                                        onClick={() => handleSelectMode(true)}
                                                        color='success'
                                                    >
                                                        ✓ Có Header
                                                    </Button>
                                                    <Button
                                                        onClick={() => handleSelectMode(false)}
                                                    >
                                                        ✗ Không có Header
                                                    </Button>
                                                </div>
                                            </div>
                                        )}

                                        {step === 'select_headers' && (
                                            <div className="space-y-4">
                                                <div className="border-2 border-gray-200 p-3 rounded-lg">
                                                    <p className="text-sm">
                                                        <strong>Đã chọn:</strong> {headerMappings.length} header
                                                    </p>
                                                </div>
                                                <div className="grid grid-cols-2 gap-2">
                                                    <Button
                                                        onClick={confirmHeaderSelection}
                                                        disabled={headerMappings.length === 0}
                                                        className="flex items-center justify-center gap-1 px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
                                                    >
                                                        <Check size={16} />
                                                        Xác nhận
                                                    </Button>
                                                    <Button
                                                        onClick={() => { setStep('select_mode'); resetConfiguration(); }}
                                                        className="px-4 py-2 bg-gray-500 text-white rounded-lg hover:bg-gray-600 transition-colors"
                                                    >
                                                        ← Quay lại
                                                    </Button>
                                                </div>
                                            </div>
                                        )}

                                        {step === 'select_data_start' && (
                                            <div className="space-y-4">
                                                <div className="bg-amber-50 p-3 rounded-lg">
                                                    <p className="text-sm text-amber-800 font-semibold">
                                                        📌 Click vào các ô để chọn điểm bắt đầu
                                                    </p>
                                                    {dataStartCells.length > 0 && (
                                                        <p className="text-sm text-amber-700 mt-2">
                                                            Đã chọn: <strong>{dataStartCells.length} ô</strong>
                                                        </p>
                                                    )}
                                                </div>
                                                <div className="grid grid-cols-2 gap-2">
                                                    <button
                                                        onClick={confirmDataStartSelection}
                                                        disabled={dataStartCells.length === 0}
                                                        className="flex items-center justify-center gap-1 px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 disabled:bg-gray-300 transition-colors"
                                                    >
                                                        <Check size={16} />
                                                        Xác nhận
                                                    </button>
                                                    <button
                                                        onClick={() => { setStep('select_mode'); resetConfiguration(); }}
                                                        className="px-4 py-2 bg-gray-500 text-white rounded-lg hover:bg-gray-600 transition-colors"
                                                    >
                                                        ← Quay lại
                                                    </button>
                                                </div>
                                            </div>
                                        )}

                                        {step === 'configure' && (
                                            <div className="space-y-3">
                                                {hasHeader ? (
                                                    <div>
                                                        <h4 className="text-sm font-semibold mb-2">
                                                            Headers ({headerMappings.length}):
                                                        </h4>
                                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-96 overflow-y-auto p-2">
                                                            {headerMappings.sort((a, b) => {
                                                                if (a.sheet !== b.sheet) {
                                                                    return a.sheet.localeCompare(b.sheet);
                                                                }
                                                                if (a.col !== b.col) {
                                                                    return a.col - b.col;
                                                                }
                                                                return a.rowIndex - b.rowIndex;
                                                            }).map((mapping, idx) => (
                                                                <Card key={idx} className="p-2 space-y-2">
                                                                    <div className='grid grid-cols-2 justify-items-stretch mb-1'>
                                                                        <p className="text-xs">
                                                                            Cột {String.fromCharCode(65 + mapping.col)}
                                                                        </p>
                                                                        <p className='justify-self-end text-xs'>{mapping.sheet}</p>
                                                                    </div>
                                                                    <Input
                                                                        type="text"
                                                                        value={mapping.displayName}
                                                                        onChange={(e) => updateHeaderName(mapping.col, mapping.rowIndex, e.target.value, mapping.sheet)}
                                                                        label='Header:'
                                                                    />
                                                                    <NumberInput
                                                                        type="number"
                                                                        value={dataStartCells[idx].row + 1}
                                                                        onChange={(e) => updateDataStartRow(mapping.col, Number(e), mapping.sheet)}
                                                                        label='Dòng bắt đầu:'
                                                                    />
                                                                    <Select
                                                                        label="Trường mapping"
                                                                        placeholder="Chọn trường"
                                                                        defaultSelectedKeys={dataStartCells[idx].field}
                                                                        disabledKeys={fields.filter(f => f.isSelected).map(f => f.field)}
                                                                        onChange={(e) => {
                                                                            updateDataField(dataStartCells[idx].col, dataStartCells[idx].sheet, e.target.value);
                                                                        }}
                                                                    >
                                                                        {fields.map((f) => (
                                                                            <SelectItem key={f.field}>{f.field}</SelectItem>
                                                                        ))}
                                                                    </Select>
                                                                </Card>
                                                            ))}
                                                        </div>
                                                    </div>
                                                ) : <div>
                                                    <h4 className="text-sm font-semibold mb-2">
                                                        Columns ({headerMappings.length}):
                                                    </h4>
                                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-96 overflow-y-auto p-2">
                                                        {dataStartCells.sort((a, b) => {
                                                            if (a.sheet !== b.sheet) {
                                                                return a.sheet.localeCompare(b.sheet);
                                                            }
                                                            if (a.col !== b.col) {
                                                                return a.col - b.col;
                                                            }
                                                            return a.row - b.row;
                                                        }).map((mapping, idx) => (
                                                            <Card key={idx} className="p-2 space-y-2">
                                                                <div className='grid grid-cols-2 justify-items-stretch mb-1'>
                                                                    <p className="text-xs">
                                                                        Cột {String.fromCharCode(65 + mapping.col)}
                                                                    </p>
                                                                    <p className='justify-self-end text-xs'>{mapping.sheet}</p>
                                                                </div>
                                                                <NumberInput
                                                                    type="number"
                                                                    value={mapping.row + 1}
                                                                    onChange={(e) => updateDataStartRow(mapping.col, Number(e), mapping.sheet)}
                                                                    label='Dòng bắt đầu:'
                                                                />
                                                                <Select
                                                                    label="Trường mapping"
                                                                    placeholder="Chọn trường"
                                                                    defaultSelectedKeys={mapping.field}
                                                                    disabledKeys={fields.filter(f => f.isSelected).map(f => f.field)}
                                                                    onChange={(e) => {
                                                                        updateDataField(mapping.col, mapping.sheet, e.target.value);
                                                                    }}
                                                                >
                                                                    {fields.map((f) => (
                                                                        <SelectItem key={f.field}>{f.field}</SelectItem>
                                                                    ))}
                                                                </Select>
                                                            </Card>
                                                        ))}
                                                    </div>
                                                </div>}

                                                <div className="grid grid-cols-2 gap-2 pt-3">
                                                    <button
                                                        onClick={extractAndDownload}
                                                        className="flex items-center justify-center gap-1 px-4 py-2 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-colors"
                                                    >
                                                        <Download size={16} />
                                                        Extract
                                                    </button>
                                                    <button
                                                        onClick={() => { setStep('select_mode'); resetConfiguration(); }}
                                                        className="px-4 py-2 bg-gray-500 text-white rounded-lg hover:bg-gray-600 transition-colors"
                                                    >
                                                        ← Cấu hình lại
                                                    </button>
                                                </div>
                                            </div>
                                        )}
                                    </div>

                                    {/* Excel Viewer */}
                                    <div className="lg:col-span-2">
                                        <ExcelViewer
                                            workbook={workbook}
                                            selectedSheet={selectedSheet}
                                            onSheetChange={setSelectedSheet}
                                            onCellClick={handleCellClick}
                                            getCellClassName={getCellStyle}
                                            readOnly={step === 'select_mode'}
                                            sheetConfigured={sheetsConfigured}
                                        />
                                    </div>
                                </div>
                            )}
                        </>
                    )}
                </div>
            </div>
        </div >
    );
}