'use client'

import React, { useState, useRef, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, CheckCircle, X, Settings, Check, Download } from 'lucide-react';
import { ExcelViewer } from '@/components/excel-viewer';
import { Input } from '@heroui/input';
import { Button } from '@heroui/button';
import { Card, Form, Modal, ModalBody, ModalContent, ModalFooter, ModalHeader, NumberInput, Select, SelectItem, Table, TableBody, TableCell, TableColumn, TableHeader, TableRow, useDisclosure } from '@heroui/react';

interface HeaderMapping {
    col: number;
    originalValue: string;
    displayName: string;
    rowIndex: number;
    sheet: string;
}

interface Cell {
    row: number;
    col: number;
    sheet: string;
}

interface CellError extends Cell {
    err?: string;
    index?: number;
}

interface DataStartCell extends Cell {
    field?: string;
}

interface Table {
    tableName: string;
    fields: Field[];
}

interface Field {
    fieldName: string;
    type: 'number' | 'string' | 'date' | 'bool';
    isSelected?: boolean;
    isRequired?: boolean;
}

interface TryCastResult<T> {
    success: boolean;
    value: T | null;
    error?: string;
}

interface SubmitResult<T = any> {
    isSuccess: boolean;
    data: T;
}

type Step = 'select_mode' | 'select_headers' | 'select_data_start' | 'configure';

const Tables: Table[] = [
    {
        tableName: 'Asset',
        fields: [{ fieldName: 'Name', type: 'string', isRequired: true },
        { fieldName: 'Code', type: 'string', isRequired: true },
        { fieldName: 'StartUsingDate', type: 'date' },
        { fieldName: 'OriginalPrice', type: 'number' }],
    },
    {
        tableName: 'User',
        fields: [{ fieldName: 'Name', type: 'string' },
        { fieldName: 'DateOfBirth', isRequired: true, type: 'date' }]
    },
]

const isEmptyValue = (value: any): boolean => {
    return value === null ||
        value === undefined ||
        value === "" ||
        (typeof value === 'string' && value.trim() === "");
};

const tryCast = (
    value: any,
    type: 'string' | 'number' | 'bool' | 'date'
): TryCastResult<any> => {

    if (value === undefined || value === null || value === '') {
        return { success: true, value: null };
    }

    try {
        switch (type) {
            case 'number': {
                const num = Number(value);
                if (isNaN(num)) {
                    return {
                        success: false,
                        value: null,
                        error: `"${value}" không phải là số`
                    };
                }
                return { success: true, value: num };
            }

            case 'bool': {
                if (typeof value === 'boolean') {
                    return { success: true, value };
                }

                if (value === 1 || value === '1' || value === 'true') {
                    return { success: true, value: true };
                }

                if (value === 0 || value === '0' || value === 'false') {
                    return { success: true, value: false };
                }

                return {
                    success: false,
                    value: null,
                    error: `"${value}" không phải boolean`
                };
            }

            case 'date': {
                const date = new Date(value);
                if (isNaN(date.getTime())) {
                    return {
                        success: false,
                        value: null,
                        error: `"${value}" không phải ngày hợp lệ`
                    };
                }
                return {
                    success: true,
                    value: date.toISOString()
                };
            }

            case 'string':
            default:
                return {
                    success: true,
                    value: value.toString()
                };
        }
    } catch (err) {
        return {
            success: false,
            value: null,
            error: (err as Error).message
        };
    }
}

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
    const [fields, setFileds] = useState<Field[]>([]);
    const [extractedData, setExtractedData] = useState<Record<string, any>[]>([]);
    const [previewData, setPreviewData] = useState<Record<string, any>[]>([]);
    const { isOpen, onOpen, onOpenChange } = useDisclosure();
    const [cellError, setCellError] = useState<CellError[]>([]);
    const [errors, setErrors] = useState<Record<string, string[]>>({});

    const tables: Table[] = (Tables)

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
        setCellError([]);
        setFileds(prev => prev.map(f => ({
            ...f,
            isSelected: false
        })));
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
        if (cellError.some(cell => cell.row === rowIdx && cell.col === colIdx && cell.sheet === sheet)) {
            return 'bg-red-600 font-semibold border-2 border-blue-500';
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

    const updateHeaderName = (index: number, newName: string) => {
        setHeaderMappings(prev => prev.map((item, idx) =>
            idx == index ? { ...item, displayName: newName } : item
        ));
    };

    const updateDataStartRow = (index: number, newRow: number) => {
        setDataStartCells(prev => prev.map((item, idx) =>
            idx == index ? { ...item, row: newRow - 1 } : item
        ));
    };

    const updateDataField = (index: number, newField?: string, oldField?: string) => {
        setDataStartCells(prev => prev.map((item, idx) =>
            idx == index ? { ...item, field: newField } : item
        ));

        if (oldField) {
            setFileds(prev => prev.map(f => ({
                ...f, isSelected: f.fieldName === oldField ? false : f.isSelected
            })));
        }

        if (newField) {
            setFileds(prev => prev.map(f => ({
                ...f, isSelected: f.fieldName === newField ? true : f.isSelected
            })));
        }
    };

    const checkRequiredFields = () => {
        const mappedFields = new Set(
            dataStartCells.map(d => d.field)
        );

        return fields
            .filter(f => f.isRequired)
            .every(f => mappedFields.has(f.fieldName));
    };

    const onSubmit = (e: React.FormEvent<HTMLFormElement>) => {
        e.preventDefault();
        if (!workbook) return;

        if (!checkRequiredFields()) {
            alert('Có trường bắt buộc chưa được mapping data!')
            return;
        }
        const result = extractDataWithConfig(workbook, dataStartCells, fields);
        setExtractedData(result.data);
        setPreviewData(result.data.map((d, i) => ({
            key: `key_${i}`,
            ...d
        })));
        if (!result.isSuccess) {
            const uniqueIndexes = new Set(cellError.map(c => c.index));

            const newErrors = Array.from(uniqueIndexes).reduce<Record<string, string[]>>(
                (acc, index) => addValidationError(acc, `field${index}`, 'Lỗi mapping kiểu dữ liệu'),
                {}
            );

            setErrors(newErrors);
        }
        onOpen();
    };

    const extractDataWithConfig = (
        workbook: XLSX.WorkBook,
        dataStartCells: DataStartCell[],
        fields: Field[]
    ): SubmitResult<Record<string, any>[]> => {
        const result: Record<string, any>[] = [];
        let isSuccess: boolean = true;

        const columnData: any[][] = dataStartCells.map(cfg => {
            const worksheet = workbook.Sheets[cfg.sheet];
            if (!worksheet) return [];

            const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null }) as any[][];
            const data: any[] = [];
            for (let i = cfg.row; i < sheetData.length; i++) {
                const row = sheetData[i] || [];
                data.push(row[cfg.col] ?? null);
            }
            return data;
        });

        const maxRows = Math.max(...columnData.map(col => col.length), 0);

        for (let rowIdx = 0; rowIdx < maxRows; rowIdx++) {
            const dataRow: Record<string, any> = {};
            let allEmpty = true;

            for (let colIdx = 0; colIdx < dataStartCells.length; colIdx++) {
                const value = columnData[colIdx][rowIdx] ?? null;
                const fieldName = dataStartCells[colIdx].field || `Column_${colIdx}`;
                const field = fields.find(f => f.fieldName === fieldName);

                let res = tryCast(value, field?.type ?? 'string')
                if (!res.success) {
                    const startCol = dataStartCells[colIdx].col;
                    const startRow = dataStartCells[colIdx].row;
                    const sheet = dataStartCells[colIdx].sheet;
                    const field = dataStartCells[colIdx].field;
                    const mess = res.error ?? '';
                    setCellError(prev => [...prev, { col: startCol, row: startRow + rowIdx, sheet: sheet, index: colIdx, err: mess }]);
                    isSuccess = false;
                }
                dataRow[fieldName] = res.success ? res.value : res.error;
                if (!isEmptyValue(value)) {
                    allEmpty = false;
                }
            }

            if (allEmpty) break;
            result.push(dataRow);
        }

        return { isSuccess: isSuccess, data: result };
    };

    const addValidationError = (
        errors: Record<string, string[]>,
        field: string,
        message: string
    ): Record<string, string[]> => {
        return {
            ...errors,
            [field]: [...(errors[field] ?? []), message]
        };
    }
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
                            <div className="mb-6 flex items-center justify-between border-2 border-blue-200 rounded-lg shadow-md p-4">
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
                                <div>
                                    <div className="border-2 border-blue-200 rounded-lg p-5 shadow-md">
                                        <div className="flex items-center gap-2 mb-4">
                                            <Settings size={20} className="text-blue-600" />
                                            <h3 className="font-bold text-lg">Cấu hình</h3>
                                        </div>

                                        <div className="mb-4">
                                            <Select
                                                className="max-w-xs"
                                                label="Chọn table"
                                                placeholder="Chọn table"
                                                variant="bordered"

                                                onChange={(e) => {
                                                    setTable(e.target.value);
                                                    const tab = tables.find(tab => tab.tableName === e.target.value);
                                                    setFileds([...tab?.fields ?? []]);
                                                    resetConfiguration();
                                                }}
                                            >
                                                {tables.map((tab) => (
                                                    <SelectItem key={tab.tableName}>{tab.tableName}</SelectItem>
                                                ))}
                                            </Select>
                                        </div>

                                        {table && step === 'select_mode' && table && (
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
                                                    <p className="text-sm  font-semibold">
                                                        📌 Click vào các ô để chọn header
                                                    </p>
                                                    <p className="text-sm mt-2">
                                                        <strong>Đã chọn:</strong> {headerMappings.length} / {fields.length} header
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
                                                <div className="border-2 border-gray-200 p-3 rounded-lg">
                                                    <p className="text-sm font-semibold">
                                                        📌 Click vào các ô để chọn điểm bắt đầu
                                                    </p>
                                                    <p className="text-sm mt-2">
                                                        <strong>Đã chọn:</strong> {dataStartCells.length} / {fields.length} ô
                                                    </p>
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
                                            <Form
                                                className='w-full grid grid-cols-1'
                                                onSubmit={onSubmit}
                                                validationErrors={errors}
                                            >
                                                {hasHeader ? (
                                                    <div>
                                                        <h4 className="text-sm font-semibold mb-2">
                                                            Headers ({headerMappings.length}):
                                                        </h4>
                                                        <div className="w-full grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2 max-h-96 overflow-y-auto p-2">
                                                            {headerMappings.map((mapping, idx) => (
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
                                                                        onChange={(e) => updateHeaderName(idx, e.target.value)}
                                                                        label='Header:'
                                                                    />
                                                                    <NumberInput
                                                                        type="number"
                                                                        value={dataStartCells[idx].row + 1}
                                                                        onChange={(e) => updateDataStartRow(idx, Number(e))}
                                                                        label='Dòng bắt đầu:'
                                                                        isRequired
                                                                        minValue={1}
                                                                    />
                                                                    <Select
                                                                        label="Trường"
                                                                        placeholder="Chọn trường"
                                                                        disabledKeys={fields.filter(f => f.isSelected && f.fieldName != dataStartCells[idx].field).map(f => f.fieldName)}
                                                                        onChange={(e) => {
                                                                            updateDataField(idx, e.target.value, dataStartCells[idx].field);
                                                                        }}
                                                                        isRequired
                                                                        name={'field' + idx}
                                                                    >
                                                                        {fields?.map((f) => (
                                                                            <SelectItem
                                                                                key={f.fieldName}
                                                                                textValue={f.fieldName}
                                                                            >
                                                                                {f.fieldName} ({f.type}) {f.isRequired && <span className='text-red-600'>*</span>}
                                                                            </SelectItem>
                                                                        ))}
                                                                    </Select>
                                                                </Card>
                                                            ))}
                                                        </div>
                                                    </div>
                                                ) : <div>
                                                    <h4 className="text-sm font-semibold mb-2">
                                                        Columns ({dataStartCells.length}):
                                                    </h4>
                                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-96 overflow-y-auto p-2">
                                                        {dataStartCells.map((mapping, idx) => (
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
                                                                    onChange={(e) => updateDataStartRow(idx, Number(e))}
                                                                    label='Dòng bắt đầu:'
                                                                />
                                                                <Select
                                                                    label="Trường mapping"
                                                                    placeholder="Chọn trường"
                                                                    disabledKeys={fields.filter(f => f.isSelected).map(f => f.fieldName)}
                                                                    onChange={(e) => {
                                                                        updateDataField(idx, e.target.value, mapping.field);
                                                                    }}
                                                                >
                                                                    {fields.map((f) => (
                                                                        <SelectItem key={f.fieldName}>{f.fieldName} {f.isRequired && '*'}</SelectItem>
                                                                    ))}
                                                                </Select>
                                                            </Card>
                                                        ))}
                                                    </div>
                                                </div>}

                                                <div className="grid grid-cols-2 gap-2 pt-3">
                                                    <button
                                                        type='submit'
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
                                            </Form>
                                        )}
                                    </div>

                                    {table && <div className="lg:col-span-2">
                                        <ExcelViewer
                                            workbook={workbook}
                                            selectedSheet={selectedSheet}
                                            onSheetChange={setSelectedSheet}
                                            onCellClick={handleCellClick}
                                            getCellClassName={getCellStyle}
                                            readOnly={step === 'select_mode'}
                                            sheetConfigured={sheetsConfigured}
                                        />
                                    </div>}
                                </div>
                            )}
                        </>
                    )}
                </div>
            </div>

            <Modal size='5xl' isOpen={isOpen} onOpenChange={onOpenChange}>
                <ModalContent>
                    {(onClose) => (
                        <>
                            <ModalHeader className="flex flex-col gap-1">Dữ liệu trích xuất</ModalHeader>
                            <ModalBody>
                                <Table
                                    aria-label="Table with dynamic content"
                                    maxTableHeight={400}
                                    isVirtualized
                                >
                                    <TableHeader columns={fields.map(f => ({ key: f.fieldName, label: f.fieldName }))}>
                                        {(column) => <TableColumn key={column.key}>{column.label}</TableColumn>}
                                    </TableHeader>
                                    <TableBody items={previewData}>
                                        {(item) => (
                                            <TableRow key={item.key}>
                                                {(columnKey) => (
                                                    <TableCell>{item[columnKey]}</TableCell>
                                                )}
                                            </TableRow>
                                        )}
                                    </TableBody>
                                </Table>
                            </ModalBody>
                            <ModalFooter>
                                <Button color="danger" onPress={onClose}>
                                    Đóng
                                </Button>
                                {cellError.length === 0 &&
                                    <Button color="primary" onPress={() => {
                                        console.log(extractedData);
                                        onClose();
                                        alert('Xuất thành công!')
                                    }}>
                                        Xuất
                                    </Button>}
                            </ModalFooter>
                        </>
                    )}
                </ModalContent>
            </Modal>
        </div >
    );
}