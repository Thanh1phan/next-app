'use client'

import React, { useEffect, useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, CheckCircle, X, Check, Download } from 'lucide-react';
import { ExcelViewer } from '@/components/excel-viewer';
import { Input } from '@heroui/input';
import { Button } from '@heroui/button';
import { Table, TableBody, TableCell, TableColumn, TableHeader, TableRow, useDisclosure, Modal, ModalBody, ModalContent, ModalFooter, ModalHeader } from '@heroui/react';

interface Field {
    fieldName: string;
    nameDisplay: string;
    type: 'number' | 'string' | 'date' | 'bool';
    isRequired: boolean;
    columnPosition: number;
    rowPosition: number;
    sheetName: string;
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
    field: string;
}

interface TryCastResult<T> {
    success: boolean;
    value: T | null;
    error?: string;
}

interface SubmitResult<T = any> {
    isSuccess: boolean;
    data: T;
    cellsErr: CellError[];
}

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
                let date: Date | null = null;

                // Xử lý Excel serial date (nếu value là số nguyên dương)
                if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
                    const excelBaseDate = new Date(1899, 11, 30);
                    date = new Date(excelBaseDate.getTime() + value * 86400000);
                }
                // Xử lý string dạng dd/MM/yyyy (hoặc dd-MM-yyyy)
                else if (typeof value === 'string') {
                    const ddmmyyyyRegex = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/;
                    const match = value.match(ddmmyyyyRegex);
                    if (match) {
                        const day = parseInt(match[1], 10);
                        const month = parseInt(match[2], 10);
                        const year = parseInt(match[3], 10);
                        date = new Date(year, month - 1, day);
                    }
                }

                // Fallback: Sử dụng new Date(value) cho các định dạng khác
                if (!date || isNaN(date.getTime())) {
                    date = new Date(value);
                }

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
};

interface ExcelImporterProps {
    configId: number;
    name: string;
}

export default function ExcelImporter({ configId, name }: ExcelImporterProps) {
    const [file, setFile] = useState<File | null>(null);
    const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
    const [selectedSheet, setSelectedSheet] = useState('');
    const [fields, setFields] = useState<Field[]>([]);
    const [dataStartCells, setDataStartCells] = useState<DataStartCell[]>([]);
    const [extractedData, setExtractedData] = useState<Record<string, any>[]>([]);
    const [previewData, setPreviewData] = useState<Record<string, any>[]>([]);
    const { isOpen: isImportOpen, onOpen: onImportOpen, onOpenChange: onImportOpenChange } = useDisclosure();
    const { isOpen: isPreviewOpen, onOpen: onPreviewOpen, onOpenChange: onPreviewOpenChange } = useDisclosure();
    const [cellError, setCellError] = useState<CellError[]>([]);
    const [loading, setLoading] = useState(false);
    const [sheetsConfigured, setSheetsConfigured] = useState<Set<string>>(new Set());

    const mapDataType = (dbType: string): Field['type'] => {
        switch (dbType) {
            case 'Integer':
            case 'Float':
                return 'number';
            case 'Date':
                return 'date';
            case 'Boolean':
                return 'bool';
            case 'String':
            default:
                return 'string';
        }
    };

    const fetchConfig = async () => {
        setLoading(true);
        try {
            const res = await fetch(`https://localhost:7034/excel-config/${configId}/details`);
            if (!res.ok) throw new Error('Lỗi fetch config details');
            const details = await res.json();
            const mappedFields: Field[] = details.map((d: any) => ({
                fieldName: d.fieldName,
                nameDisplay: d.displayName,
                type: mapDataType(d.dataType),
                isRequired: d.isRequired,
                columnPosition: d.columnPosition,
                rowPosition: d.rowPosition,
                sheetName: d.sheetName
            }));
            setFields(mappedFields);
            setDataStartCells(mappedFields.map((f: Field) => ({
                row: f.rowPosition,
                col: f.columnPosition,
                sheet: f.sheetName,
                field: f.fieldName,
            })));

            setSheetsConfigured(new Set(mappedFields.map(m => m.sheetName)));
        } catch (error) {
            alert('Lỗi fetch config: ' + (error as Error).message);
        } finally {
            setLoading(false);
        }
    };

    const filterVisibleWorkbook = (wb: XLSX.WorkBook): XLSX.WorkBook => {
        const visibleSheetNames = wb.SheetNames.filter(name => {
            const sheetMeta = wb.Workbook?.Sheets?.find(s => s.name === name);
            return !sheetMeta || sheetMeta.Hidden === 0;
        });

        const newWb = XLSX.utils.book_new();

        visibleSheetNames.forEach(name => {
            XLSX.utils.book_append_sheet(newWb, wb.Sheets[name], name);
        });

        return newWb;
    };

    const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const uploadedFile = e.target.files?.[0];
        if (!uploadedFile) return;

        setFile(uploadedFile);
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                const wbRaw = XLSX.read(event.target?.result, { type: 'binary' });
                const wb = filterVisibleWorkbook(wbRaw);
                setWorkbook(wb);
                const firstSheet = wb.SheetNames[0];
                setSelectedSheet(firstSheet);
            } catch (error) {
                alert('Lỗi khi đọc file Excel: ' + (error as Error).message);
            }
        };

        reader.readAsBinaryString(uploadedFile);
    };

    const checkRequiredFields = () => {
        const mappedFields = new Set(dataStartCells.map(d => d.field));
        return fields.filter(f => f.isRequired).every(f => mappedFields.has(f.fieldName));
    };

    const extractDataWithConfig = (
        workbook: XLSX.WorkBook,
        dataStartCells: DataStartCell[],
        fields: Field[]
    ): SubmitResult<Record<string, any>[]> => {
        const result: Record<string, any>[] = [];
        let isSuccess: boolean = true;
        let cellsErr: CellError[] = [];

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
                const fieldName = dataStartCells[colIdx].field;
                const field = fields.find(f => f.fieldName === fieldName);

                let res = tryCast(value, field?.type ?? 'string');
                if (!res.success) {
                    const startCol = dataStartCells[colIdx].col;
                    const startRow = dataStartCells[colIdx].row;
                    const sheet = dataStartCells[colIdx].sheet;
                    const mess = res.error ?? '';
                    cellsErr.push({
                        col: startCol,
                        row: startRow + rowIdx,
                        sheet: sheet,
                        index: colIdx,
                        err: mess
                    });
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

        return { isSuccess, data: result, cellsErr };
    };

    const handleExtract = () => {
        if (!workbook || fields.length === 0) return;

        if (!checkRequiredFields()) {
            alert('Config thiếu trường bắt buộc!');
            return;
        }

        const result = extractDataWithConfig(workbook, dataStartCells, fields);
        setExtractedData(result.data);
        setPreviewData(result.data.map((d, i) => ({
            key: `key_${i}`,
            stt: i + 1,
            ...d
        })));

        if (!result.isSuccess) {
            setCellError(result.cellsErr);
        } else {
            setCellError([]);
        }

        onPreviewOpen();
    };

    const getCellStyle = (rowIdx: number, colIdx: number, sheet: string) => {
        if (cellError.some(cell => cell.row === rowIdx && cell.col === colIdx && cell.sheet === sheet)) {
            return 'bg-red-500 text-white font-semibold border-2 border-blue-500';
        }
        if (dataStartCells.some(cell => cell.row === rowIdx && cell.col === colIdx && cell.sheet === sheet)) {
            return 'bg-blue-200 border-2 border-gray-400';
        }
        return 'bg-white';
    };

    const resetFile = () => {
        setFile(null);
        setWorkbook(null);
        setSelectedSheet('');
        setCellError([]);
        setPreviewData([]);
        setFields([]);
        setDataStartCells([]);
    };

    useEffect(() => {
        if (workbook) {
            fetchConfig();
        }
    }, [workbook]);

    return (
        <>
            <Button onPress={onImportOpen}>
                <Upload className="mx-auto" size={24} /> {name}
            </Button>

            <Modal
                size='5xl'
                isOpen={isImportOpen}
                onOpenChange={onImportOpenChange}
                onClose={() => resetFile()}
                isDismissable={false}
                isKeyboardDismissDisabled={true}
            >
                <ModalContent>
                    {(onClose) => (
                        <>
                            <ModalHeader>Import Excel</ModalHeader>
                            <ModalBody>
                                {!file ? (
                                    <div className="border-2 border-dashed border-gray-300 rounded-lg p-12 text-center hover:border-blue-400 transition-colors">
                                        <Upload className="mx-auto mb-4" size={48} />
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
                                )}

                                {workbook && (
                                    <ExcelViewer
                                        workbook={workbook}
                                        selectedSheet={selectedSheet}
                                        onSheetChange={setSelectedSheet}
                                        getCellClassName={getCellStyle}
                                        readOnly={true}
                                        sheetConfigured={sheetsConfigured}
                                    />
                                )}

                                {loading && <p>Đang tải config...</p>}
                            </ModalBody>
                            <ModalFooter>
                                <Button color="danger" onPress={onClose}>
                                    Đóng
                                </Button>
                                {file && (<Button color="success" onPress={handleExtract}>
                                    Trích xuất
                                </Button>)}
                            </ModalFooter>
                        </>
                    )}
                </ModalContent>
            </Modal>

            <Modal
                size='5xl'
                isOpen={isPreviewOpen}
                onOpenChange={onPreviewOpenChange}
                isDismissable={false}
                isKeyboardDismissDisabled={true}
            >
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
                                    <TableHeader columns={[
                                        { key: 'stt', label: 'STT' },
                                        ...fields.map(f => ({ key: f.fieldName, label: f.nameDisplay }))
                                    ]}>
                                        {(column) => <TableColumn key={column.key}>{column.label}</TableColumn>}
                                    </TableHeader>
                                    <TableBody items={previewData}>
                                        {(item) => (
                                            <TableRow key={item.key}>
                                                {(columnKey) => (
                                                    <TableCell>{item[columnKey as keyof typeof item]}</TableCell>
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
        </>
    );
}