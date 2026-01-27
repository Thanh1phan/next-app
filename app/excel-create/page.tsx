'use client'

import { ExcelViewer } from "@/components/excel-viewer";
import { Button, Input, Select, SelectItem, Table, TableHeader, TableColumn, TableBody, TableRow, TableCell, Checkbox, Modal, ModalContent, ModalHeader, ModalBody, ModalFooter, useDisclosure, NumberInput, Form, Card } from "@heroui/react";
import { RefreshCw, Save, Plus, Trash2, Edit, Upload, CheckCircle, X, Settings, Check, Download, View } from "lucide-react";
import { ChangeEvent, useEffect, useState } from "react";
import * as XLSX from 'xlsx';

// Enums
enum DataTypes {
    String = 0,
    Number = 1,
    Date = 2,
    Boolean = 3,
    Decimal = 4
}

enum ConfigType {
    Salary = 0,
    Insurance = 1
}

enum DepartmentId {
    DepartmentA = 1,
    DepartmentB = 2
}

// Types
interface ExcelConfigDetail {
    id?: number;
    configId?: number;
    fieldName: string;
    displayName: string;
    columnPosition: number;
    rowPosition: number;
    sheetName: string;
    dataType: DataTypes;
    isRequired: boolean;
}

interface ExcelConfig {
    id: number;
    templateFileName: string;
    configName: string;
    departmentId: number;
    configType: ConfigType;
    details?: ExcelConfigDetail[];
    acctions: string;
}

const API_BASE_URL = 'https://localhost:7034';

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
    nameDisplay: string;
    type: DataTypes;
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
    cellsErr: CellError[];
}

type Step = 'select_mode' | 'select_headers' | 'set_row_start' | 'select_data_start' | 'configure';

const Tables: Field[] = [
    { fieldName: 'fullName', nameDisplay: 'Họ và tên', type: 0, isRequired: false },
    { fieldName: 'ctvCode', nameDisplay: 'Mã CTV', type: 0, isRequired: false },

    { fieldName: 'firstContractStartDateT9_2024', nameDisplay: 'Ngày bắt đầu hợp đồng lần (T9/2024)', type: 2, isRequired: false },
    { fieldName: 'contractStartDate', nameDisplay: 'Ngày bắt đầu hợp đồng', type: 2, isRequired: false },

    { fieldName: 'organization', nameDisplay: 'Đơn vị', type: 0, isRequired: false },
    { fieldName: 'jobPosition', nameDisplay: 'Vị trí công việc', type: 0, isRequired: false },

    { fieldName: 'actualWorkingDays', nameDisplay: 'Ngày công thực tế', type: 1, isRequired: false },
    { fieldName: 'leaveDays', nameDisplay: 'Ngày công phép', type: 1, isRequired: false },
    { fieldName: 'holidayDays', nameDisplay: 'Ngày công lễ', type: 1, isRequired: false },
    { fieldName: 'nightShiftDays', nameDisplay: 'Ngày công ca đêm', type: 1, isRequired: false },
    { fieldName: 'policyLeaveDays', nameDisplay: 'Nghỉ chế độ', type: 1, isRequired: false },
    { fieldName: 'bhxhLeaveDays', nameDisplay: 'Nghỉ BHXH', type: 1, isRequired: false },
    { fieldName: 'unpaidLeaveDays', nameDisplay: 'Ngày nghỉ không lương', type: 1, isRequired: false },

    { fieldName: 'vtcvSalaryWorkingDays', nameDisplay: 'Tổng công tính lương VTCV', type: 1, isRequired: false },
    { fieldName: 'performanceSalaryWorkingDays', nameDisplay: 'Tổng công tính lương hiệu quả', type: 1, isRequired: false },
    { fieldName: 'actualSalaryWorkingDaysHidden', nameDisplay: 'Tổng công tính lương thực tế ẩn', type: 1, isRequired: false },

    { fieldName: 'nightShiftWorkingDays', nameDisplay: 'Ngày công ca đêm', type: 1, isRequired: false },
    { fieldName: 'holidayDutyWorkingDays', nameDisplay: 'Ngày công trực ca lễ tết', type: 1, isRequired: false },
    { fieldName: 'standardWorkingDaysOfMonth', nameDisplay: 'Ngày công chuẩn của tháng', type: 1, isRequired: false },

    { fieldName: 'bhxhBaseSalary', nameDisplay: 'Mức lương làm căn cứ đóng BHXH', type: 1, isRequired: false },
    { fieldName: 'vtcvSalary', nameDisplay: 'Tiền lương VTCV', type: 1, isRequired: false },
    { fieldName: 'workCompletionRate', nameDisplay: 'Tỉ lệ hoàn thành công việc', type: 1, isRequired: false },
    { fieldName: 'performanceSalary', nameDisplay: 'Lương hiệu quả', type: 1, isRequired: false },
    { fieldName: 'nightAndHolidaySalary', nameDisplay: 'Lương ca đêm và trực ca lễ tết', type: 1, isRequired: false },

    { fieldName: 'totalVtcvAndPerformanceSalary', nameDisplay: 'Tổng lương VTCV và hiệu quả', type: 1, isRequired: false },
    { fieldName: 'agreedSalaryColumn', nameDisplay: 'Cột lương thỏa thuận trả cho người lao động', type: 1, isRequired: false },
    { fieldName: 'salaryArrears', nameDisplay: 'Truy lĩnh tiền lương', type: 1, isRequired: false },
];

const isEmptyValue = (value: any): boolean => {
    return value === null ||
        value === undefined ||
        value === "" ||
        (typeof value === 'string' && value.trim() === "");
};

const tryCast = (
    value: any,
    type: DataTypes
): TryCastResult<any> => {

    if (value === undefined || value === null || value === '') {
        return { success: true, value: null };
    }

    try {
        switch (type) {
            case DataTypes.Number: {
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

            case DataTypes.Boolean: {
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

            case DataTypes.Date: {
                let date: Date | null = null;

                // Xử lý Excel serial date (nếu value là số nguyên dương)
                if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
                    // Excel serial date bắt đầu từ 1900-01-01 (giá trị 1)
                    // Công thức: Date(1899, 11, 30) + value (vì Excel có bug ở 1900 không nhuận, nhưng ta dùng offset chuẩn)
                    const excelBaseDate = new Date(1899, 11, 30); // Base cho serial
                    date = new Date(excelBaseDate.getTime() + value * 86400000); // 86400000 ms = 1 ngày
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

                // Fallback: Sử dụng new Date(value) cho các định dạng khác (ISO, MM/dd/yyyy, etc.)
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

            case DataTypes.String:
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

export default function ExcelCreate() {
    const [config, setConfig] = useState<ExcelConfig>({
        id: 0,
        templateFileName: '',
        configName: '',
        departmentId: 0,
        configType: ConfigType.Salary,
        details: [],
        acctions: ''
    });

    const [details, setDetails] = useState<ExcelConfigDetail[]>([]);
    const [isSaving, setIsSaving] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [editingDetail, setEditingDetail] = useState<ExcelConfigDetail | null>(null);
    const [selectedFile, setSelectedFile] = useState<File | null>(null);
    const [fileName, setFileName] = useState<string>('');


    const [file, setFile] = useState<File | null>(null);
    const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
    const [selectedSheet, setSelectedSheet] = useState('');

    const [hasHeader, setHasHeader] = useState<boolean | null>(null);
    const [step, setStep] = useState<Step>('select_mode');
    const [headerMappings, setHeaderMappings] = useState<HeaderMapping[]>([]);
    const [selectedHeaderCells, setSelectedHeaderCells] = useState<Set<string>>(new Set());
    const [sheetsConfigured, setsheetsConfigured] = useState<Set<string>>(new Set());
    const [fields, setFileds] = useState<Field[]>([]);
    const [extractedData, setExtractedData] = useState<Record<string, any>[]>([]);
    const [previewData, setPreviewData] = useState<Record<string, any>[]>([]);
    const { isOpen, onOpen, onClose, onOpenChange } = useDisclosure();
    const [cellError, setCellError] = useState<CellError[]>([]);
    const [errors, setErrors] = useState<Record<string, string[]>>({});
    const [numberSelected, setNumberSelected] = useState<number>();


    const dataTypeLabels = {
        [DataTypes.String]: 'Chuỗi',
        [DataTypes.Number]: 'Số',
        [DataTypes.Date]: 'Ngày',
        [DataTypes.Boolean]: 'Boolean',
        [DataTypes.Decimal]: 'Số thập phân'
    };

    const configTypeLabels = {
        [ConfigType.Salary]: 'Lương',
        [ConfigType.Insurance]: 'Bảo hiểm'
    };

    const configTypeDepartments = {
        [DepartmentId.DepartmentA]: 'DepartmentA',
        [DepartmentId.DepartmentB]: 'DepartmentB'
    };

    // Save config
    const handleSaveConfig = async (): Promise<ExcelConfig> => {
        setIsSaving(true);
        setError(null);

        try {
            const response = await fetch(`${API_BASE_URL}/excel-config`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(config),
            });

            if (!response.ok) {
                throw new Error('Không thể lưu cấu hình');
            }

            const data: ExcelConfig = await response.json();
            alert('Lưu cấu hình thành công!');
            return data;

        } catch (err) {
            const message =
                err instanceof Error ? err.message : 'Đã xảy ra lỗi khi lưu';

            setError(message);
            console.error('Error saving config:', err);

            throw err;
        } finally {
            setIsSaving(false);
        }
    };


    // Open modal for new detail
    const handleAddDetail = () => {
        setEditingDetail({
            id: 0,
            configId: config.id,
            fieldName: '',
            displayName: '',
            columnPosition: 0,
            rowPosition: 0,
            sheetName: '',
            dataType: DataTypes.String,
            isRequired: false
        });
        onOpen();
    };

    // Open modal for editing
    const handleEditDetail = (detail: ExcelConfigDetail) => {
        setEditingDetail({ ...detail });
        onOpen();
    };

    // Save detail
    const handleSaveDetail = async (configId: number) => {
        if (!configId) return;

        const data = details.map(prev => ({ ...prev, configId: configId }))

        try {
            const response = await fetch(`${API_BASE_URL}/excel-config/${config.id}/details`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(data),
            });

            if (!response.ok) throw new Error('Không thể lưu chi tiết');

            onClose();
            setEditingDetail(null);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi khi lưu');
            console.error('Error saving detail:', err);
        }
    };

    // Delete detail
    const handleDeleteDetail = async (detailId: number) => {
        if (!confirm('Bạn có chắc chắn muốn xóa chi tiết này?')) return;

        try {
            const response = await fetch(`${API_BASE_URL}/excel-config/${config.id}/details/${detailId}`, {
                method: 'DELETE',
            });

            if (!response.ok) throw new Error('Không thể xóa chi tiết');
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi khi xóa');
            console.error('Error deleting detail:', err);
        }
    };

    const onSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
        e.preventDefault();

        //if (!workbook || !errors || cellError!) return;

        if (!details) {
            alert('Chưa setting chi tiết!')
            return;
        }

        const configRes = await handleSaveConfig();

        await handleSaveDetail(configRes.id);

        await handleUpload();

        alert("Lưu thành công!");
    };

    const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        const fileName = file?.name ?? "";
        if (file && (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.type === 'application/vnd.ms-excel')) {
            setSelectedFile(file);
            const guid = crypto.randomUUID();
            setFileName(fileName)
        };

        const uploadedFile = event.target.files?.[0];
        if (!uploadedFile) return;

        setFile(uploadedFile);
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                const wbRaw = XLSX.read(event.target?.result, { type: 'binary' });
                const wb = filterVisibleWorkbook(wbRaw);
                setWorkbook(wb);
                setSelectedSheet(wb.SheetNames[0]);
            } catch (error) {
                alert('Lỗi khi đọc file Excel: ' + (error as Error).message);
            }
        };

        reader.readAsBinaryString(uploadedFile);
        setConfig(prev => ({
            ...prev,
            templateFileName: fileName
        }));

        setFileds(Tables);
    };

    // Hàm gửi file lên BE
    const handleUpload = async () => {
        if (!selectedFile) {
            return;
        }

        const formData = new FormData();
        formData.append(fileName, selectedFile);

        try {
            const response = await fetch(`${API_BASE_URL}/excel-config/upload`, {
                headers: {
                    'Content-Type': 'multipart/form-data',
                },
                body: formData,
                method: 'POST'
            });
        } catch (error) {

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

    const resetConfiguration = () => {
        setHasHeader(null);
        setStep('select_mode');
        setHeaderMappings([]);
        setSelectedHeaderCells(new Set());
        setDetails([]);
        setsheetsConfigured(new Set());
        setCellError([]);
        setFileds(prev => prev.map(f => ({
            ...f,
            isSelected: false
        })));
        setErrors({});
        setCellError([]);
        setNumberSelected(undefined);
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
                setDetails(prev => prev.filter(m => !(m.columnPosition === colIdx && m.rowPosition === rowIdx + 1 && m.sheetName === sheet)));
            } else {
                if (details.length >= fields.length) {
                    return
                }
                newSelected.add(cellKey);
                const worksheet = workbook?.Sheets[sheet];
                const data = XLSX.utils.sheet_to_json(worksheet!, { header: 1, defval: '' }) as any[][];
                const cellValue = data[rowIdx]?.[colIdx] || '';
                const field = fields[details.length];

                setHeaderMappings(prev => [...prev, {
                    col: colIdx,
                    originalValue: cellValue.toString(),
                    displayName: cellValue.toString(),
                    rowIndex: rowIdx,
                    sheet: sheet
                }]);
                setDetails(prev => [...prev,
                {
                    columnPosition: colIdx,
                    rowPosition: rowIdx + 1,
                    sheetName: sheet,
                    fieldName: field.fieldName,
                    displayName: field.nameDisplay,
                    isRequired: field.isRequired ?? false,
                    dataType: field.type,
                }]);
            }

            setSelectedHeaderCells(newSelected);
        } else if (step === 'select_data_start') {
            const existingIndex = details.findIndex(cell => cell.rowPosition === rowIdx && cell.columnPosition === colIdx && cell.sheetName === sheet);

            if (existingIndex !== -1) {
                setDetails(prev => prev.filter((_, idx) => idx !== existingIndex));
            } else {
                if (details.length >= fields.length) {
                    return
                }
                const field = fields[details.length];
                setDetails(prev => [...prev,
                {
                    columnPosition: colIdx,
                    rowPosition: rowIdx + 1,
                    sheetName: sheet,
                    fieldName: field.fieldName,
                    displayName: field.nameDisplay,
                    isRequired: field.isRequired ?? false,
                    dataType: field.type,
                }]);
            }
        }
    };

    const getCellStyle = (rowIdx: number, colIdx: number, sheet: string) => {
        if (step === 'select_headers' && selectedHeaderCells.has(`${rowIdx}-${colIdx}-${sheet}`)) {
            return 'bg-green-200 font-bold border-2 border-green-500 cursor-pointer';
        }
        if (cellError.some(cell => cell.row === rowIdx && cell.col === colIdx && cell.sheet === sheet)) {
            return 'bg-red-500 text-white font-semibold border-2 border-blue-500';
        }
        if (step === 'select_data_start' && details.some(cell => cell.rowPosition === rowIdx && cell.columnPosition === colIdx && cell.sheetName === sheet)) {
            return 'bg-blue-200 font-bold border-2 border-gray-400 cursor-pointer';
        }
        if ((step === 'configure' || step === 'set_row_start') && selectedHeaderCells.has(`${rowIdx}-${colIdx}-${sheet}`)) {
            return 'bg-green-200 font-bold border-2 border-green-500';
        }
        if ((step === 'configure' || step === 'set_row_start') && details.some(cell => cell.rowPosition === rowIdx && cell.columnPosition === colIdx && cell.sheetName === sheet)) {
            return 'bg-blue-200 border-2 border-gray-400';
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
        setStep('set_row_start');
        setFileds(prev => prev.map(f => ({
            ...f,
            isSelected: true
        })));
    };

    const confirmDataStartSelection = () => {
        if (details.length === 0) {
            alert('Vui lòng chọn ít nhất một ô để bắt đầu lấy dữ liệu!');
            return;
        }
        setStep('configure');
        setFileds(prev => prev.map(f => ({
            ...f,
            isSelected: true
        })));
    };

    const updateHeaderName = (index: number, newName: string) => {
        setHeaderMappings(prev => prev.map((item, idx) =>
            idx == index ? { ...item, displayName: newName } : item
        ));
    };

    const updateDataStartRow = (index: number, newRow: number) => {
        setDetails(prev => prev.map((item, idx) =>
            idx == index ? { ...item, row: newRow - 1 } : item
        ));
    };

    const updateDataField = (index: number, newField?: string, oldField?: string) => {
        setDetails(prev => prev.map((item, idx) =>
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

    const excelColName = (col: number): string => {
        let name = '';
        while (col >= 0) {
            name = String.fromCharCode((col % 26) + 65) + name;
            col = Math.floor(col / 26) - 1;
        }
        return name;
    };

    const checkRequiredFields = () => {
        const mappedFields = new Set(
            details.map(d => d.fieldName)
        );

        return fields
            .filter(f => f.isRequired)
            .every(f => mappedFields.has(f.fieldName));
    };

    const preViewData = () => {
        setCellError([]);
        setErrors({});
        if (!workbook) return;

        // if (!checkRequiredFields()) {
        //     alert('Có trường bắt buộc chưa được mapping data!')
        //     return;
        // }

        if (!details) {
            alert('Chưa setting chi tiết!')
            return;
        }

        const result = extractDataWithConfig(workbook, details, fields);
        setExtractedData(result.data);
        setPreviewData(result.data.map((d, i) => ({
            key: `key_${i}`,
            stt: i + 1,
            ...d
        })));

        if (!result.isSuccess) {
            const uniqueIndexes = new Set(result.cellsErr.map(c => c.index));

            const newErrors = Array.from(uniqueIndexes).reduce<Record<string, string[]>>(
                (acc, index) => addValidationError(acc, `field${index}`, 'Lỗi mapping kiểu dữ liệu'),
                {}
            );
            setCellError(result.cellsErr);
            setErrors(newErrors);
        }

        onOpen();
    };

    const extractDataWithConfig = (
        workbook: XLSX.WorkBook,
        dataStartCells: ExcelConfigDetail[],
        fields: Field[]
    ): SubmitResult<Record<string, any>[]> => {
        const result: Record<string, any>[] = [];
        let isSuccess: boolean = true;
        let cellsErr: CellError[] = [];

        const columnData: any[][] = dataStartCells.map(cfg => {
            const worksheet = workbook.Sheets[cfg.sheetName];
            if (!worksheet) return [];

            const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null }) as any[][];
            const data: any[] = [];
            for (let i = cfg.rowPosition; i < sheetData.length; i++) {
                const row = sheetData[i] || [];
                data.push(row[cfg.columnPosition] ?? null);
            }
            return data;
        });

        const maxRows = Math.max(...columnData.map(col => col.length), 0);

        for (let rowIdx = 0; rowIdx < maxRows; rowIdx++) {
            const dataRow: Record<string, any> = {};
            let allEmpty = true;

            for (let colIdx = 0; colIdx < dataStartCells.length; colIdx++) {
                const value = columnData[colIdx][rowIdx] ?? null;
                const fieldName = dataStartCells[colIdx].fieldName || `Column_${colIdx}`;
                const field = fields.find(f => f.fieldName === fieldName);

                let res = tryCast(value, field?.type ?? DataTypes.String)
                if (!res.success) {
                    const startCol = dataStartCells[colIdx].columnPosition;
                    const startRow = dataStartCells[colIdx].rowPosition;
                    const sheet = dataStartCells[colIdx].sheetName;
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

        return { isSuccess: isSuccess, data: result, cellsErr: cellsErr };
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

    useEffect(() => {
        const next = new Set(details.map(x => x.sheetName));

        setsheetsConfigured(prev => {
            if (prev.size === next.size &&
                [...prev].every(x => next.has(x))) {
                return prev;
            }
            return next;
        });
    }, [details]);

    return (
        <Form
            className="min-h-screen p-6"
            onSubmit={onSubmit}
            validationErrors={errors}
        >
            <div className="mb-6 w-full">
                <div className="flex justify-between items-center w-full">
                    <div>
                        <h1 className="text-3xl font-bold">Cấu hình Extract Excel</h1>
                        <p className="mt-2">Quản lý cấu hình import/export Excel</p>
                    </div>
                    <div>
                        <Button
                            color="success"
                            type="submit"
                            isLoading={isSaving}
                            startContent={<Save className="w-4 h-4" />}
                        >
                            Lưu
                        </Button>
                    </div>
                </div>
            </div>

            {error && (
                <div className="mb-4 p-4 bg-red-100 border border-red-400 text-red-700 rounded">
                    {error}
                </div>
            )}

            {/* Config Form */}
            <div className="w-full mb-6 border-2 border-blue-200 rounded-lg shadow-md p-6">
                <h2 className="text-xl font-semibold mb-4">Thông tin cấu hình</h2>
                {!selectedFile ? (
                    <div className="border-2 border-dashed border-gray-300 rounded-lg p-2 mb-4 text-center hover:border-blue-400 transition-colors">
                        <Upload className="mx-auto  mb-4" size={24} />
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
                                <p className="font-semibold">{selectedFile.name}</p>
                                <p className="text-sm">
                                    {(selectedFile.size / 1024).toFixed(2)} KB
                                </p>
                            </div>
                        </div>
                        <Button
                            onPress={() => setSelectedFile(null)}
                            color='danger'
                            startContent={<X size={16} />}
                        >
                            Xóa file
                        </Button>
                    </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    <Input
                        type="text"
                        label="Tên cấu hình"
                        value={config.configName ?? ''}
                        onChange={(e) => setConfig(prev => ({ ...prev, configName: e.target.value }))}
                        isRequired
                    />
                    <Select
                        label="Department ID"
                        onChange={(e) => setConfig(prev => ({ ...prev, departmentId: parseInt(e.target.value) || 0 }))}
                        isRequired
                    >
                        {Object.entries(configTypeDepartments).map(([key, value]) => (
                            <SelectItem key={key} textValue={value}>
                                {value}
                            </SelectItem>
                        ))}
                    </Select>
                    <Select
                        label="Loại cấu hình"
                        onChange={(e) => setConfig(prev => ({ ...prev, configType: parseInt(e.target.value) as ConfigType }))}
                        isRequired
                    >
                        {Object.entries(configTypeLabels).map(([key, value]) => (
                            <SelectItem key={key} textValue={value}>
                                {value}
                            </SelectItem>
                        ))}
                    </Select>
                </div>
            </div>


            <div>
                {workbook && (
                    <div className={fields && 'grid grid-cols-3 gap-1'}>
                        <div className=" border-2 border-blue-200 rounded-lg p-2 max-h-[640px] shadow-md">
                            <div className="flex items-center gap-2 mb-4">
                                <Settings size={20} className="text-blue-600" />
                                <h3 className="font-bold text-lg">Cấu hình chi tiết</h3>
                            </div>

                            {fields && step === 'select_mode' && fields && (
                                <div className="space-y-4">
                                    <p className="text-sm font-semibold mb-3">
                                        Dữ liệu của bạn có header không?
                                    </p>
                                    <div className="grid grid-cols-2 gap-2">
                                        <Button
                                            onPress={() => handleSelectMode(true)}
                                            color='success'
                                        >
                                            ✓ Có Header
                                        </Button>
                                        <Button
                                            onPress={() => handleSelectMode(false)}
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
                                        {headerMappings.length / fields.length < 1 && <p className="text-sm mt-2">
                                            <strong>Chọn header cho trường:</strong> {fields[headerMappings.length].nameDisplay}
                                        </p>}
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


                            {step === 'set_row_start' && (
                                <div className="space-y-4">
                                    <div className="border-2 border-gray-200 p-3 rounded-lg space-y-3">
                                        <p className="text-sm font-semibold">
                                            📌 Chọn dòng bắt đầu lấy dữ liệu
                                        </p>
                                        <p className="text-sm">
                                            Nếu không chọn, mặc định ví trí bắt đầu lấy dữ liệu là dòng header + 1
                                        </p>

                                        <NumberInput
                                            type="number"
                                            onValueChange={setNumberSelected}
                                            label='Dòng bắt đầu:'
                                            minValue={1}
                                        />
                                    </div>
                                    <div className="grid grid-cols-2 gap-2">
                                        <Button
                                            onClick={() => {
                                                setStep('configure');
                                                if (!numberSelected) return;

                                                setDetails(prev =>
                                                    prev.map(cell => ({
                                                        ...cell,
                                                        rowPosition: numberSelected - 1
                                                    }))
                                                );
                                            }}
                                            disabled={headerMappings.length === 0}
                                            className="flex items-center justify-center gap-1 px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
                                        >
                                            <Check size={16} />
                                            Xác nhận
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
                                            <strong>Đã chọn:</strong> {details.length} / {fields.length} ô
                                        </p>
                                        {details.length / fields.length < 1 && <p className="text-sm mt-2">
                                            <strong>Chọn vị trí bắt đầu cho trường:</strong> {fields[headerMappings.length].nameDisplay}
                                        </p>}
                                    </div>
                                    <div className="grid grid-cols-2 gap-2">
                                        <button
                                            onClick={confirmDataStartSelection}
                                            disabled={details.length === 0}
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
                                <div className='w-full grid grid-cols-1'>
                                    {hasHeader ? (
                                        <div>
                                            <h4 className="text-sm font-semibold mb-2">
                                                Headers ({headerMappings.length}):
                                            </h4>
                                            <div className="w-full grid grid-cols-1 gap-2 max-h-[465px] overflow-y-auto">
                                                {headerMappings.map((mapping, idx) => (
                                                    <Card key={idx} className="p-1 space-y-2">
                                                        <div className='grid grid-cols-2 justify-items-stretch mb-1'>
                                                            <p className="text-xs">
                                                                Cột {excelColName(mapping.col)}
                                                            </p>
                                                            <p className='justify-self-end text-xs'>{mapping.sheet}</p>
                                                        </div>

                                                        <Input
                                                            type="text"
                                                            value={mapping.displayName}
                                                            onChange={(e) => updateHeaderName(idx, e.target.value)}
                                                            label='Header:'
                                                            disabled
                                                        />
                                                        <Select
                                                            label="Trường"
                                                            placeholder="Chọn trường"
                                                            disabledKeys={fields.filter(f => f.isSelected && f.fieldName != details[idx].fieldName).map(f => f.fieldName)}
                                                            onChange={(e) => {
                                                                updateDataField(idx, e.target.value, details[idx].fieldName);
                                                            }}
                                                            isRequired
                                                            defaultSelectedKeys={[details[idx].fieldName ?? '']}
                                                            name={'field' + idx}
                                                        >
                                                            {fields?.map((f) => (
                                                                <SelectItem
                                                                    key={f.fieldName}
                                                                    textValue={f.nameDisplay}
                                                                >
                                                                    {f.nameDisplay} ({f.type}) {f.isRequired && <span className='text-red-600'>*</span>}
                                                                </SelectItem>
                                                            ))}
                                                        </Select>
                                                        <NumberInput
                                                            type="number"
                                                            value={details[idx].rowPosition + 1}
                                                            onChange={(e) => updateDataStartRow(idx, Number(e))}
                                                            label='Dòng bắt đầu:'
                                                            isRequired
                                                            minValue={1}
                                                        />

                                                    </Card>
                                                ))}
                                            </div>
                                        </div>
                                    ) : <div>
                                        <h4 className="text-sm font-semibold mb-2">
                                            Columns ({details.length}):
                                        </h4>
                                        <div className="grid grid-cols-1 gap-2 max-h-96 overflow-y-auto p-2">
                                            {details.map((mapping, idx) => (
                                                <Card key={idx} className="p-2 space-y-2">
                                                    <div className='grid grid-cols-2 justify-items-stretch mb-1'>
                                                        <p className="text-xs">
                                                            Cột {excelColName(mapping.columnPosition)}
                                                        </p>
                                                        <p className='justify-self-end text-xs'>{mapping.sheetName}</p>
                                                    </div>
                                                    <Select
                                                        label="Trường:"
                                                        placeholder="Chọn trường"
                                                        defaultSelectedKeys={[mapping.fieldName ?? '']}
                                                        disabledKeys={fields.filter(f => f.isSelected && f.fieldName != details[idx].fieldName).map(f => f.fieldName)}
                                                        onChange={(e) => {
                                                            updateDataField(idx, e.target.value, mapping.fieldName);
                                                        }}
                                                    >
                                                        {fields.map((f) => (
                                                            <SelectItem key={f.fieldName} textValue={f.nameDisplay}>{f.nameDisplay}</SelectItem>
                                                        ))}
                                                    </Select>
                                                    <NumberInput
                                                        type="number"
                                                        value={mapping.rowPosition + 1}
                                                        onChange={(e) => updateDataStartRow(idx, Number(e))}
                                                        label='Dòng bắt đầu:'
                                                    />
                                                </Card>
                                            ))}
                                        </div>
                                    </div>}

                                    <div className="grid grid-cols-2 gap-2 pt-3">
                                        <Button
                                            onPress={preViewData}
                                            className="flex items-center justify-center gap-1 px-4 py-2 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-colors"
                                        >
                                            <View size={16} />
                                            Xem trước
                                        </Button>
                                        <Button
                                            onClick={() => { setStep('select_mode'); resetConfiguration(); }}
                                            className="px-4 py-2 bg-gray-500 text-white rounded-lg hover:bg-gray-600 transition-colors"
                                        >
                                            ← Cấu hình lại
                                        </Button>
                                    </div>
                                </div>
                            )}
                        </div>

                        {fields && <div className="col-span-2">
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
        </Form >
    );
}