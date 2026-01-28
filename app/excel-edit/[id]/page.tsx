'use client'

import { ExcelViewer } from "@/components/excel-viewer";
import { API_BASE_URL, CellError, ConfigType, DataTypes, DepartmentId, ExcelConfig, ExcelConfigDetail, Field, HeaderMapping, Step, SubmitResult, Tables, TryCastResult } from "@/types/excel-type";
import { extractDataWithConfig, isEmptyValue, mappingDataType, tryCast } from "@/utils/excel";
import { Button, Input, Select, SelectItem, Table, TableHeader, TableColumn, TableBody, TableRow, TableCell, Checkbox, Modal, ModalContent, ModalHeader, ModalBody, ModalFooter, useDisclosure, NumberInput, Form, Card } from "@heroui/react";
import { RefreshCw, Save, Plus, Trash2, Edit, Upload, CheckCircle, X, Settings, Check, Download, View } from "lucide-react";
import { ChangeEvent, useEffect, useState } from "react";
import * as XLSX from 'xlsx';
import { useParams } from 'next/navigation';

export default function ExcelEdit() {

    const { id } = useParams<{ id: string }>();
    const [config, setConfig] = useState<ExcelConfig>({
        id: Number(id),
        templateFileName: '',
        configName: '',
        departmentId: 0,
        configType: ConfigType.Salary,
        acctions: ''
    });
    const [details, setDetails] = useState<ExcelConfigDetail[]>([]);
    const [isSaving, setIsSaving] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [selectedFile, setSelectedFile] = useState<File | null>(null);
    const [fileName, setFileName] = useState<string>('');
    const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
    const [selectedSheet, setSelectedSheet] = useState('');

    const [hasHeader, setHasHeader] = useState<boolean | null>(null);
    const [step, setStep] = useState<Step>('configure');
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

    const configTypeLabels = {
        [ConfigType.Salary]: 'Lương',
        [ConfigType.Insurance]: 'Bảo hiểm'
    };

    const configTypeDepartments = {
        [DepartmentId.All]: 'Tất cả',
        [DepartmentId.DepartmentA]: 'DepartmentA',
        [DepartmentId.DepartmentB]: 'DepartmentB'
    };

    // Fetch config
    const fetchConfig = async (): Promise<ExcelConfig> => {
        setError(null);
        try {
            const response = await fetch(`${API_BASE_URL}/excel-config/${id}`);
            if (!response.ok) throw new Error('Không thể tải cấu hình');
            const data = await response.json();
            setConfig(data);
            return data;
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi');
            console.error('Error fetching config:', err);
            throw err;
        } finally {
        }
    };

    // Fetch details
    const fetchDetails = async () => {
        setError(null);
        try {
            const response = await fetch(`${API_BASE_URL}/excel-config/${id}/details`);
            if (!response.ok) throw new Error('Không thể tải chi tiết cấu hình');
            const data = await response.json();
            setDetails(data);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi');
            console.error('Error fetching details:', err);
        } finally {
        }
    };

    const downloadAsFile = async (fileName: string): Promise<File> => {
        console.log(fileName)
        const response = await fetch(
            `${API_BASE_URL}/excel-config/download?fileName=${encodeURIComponent(fileName)}`
        );

        if (!response.ok) {
            throw new Error('Không thể tải file');
        }

        const blob = await response.blob();

        return new File([blob], fileName, {
            type: blob.type
        });
    };

    // Initial load
    useEffect(() => {
        const loadData = async () => {
            const configRes = await fetchConfig();
            await fetchDetails();
            const file = await downloadAsFile(configRes.templateFileName);
            if (!!file) {
                setSelectedFile(file);
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

                reader.readAsBinaryString(file);
                setFileds(Tables);
            }
        };
        loadData();
    }, []);

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

        if (details.length === 0) {
            alert('Chưa setting chi tiết!')
            return;
        }

        const configRes = await handleSaveConfig();

        if (!configRes)
            return

        await handleSaveDetail(configRes.id);

        await handleUpload();
    };

    const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        const fileName = file?.name ?? "";
        if (file && (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.type === 'application/vnd.ms-excel')) {
            setSelectedFile(file);
            const guid = crypto.randomUUID();
            setFileName(fileName)
        };

        if (!file) return;

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

        reader.readAsBinaryString(file);
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
        setConfig({
            id: Number(id),
            templateFileName: '',
            configName: '',
            departmentId: 0,
            configType: ConfigType.Salary,
            acctions: ''
        });
        setWorkbook(null);
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
            idx == index ? { ...item, fieldName: newField ?? '' } : item
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
                            onPress={() => {
                                setSelectedFile(null);
                                resetConfiguration();
                            }}
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
                        selectedKeys={[config.departmentId.toString()]}
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
                        selectedKeys={[config.configType.toString()]}
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
                                            onPress={confirmHeaderSelection}
                                            disabled={headerMappings.length === 0}
                                            className="flex items-center justify-center gap-1 px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
                                        >
                                            <Check size={16} />
                                            Xác nhận
                                        </Button>
                                        <Button
                                            onPress={() => { setStep('select_mode'); resetConfiguration(); }}
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
                                            onPress={() => {
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
                                                                    {f.nameDisplay} ({mappingDataType(f.type)}) {f.isRequired && <span className='text-red-600'>*</span>}
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
                                        <div className="grid grid-cols-1 gap-2 max-h-[465px] overflow-y-auto p-2">
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
                                sheetsConfigured={sheetsConfigured}
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