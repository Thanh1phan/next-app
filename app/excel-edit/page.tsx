'use client'

import { Button, Input, Select, SelectItem, Table, TableHeader, TableColumn, TableBody, TableRow, TableCell, Checkbox, Modal, ModalContent, ModalHeader, ModalBody, ModalFooter, useDisclosure } from "@heroui/react";
import { RefreshCw, Save, Plus, Trash2, Edit } from "lucide-react";
import { ChangeEvent, useEffect, useState } from "react";

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

// Types
interface ExcelConfigDetail {
    id: number;
    configId: number;
    fieldName: string;
    displayName: string;
    columnPosition: number;
    rowPosition: number;
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
    const [isLoadingConfigs, setIsLoadingConfigs] = useState(true);
    const [isLoadingDetails, setIsLoadingDetails] = useState(false);
    const [isSaving, setIsSaving] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [editingDetail, setEditingDetail] = useState<ExcelConfigDetail | null>(null);
    const [selectedFile, setSelectedFile] = useState<File | null>(null);
    const [fileName, setFileName] = useState<string>('');
    const { isOpen, onOpen, onClose } = useDisclosure();

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

    // Fetch config
    const fetchConfig = async () => {
        setIsLoadingConfigs(true);
        setError(null);
        try {
            const response = await fetch(`${API_BASE_URL}/excel-config/1`);
            if (!response.ok) throw new Error('Không thể tải cấu hình');
            const data = await response.json();
            setConfig(data);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi');
            console.error('Error fetching config:', err);
        } finally {
            setIsLoadingConfigs(false);
        }
    };

    // Fetch details
    const fetchDetails = async () => {
        setIsLoadingDetails(true);
        setError(null);
        try {
            const response = await fetch(`${API_BASE_URL}/excel-config/1/details`);
            if (!response.ok) throw new Error('Không thể tải chi tiết cấu hình');
            const data = await response.json();
            console.log('Details data received:', data);
            setDetails(Array.isArray(data) ? data : []);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi');
            console.error('Error fetching details:', err);
        } finally {
            setIsLoadingDetails(false);
        }
    };

    // Initial load
    useEffect(() => {
        const loadData = async () => {
            await fetchConfig();
            await fetchDetails();
        };
        loadData();
    }, []);

    // Save config
    const handleSaveConfig = async () => {
        setIsSaving(true);
        setError(null);
        try {
            const response = await fetch(`${API_BASE_URL}/excel-config/${config.id}`, {
                method: 'PUT',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(config),
            });
            if (!response.ok) throw new Error('Không thể lưu cấu hình');
            alert('Lưu cấu hình thành công!');
            await fetchConfig();
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi khi lưu');
            console.error('Error saving config:', err);
        } finally {
            setIsSaving(false);
        }
    };

    // Refresh all data
    const handleRefresh = async () => {
        await Promise.all([fetchConfig(), fetchDetails()]);
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
    const handleSaveDetail = async () => {
        if (!editingDetail) return;

        try {
            const method = editingDetail.id === 0 ? 'POST' : 'PUT';
            const url = editingDetail.id === 0
                ? `${API_BASE_URL}/excel-config/${config.id}/details`
                : `${API_BASE_URL}/excel-config/${config.id}/details/${editingDetail.id}`;

            const response = await fetch(url, {
                method,
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(editingDetail),
            });

            if (!response.ok) throw new Error('Không thể lưu chi tiết');

            await fetchDetails();
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

            await fetchDetails();
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi khi xóa');
            console.error('Error deleting detail:', err);
        }
    };

    const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (file && (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.type === 'application/vnd.ms-excel')) {
            setSelectedFile(file);
            const guid = crypto.randomUUID();
            setFileName(file.name + guid)
        } else {
        }
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

    return (
        <div className="min-h-screen p-6 bg-gray-50">
            <div className="mb-6">
                <div className="flex justify-between items-center">
                    <div>
                        <h1 className="text-3xl font-bold text-gray-800">Cấu hình Extract Excel</h1>
                        <p className="mt-2 text-gray-600">Quản lý cấu hình import/export Excel</p>
                    </div>
                    <div className="flex gap-2">
                        <Button
                            color="success"
                            onPress={handleSaveConfig}
                            isLoading={isSaving}
                            startContent={<Save className="w-4 h-4" />}
                        >
                            Lưu
                        </Button>
                        <Button
                            onPress={handleRefresh}
                            isLoading={isLoadingConfigs || isLoadingDetails}
                            startContent={<RefreshCw className={`w-4 h-4 ${(isLoadingConfigs || isLoadingDetails) ? 'animate-spin' : ''}`} />}
                        >
                            Làm mới
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
            <div className="mb-6 bg-white border-2 border-blue-200 rounded-lg shadow-md p-6">
                <h2 className="text-xl font-semibold mb-4 text-gray-700">Thông tin cấu hình</h2>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <Input
                        type="text"
                        label="Tên cấu hình"
                        value={config.configName ?? ''}
                        onChange={(e) => setConfig(prev => ({ ...prev, configName: e.target.value }))}
                        isDisabled={isLoadingConfigs}
                    />
                    <Input
                        type="text"
                        label="Tên file template"
                        value={config.templateFileName ?? ''}
                        onChange={(e) => setConfig(prev => ({ ...prev, templateFileName: e.target.value }))}
                        isDisabled={isLoadingConfigs}
                    />
                    <Input
                        type="number"
                        label="Department ID"
                        value={config.departmentId?.toString() ?? '0'}
                        onChange={(e) => setConfig(prev => ({ ...prev, departmentId: parseInt(e.target.value) || 0 }))}
                        isDisabled={isLoadingConfigs}
                    />
                    <Select
                        label="Loại cấu hình"
                        selectedKeys={[config.configType?.toString()]}
                        onChange={(e) => setConfig(prev => ({ ...prev, configType: parseInt(e.target.value) as ConfigType }))}
                        isDisabled={isLoadingConfigs}
                    >
                        {Object.entries(configTypeLabels).map(([key, value]) => (
                            <SelectItem key={key} textValue={key}>
                                {value}
                            </SelectItem>
                        ))}
                    </Select>
                </div>
            </div>

            {/* Details Table */}
            <div className="bg-white rounded-lg shadow-md p-6">
                <div className="flex justify-between items-center mb-4">
                    <h2 className="text-xl font-semibold text-gray-700">Chi tiết cấu hình</h2>
                    <Button
                        color="primary"
                        onPress={handleAddDetail}
                        startContent={<Plus className="w-4 h-4" />}
                    >
                        Thêm chi tiết
                    </Button>
                </div>

                <Table aria-label="Config details table">
                    <TableHeader>
                        <TableColumn>TÊN TRƯỜNG</TableColumn>
                        <TableColumn>TÊN HIỂN THỊ</TableColumn>
                        <TableColumn>CỘT</TableColumn>
                        <TableColumn>HÀNG</TableColumn>
                        <TableColumn>KIỂU DỮ LIỆU</TableColumn>
                        <TableColumn>BẮT BUỘC</TableColumn>
                        <TableColumn>THAO TÁC</TableColumn>
                    </TableHeader>
                    <TableBody
                        items={details}
                        isLoading={isLoadingDetails}
                        emptyContent="Chưa có chi tiết cấu hình"
                    >
                        {(detail) => (
                            <TableRow key={detail.id}>
                                <TableCell>{detail.fieldName}</TableCell>
                                <TableCell>{detail.displayName}</TableCell>
                                <TableCell>{detail.columnPosition}</TableCell>
                                <TableCell>{detail.rowPosition}</TableCell>
                                <TableCell>{dataTypeLabels[detail.dataType]}</TableCell>
                                <TableCell>
                                    <Checkbox isSelected={detail.isRequired} isDisabled />
                                </TableCell>
                                <TableCell>
                                    <div className="flex gap-2">
                                        <Button
                                            size="sm"
                                            color="primary"
                                            variant="light"
                                            onPress={() => handleEditDetail(detail)}
                                            startContent={<Edit className="w-4 h-4" />}
                                        >
                                            Sửa
                                        </Button>
                                        <Button
                                            size="sm"
                                            color="danger"
                                            variant="light"
                                            onPress={() => handleDeleteDetail(detail.id)}
                                            startContent={<Trash2 className="w-4 h-4" />}
                                        >
                                            Xóa
                                        </Button>
                                    </div>
                                </TableCell>
                            </TableRow>
                        )}
                    </TableBody>
                </Table>
            </div>

            {/* Edit/Add Detail Modal */}
            <Modal isOpen={isOpen} onClose={onClose} size="2xl">
                <ModalContent>
                    <ModalHeader>
                        {editingDetail?.id === 0 ? 'Thêm chi tiết mới' : 'Chỉnh sửa chi tiết'}
                    </ModalHeader>
                    <ModalBody>
                        {editingDetail && (
                            <div className="grid grid-cols-2 gap-4">
                                <Input
                                    label="Tên trường"
                                    value={editingDetail.fieldName}
                                    onChange={(e) => setEditingDetail({ ...editingDetail, fieldName: e.target.value })}
                                />
                                <Input
                                    label="Tên hiển thị"
                                    value={editingDetail.displayName}
                                    onChange={(e) => setEditingDetail({ ...editingDetail, displayName: e.target.value })}
                                />
                                <Input
                                    type="number"
                                    label="Vị trí cột"
                                    value={editingDetail.columnPosition?.toString()}
                                    onChange={(e) => setEditingDetail({ ...editingDetail, columnPosition: parseInt(e.target.value) || 0 })}
                                />
                                <Input
                                    type="number"
                                    label="Vị trí hàng"
                                    value={editingDetail.rowPosition?.toString()}
                                    onChange={(e) => setEditingDetail({ ...editingDetail, rowPosition: parseInt(e.target.value) || 0 })}
                                />
                                <Select
                                    label="Kiểu dữ liệu"
                                    selectedKeys={[editingDetail.dataType?.toString()]}
                                    onChange={(e) => setEditingDetail({ ...editingDetail, dataType: parseInt(e.target.value) as DataTypes })}
                                >
                                    {Object.entries(dataTypeLabels).map(([key, value]) => (
                                        <SelectItem key={key} textValue={key}>
                                            {value}
                                        </SelectItem>
                                    ))}
                                </Select>
                                <div className="flex items-center">
                                    <Checkbox
                                        isSelected={editingDetail.isRequired}
                                        onValueChange={(checked) => setEditingDetail({ ...editingDetail, isRequired: checked })}
                                    >
                                        Bắt buộc
                                    </Checkbox>
                                </div>
                            </div>
                        )}
                    </ModalBody>
                    <ModalFooter>
                        <Button color="danger" variant="light" onPress={onClose}>
                            Hủy
                        </Button>
                        <Button color="primary" onPress={handleSaveDetail}>
                            Lưu
                        </Button>
                    </ModalFooter>
                </ModalContent>
            </Modal>
        </div>
    );
}