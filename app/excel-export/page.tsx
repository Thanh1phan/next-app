import UploadButton from "@/components/upload-button";

export default function ExcelExport() {
    return (
        <div>
            <UploadButton departmentId={1} configType={0} name="Upload bảng lương" color="success" ></UploadButton >
            <UploadButton departmentId={1} configType={1} name="Upload bảng lương" size="lg"></UploadButton >
        </div>
    )
}