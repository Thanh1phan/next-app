import ExcelImporter from "@/components/upload-button";

export default function ExcelExport() {
    return (
        <div>
            <ExcelImporter configId={1009} name="Upload bảng lương" ></ExcelImporter >
            <ExcelImporter configId={1010} name="Upload bảng lương" ></ExcelImporter >
        </div>
    )
}