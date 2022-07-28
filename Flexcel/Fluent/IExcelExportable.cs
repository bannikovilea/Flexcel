using OfficeOpenXml;

namespace Flexcel.Fluent;

public interface IExcelExportable
{
    ExcelPackage GetPackage();
}