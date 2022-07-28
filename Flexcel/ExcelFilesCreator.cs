
using Flexcel.Fluent;

namespace Flexcel;

public static class ExcelFilesCreator
{
    public static IExcelFluent NewExcel() => new ExcelFluent();
}