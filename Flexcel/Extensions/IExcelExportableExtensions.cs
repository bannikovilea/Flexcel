using Flexcel.Fluent;

namespace Flexcel.Extensions;

public static class ExcelExportableExtensions
{
    public static byte[] ToByteArray(this IExcelExportable fluent) => fluent.GetPackage().GetAsByteArray();
    
    public static void ToFile(this IExcelExportable fluent, string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        var bytes = fluent.ToByteArray();
        fs.Write(bytes, 0, bytes.Length);
    }
    
    public static Task ToFileAsync(this IExcelExportable fluent, string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        var bytes = fluent.ToByteArray();
        return fs.WriteAsync(bytes, 0, bytes.Length);
    }

    public static MemoryStream ToMemoryStream(this IExcelExportable fluent) => new MemoryStream(fluent.ToByteArray());
}
