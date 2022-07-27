using Flexcel.Fluent;

namespace Flexcel.Extensions;

public static class IExcelExportableExtensions
{
    public static void ToFile(this IExcelExportable fluent, string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        var bytes = fluent.GetPackage().GetAsByteArray();
        fs.Write(bytes, 0, bytes.Length);
    }
    
    public static Task ToFileAsync(this IExcelExportable fluent, string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        var bytes = fluent.GetPackage().GetAsByteArray();
        return fs.WriteAsync(bytes, 0, bytes.Length);
    }

    public static byte[] ToByteArray(this IExcelExportable fluent) => fluent.GetPackage().GetAsByteArray();
}