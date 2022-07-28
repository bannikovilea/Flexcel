namespace Flexcel.Internal;

public static class DateTypeHelper
{
    private const string DefaultExportFormat = "yyyy-MM-dd HH:mm:ss";
    
    public static bool IsDateType(Type type) 
        => type == typeof(DateTime);

    public static string? GetDefaultDateFormatIfDateType(Type type)
        => IsDateType(type)
            ? DefaultExportFormat
            : null;
}