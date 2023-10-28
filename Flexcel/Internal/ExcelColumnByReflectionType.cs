using System.Globalization;
using System.Reflection;

namespace Flexcel.Internal;

internal class ExcelColumnByReflectionType<TRowDocument> : IExcelColumn<TRowDocument, object?>
{
    private readonly string _title;
    private readonly Func<TRowDocument, object?> _extractFunc;
    private readonly string? _cellFormat;
    
    public ExcelColumnByReflectionType(FieldInfo fieldInfo)
    {
        _title = fieldInfo.Name;
        _extractFunc = document => fieldInfo.GetValue(document);
        _cellFormat = DataTypeFormatHelper.GetDefaultDateFormatIfDateType(fieldInfo.FieldType);
    }
    
    public ExcelColumnByReflectionType(PropertyInfo fieldInfo)
    {
        _title = fieldInfo.Name;
        _extractFunc = document => fieldInfo.GetValue(document);
        _cellFormat = DataTypeFormatHelper.GetDefaultDateFormatIfDateType(fieldInfo.PropertyType);
    }

    public string? GetTitle() => _title;

    public object? GetValue(TRowDocument source) 
        => source != null 
               ? _extractFunc(source) 
               : null;

    public string? CellFormat() => _cellFormat;

    object? IExcelColumn<TRowDocument>.GetValue(TRowDocument source) => GetValue(source);
}