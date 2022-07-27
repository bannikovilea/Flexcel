using System.Reflection;

namespace Flexcel.Internal;

public class ExcelColumnByReflectionType<TRowDocument> : IExcelColumn<TRowDocument, object?>
{
    private readonly string title;
    private readonly Func<TRowDocument, object?> extractFunc;
    
    public ExcelColumnByReflectionType(FieldInfo fieldInfo)
    {
        title = fieldInfo.Name;
        extractFunc = document => fieldInfo.GetValue(document);
    }
    
    public ExcelColumnByReflectionType(PropertyInfo fieldInfo)
    {
        title = fieldInfo.Name;
        extractFunc = document => fieldInfo.GetValue(document);
    }

    public string? GetTitle() => title;

    public object? GetValue(TRowDocument source) 
        => source != null 
               ? extractFunc(source) 
               : null;

    object? IExcelColumn<TRowDocument>.GetValue(TRowDocument source) => GetValue(source);
}