using System.Linq.Expressions;
using Flexcel.Extensions;

namespace Flexcel.Internal;

public class ExcelColumn<TRowDocument, TValue> : IExcelColumn<TRowDocument, TValue>
{
    private readonly string? columnTitle;
    private readonly Func<TRowDocument, TValue> extractFunc;
    
    public ExcelColumn(Expression<Func<TRowDocument, TValue>> extractFunc, string? columnTitle = null)
    {
        this.extractFunc = extractFunc.Compile();
        this.columnTitle = columnTitle ?? extractFunc.GetMappedParameterName();
    }

    public string? GetTitle() => columnTitle;
    
    public TValue? GetValue(TRowDocument source)
    {
        return source != null ? extractFunc(source) : default;
    }

    object? IExcelColumn<TRowDocument>.GetValue(TRowDocument source)
    {
        return source != null ? extractFunc(source) : null;
    }
}