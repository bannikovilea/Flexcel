using System.Globalization;
using System.Linq.Expressions;
using Flexcel.Extensions;

namespace Flexcel.Internal;

internal class ExcelColumn<TRowDocument, TValue> : IExcelColumn<TRowDocument, TValue>
{
    private readonly string? _columnTitle;
    private readonly Func<TRowDocument, TValue> _extractFunc;
    private readonly string? _cellFormat;
    
    public ExcelColumn(Expression<Func<TRowDocument, TValue>> extractFunc, string? columnTitle = null)
    {
        this._extractFunc = extractFunc.Compile();
        this._columnTitle = columnTitle ?? extractFunc.GetMappedParameterName();
        _cellFormat =
            DataTypeFormatHelper.GetDefaultDateFormatIfDateType(extractFunc.GetMappedParameterType() ?? typeof(TValue));
    }

    public string? GetTitle() => _columnTitle;
    
    public TValue? GetValue(TRowDocument source) 
        => source != null 
            ? _extractFunc(source) 
            : default;

    public string? CellFormat() => _cellFormat;

    object? IExcelColumn<TRowDocument>.GetValue(TRowDocument source)
    {
        return source != null ? _extractFunc(source) : null;
    }
}