namespace Flexcel.Internal;

public interface IExcelColumn<TRowDocument>
{
    string? GetTitle();
    object? GetValue(TRowDocument source);
    ColumnSettings GetSettings();
}

public interface IExcelColumn<TRowDocument, TValue> : IExcelColumn<TRowDocument>
{
    new TValue? GetValue(TRowDocument source);
}