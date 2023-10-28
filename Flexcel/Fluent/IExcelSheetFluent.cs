using System.Linq.Expressions;

namespace Flexcel.Fluent;

public interface IExcelSheetFluent
{
    Guid GetId();
    string GetTittle();
    void InitColumns();
    void ApplySettings();
}

public interface IExcelSheetFluent<TRowDocument> : IExcelSheetFluent, IExcelExportable
{
    IExcelSheetFluent<TNewRowDocument> AddSheetAndSelect<TNewRowDocument>(string? sheetTitle = null);
    IExcelSheetFluent<TNewRowDocument> AddSheetAndSelect<TNewRowDocument>(string sheetTitle, out Guid sheetId);
    IExcelConfiguredColumnFluent<TRowDocument> AddColumn<TValue>(Expression<Func<TRowDocument, TValue>> mapFunc, string? columnTitle = null);
    IExcelSheetFluent<TRowDocument> AddColumns(IEnumerable<Expression<Func<TRowDocument, object>>> columns);
    IExcelSheetWithDataFluent<TRowDocument> AddRow(TRowDocument row);
    IExcelSheetWithDataFluent<TRowDocument> AddRows(IEnumerable<TRowDocument> rows);
}