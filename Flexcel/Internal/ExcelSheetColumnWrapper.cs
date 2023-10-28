using System.Linq.Expressions;
using Flexcel.Fluent;
using OfficeOpenXml;

namespace Flexcel.Internal;

internal class ExcelSheetColumnWrapper<TRowDocument> : IExcelConfiguredColumnFluent<TRowDocument>
{
    private readonly IExcelSheetFluent<TRowDocument> _sheet;
    private readonly IExcelColumn<TRowDocument> _column;

    public ExcelSheetColumnWrapper(IExcelSheetFluent<TRowDocument> sheet, ref IExcelColumn<TRowDocument> column)
    {
        _sheet = sheet;
        _column = column;
    }

    public Guid GetId() => _sheet.GetId();

    public string GetTittle() => _sheet.GetTittle();

    public void InitColumns() => _sheet.InitColumns();
    public void ApplySettings() => _sheet.ApplySettings();

    public ExcelPackage GetPackage() => _sheet.GetPackage();

    public IExcelSheetFluent<TNewRowDocument> AddSheetAndSelect<TNewRowDocument>(string? sheetTitle = null) 
        => _sheet.AddSheetAndSelect<TNewRowDocument>(sheetTitle);

    public IExcelSheetFluent<TNewRowDocument> AddSheetAndSelect<TNewRowDocument>(string sheetTitle, out Guid sheetId) 
        => _sheet.AddSheetAndSelect<TNewRowDocument>(sheetTitle, out sheetId);

    public IExcelConfiguredColumnFluent<TRowDocument> AddColumn<TValue1>(
        Expression<Func<TRowDocument, TValue1>> mapFunc, string? columnTitle = null)
        => _sheet.AddColumn(mapFunc, columnTitle);

    public IExcelSheetFluent<TRowDocument> AddColumns(IEnumerable<Expression<Func<TRowDocument, object>>> columns)
        => _sheet.AddColumns(columns);

    public IExcelSheetWithDataFluent<TRowDocument> AddRow(TRowDocument row) => _sheet.AddRow(row);

    public IExcelSheetWithDataFluent<TRowDocument> AddRows(IEnumerable<TRowDocument> rows) => _sheet.AddRows(rows);

    public IExcelConfiguredColumnFluent<TRowDocument> SetWidth(int minWidth)
    {
        _column.GetSettings().SetWidth(minWidth);
        return this;
    }

    public IExcelConfiguredColumnFluent<TRowDocument> SetAutoFit(bool autoFit)
    {
        _column.GetSettings().SetAutoFit(autoFit);
        return this;
    }

    public IExcelConfiguredColumnFluent<TRowDocument> SetNumberFormat(string format)
    {
        _column.GetSettings().SetFormat(format);
        return this;
    }
}