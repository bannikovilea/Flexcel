using System.Linq.Expressions;
using Flexcel.Internal;
using OfficeOpenXml;

namespace Flexcel.Fluent;

internal class ExcelSheetFluent<TRowDocument> : IExcelSheetFluent<TRowDocument>, IExcelSheetWithDataFluent<TRowDocument>
{
    private readonly IExcelFluent _parent;
    private readonly Guid _id;
    private readonly string _title;
    private readonly IExcelSheet<TRowDocument> _sheet;

    public ExcelSheetFluent(IExcelFluent parent, string title, Guid id)
    {
        _parent = parent;
        _id = id;
        _title = title;
        _sheet = new ExcelSheet<TRowDocument>(parent.GetPackage(), title);
    }

    public IExcelSheetFluent<TNewRowDocument> AddSheetAndSelect<TNewRowDocument>(string? sheetTitle = null) 
        => _parent.AddSheetAndSelect<TNewRowDocument>(sheetTitle);

    public IExcelSheetFluent<TNewRowDocument> AddSheetAndSelect<TNewRowDocument>(string sheetTitle, out Guid sheetId)
        => _parent.AddSheetAndSelect<TNewRowDocument>(sheetTitle, out sheetId);

    public ExcelPackage GetPackage() => _parent.GetPackage();

    public IExcelSheetFluent<TRowDocument> AddColumn<TValue>(Expression<Func<TRowDocument, TValue>> mapFunc, string? columnTitle = null)
    {
        _sheet.AddColumn(mapFunc, columnTitle);
        return this;
    }

    public IExcelSheetFluent<TRowDocument> AddColumns(IEnumerable<Expression<Func<TRowDocument, object>>> columns)
    {
        foreach (var column in columns)
            _sheet.AddColumn(column);

        return this;
    }

    public IExcelSheetWithDataFluent<TRowDocument> AddRow(TRowDocument row)
    {
        _sheet.AddRow(row);
        return this;
    }

    public IExcelSheetWithDataFluent<TRowDocument> AddRows(IEnumerable<TRowDocument> rows)
    {
        _sheet.AddRows(rows);
        return this;
    }

    public Guid GetId() => _id;
    
    public string GetTittle() => _title;

    public void InitColumns() => _sheet.InitColumns();
}