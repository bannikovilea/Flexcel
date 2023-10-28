using System.Linq.Expressions;
using System.Reflection;
using OfficeOpenXml;

namespace Flexcel.Internal;

internal class ExcelSheet<TSourceDocument> : IExcelSheet<TSourceDocument>
{
    private readonly ExcelWorksheet _worksheet;
    private readonly List<IExcelColumn<TSourceDocument>> _columns;
    private bool _columnsInited;
    private long _currentRow;

    public ExcelSheet(ExcelPackage xlsx, string sheetTitle)
    {
        _worksheet = xlsx.Workbook.Worksheets.Add(sheetTitle);
        _columns = new List<IExcelColumn<TSourceDocument>>();
    }

    public IExcelColumn<TSourceDocument> AddColumn<TValue>(Expression<Func<TSourceDocument, TValue>> mapFunction, string? columnTitle)
    {
        if (_currentRow != 0)
            throw new Exception("Sheet already has rows");
        _columns.Add(new ExcelColumn<TSourceDocument, TValue>(mapFunction, columnTitle));
        return _columns.Last();
    }

    public void AddRow(TSourceDocument row)
    {
        InitColumnsIfNeeded();

        var columnNumber = 0;
        foreach (var column in _columns)
        {
            _worksheet.Cells[AddressHelper.GetCellAddress(columnNumber, _currentRow)].Value = column.GetValue(row);
            ++columnNumber;
        }

        ++_currentRow;
    }

    public void AddRows(IEnumerable<TSourceDocument> rows)
    {
        foreach (var row in rows)
            AddRow(row);
    }

    public void InitColumns()
    {
        if (_columnsInited)
            return;
        
        var columnNumber = 0;
        if (!_columns.Any())
            InitColumnsByType();
        
        foreach (var title in _columns.Select(c => c.GetTitle()))
        {
            _worksheet.Cells[AddressHelper.GetCellAddress(columnNumber, 0)].Value = title;
            ++columnNumber;
        }

        _columnsInited = true;
    }

    public void ApplySettings()
    {
        var columnNumber = 0;
        foreach (var column in _columns)
        {
            var settings = column.GetSettings();
            var startCellRow = _currentRow > 0 ? 1 : _currentRow;
            _worksheet.Cells[$"{AddressHelper.GetCellAddress((int)startCellRow, columnNumber)}:{AddressHelper.GetCellAddress((int)_currentRow, columnNumber)}"].Style.Numberformat.Format = settings.CellFormat;
            
            if (settings.Width.HasValue)
                _worksheet.Column(columnNumber).Width = settings.Width.Value;
            if (settings.AutoFit == true)
                _worksheet.Column(columnNumber).AutoFit();
            
            ++columnNumber;
        }
    }

    public void AutoFitAllColumns()
    {
        _worksheet.Cells.AutoFitColumns();
    }

    private void InitColumnsIfNeeded()
    {
        if (_currentRow != 0)
            return;

        InitColumns();

        ++_currentRow;
    }

    private void InitColumnsByType()
    {
        foreach (var member in typeof(TSourceDocument).GetMembers())
        {
            switch (member.MemberType)
            {
                case MemberTypes.Field:
                    _columns.Add(new ExcelColumnByReflectionType<TSourceDocument>((FieldInfo)member));
                    break;
                case MemberTypes.Property:
                    _columns.Add(new ExcelColumnByReflectionType<TSourceDocument>((PropertyInfo)member));
                    break;
                default:
                    continue;
            }
            
        }
    }
}