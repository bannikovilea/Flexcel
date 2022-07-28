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

    public void AddColumn<TValue>(Expression<Func<TSourceDocument, TValue>> mapFunction, string? columnTitle)
    {
        if (_currentRow != 0)
            throw new Exception("Sheet already has rows");
        _columns.Add(new ExcelColumn<TSourceDocument, TValue>(mapFunction, columnTitle));
    }

    public void AddRow(TSourceDocument row)
    {
        InitColumnsIfNeeded();

        var columnNumber = 0;
        foreach (var column in _columns)
        {
            _worksheet.Cells[AddressHelper.GetCellAddress(columnNumber, _currentRow)].Value = column.GetValue(row);
            _worksheet.Cells[AddressHelper.GetCellAddress(columnNumber, _currentRow)].Style.Numberformat.Format = column.CellFormat();
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