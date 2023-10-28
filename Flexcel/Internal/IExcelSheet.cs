using System.Linq.Expressions;

namespace Flexcel.Internal;

public interface IExcelSheet<TSourceDocument>
{
    IExcelColumn<TSourceDocument> AddColumn<TValue>(Expression<Func<TSourceDocument, TValue>> mapFunction, string? columnTitle = null);
    void AddRow(TSourceDocument row);
    void AddRows(IEnumerable<TSourceDocument> rows);
    void InitColumns();
    void ApplySettings();
    void AutoFitAllColumns();
}