using System.Linq.Expressions;

namespace Flexcel.Internal;

public interface IExcelSheet<TSourceDocument>
{
    void AddColumn<TValue>(Expression<Func<TSourceDocument, TValue>> mapFunction, string? columnTitle = null);
    void AddRow(TSourceDocument row);
    void AddRows(IEnumerable<TSourceDocument> rows);
    void InitColumns();
}