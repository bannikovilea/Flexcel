using System.Linq.Expressions;

namespace Flexcel.Extensions;

public static class ExpressionExtensions
{
    public static string? GetMappedParameterName<TSourceDoc, TValue>(this Expression<Func<TSourceDoc, TValue>> expression)
    {
        var expressionBody = expression.Body;
        if (expressionBody is UnaryExpression { NodeType: ExpressionType.Convert } unaryExpression)
        {
            expressionBody = unaryExpression.Operand;
        }
        return (expressionBody as MemberExpression)?.Member.Name;
    }
}