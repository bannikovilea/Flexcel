using System.Linq.Expressions;
using System.Reflection;

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
    
    public static Type? GetMappedParameterType<TSourceDoc, TValue>(this Expression<Func<TSourceDoc, TValue>> expression)
    {
        var expressionBody = expression.Body;
        if (expressionBody is UnaryExpression { NodeType: ExpressionType.Convert } unaryExpression)
        {
            expressionBody = unaryExpression.Operand;
        }
        var member = (expressionBody as MemberExpression)?.Member;
        return member?.MemberType switch
        {
            MemberTypes.Property => (member as PropertyInfo)?.PropertyType,
            MemberTypes.Field => (member as FieldInfo)?.FieldType,
            _ => null
        };
    }
}