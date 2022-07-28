using Flexcel.Internal;
using Flexcel.Tests.Infrastructure;
using NUnit.Framework;

namespace Flexcel.Tests.Internal;

[Parallelizable]
public class DateFormatHelperTests : TestBaseSimple
{
    [Test]
    [TestCase(typeof(DateTime), true)]
    [TestCase(typeof(DateTimeOffset), false)]
    [TestCase(typeof(object), false)]
    [TestCase(typeof(int), false)]
    [TestCase(typeof(long), false)]
    [TestCase(typeof(float), false)]
    [TestCase(typeof(string), false)]
    [TestCase(typeof(TestExportClass), false)]
    public void IsDateType_Always_ReturnValid(Type type, bool expected)
    {
        Assert.That(DateTypeHelper.IsDateType(type), Is.EqualTo(expected));
    }
}