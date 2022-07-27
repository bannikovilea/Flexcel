using Flexcel.Internal;
using Flexcel.Tests.Infrastructure;
using NUnit.Framework;

namespace Flexcel.Tests.Internal;

[Parallelizable]
public class AddressHelperTests : TestBaseSimple
{
    [Test]
    [TestCase(0, 0, "A1")]
    [TestCase(0, 1, "A2")]
    [TestCase(0, 50, "A51")]
    [TestCase(1, 0, "B1")]
    [TestCase(30, 0, "AE1")]
    [TestCase(30, 30, "AE31")]
    public void GetCellAddress_Always_ShouldReturnCorrectAddress(int col, long row, string expectedAddress)
    {
        Assert.That(AddressHelper.GetCellAddress(col, row), Is.EqualTo(expectedAddress));
    }

    [Test]
    [TestCase(-1, 0)]
    [TestCase(-1, 5)]
    [TestCase(0, -1)]
    [TestCase(5, -1)]
    [TestCase(-1, -1)]
    public void GetCellAddress_InvalidArgs_ShouldThrow(int col, long row)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => AddressHelper.GetCellAddress(col, row));
    }
}