using System.Text.RegularExpressions;
using Flexcel.Internal;
using Flexcel.Tests.Infrastructure;
using NUnit.Framework;
using AutoFixture;

namespace Flexcel.Tests.Internal;

[Parallelizable]
public class AddressHelperTests : TestBaseSimple
{
    private static readonly Regex OnlyLettersAndDigits = new Regex("^[A-Z0-9]+$", RegexOptions.Compiled);
    
    [Test]
    [TestCase(0, 0, "A1")]
    [TestCase(0, 1, "A2")]
    [TestCase(25, 1, "Z2")]
    [TestCase(0, 50, "A51")]
    [TestCase(1, 0, "B1")]
    [TestCase(26, 0, "AA1")]
    [TestCase(30, 0, "AE1")]
    [TestCase(30, 30, "AE31")]
    [TestCase(1169, 18423, "ARZ18424")]
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

    [Test]
    [Parallelizable]
    public void GetCellAddress_Always_ShouldReturnAddressWithOnlyDigitsAndLetters()
    {
        for (var column = 0; column < 10000; ++column) {
            var row = Math.Abs(fixture.Create<int>());
            var actual = AddressHelper.GetCellAddress(column, row);
            
            Assert.IsTrue(OnlyLettersAndDigits.IsMatch(actual));
        }
    }
}