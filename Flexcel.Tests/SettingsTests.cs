using AutoFixture;
using NUnit.Framework;

namespace Flexcel.Tests;

public partial class SheetFluentTests
{
    [Test]
    public void ApplySettings_CellFormat_ShouldSet()
    {
        var testItems = fixture.CreateMany<TestExportClass>(30).ToList();

        var actualCells = sut.AddColumn(x => x.DateTime)
            .SetNumberFormat("Date")
            .AddColumn(x => x.DateTimeOffset)
            .SetNumberFormat("Text")
            .AddColumn(x => x.SomeString)
            .SetAutoFit()
            .AddRows(testItems)
            .GetPackage().Workbook
            .Worksheets[0]
            .Cells;
        
        for (var i = 0; i < testItems.Count; ++i)
        {
            Assert.That(actualCells[$"A{i+2}"].Style.Numberformat.Format, Is.EqualTo("Date"));
            Assert.That(actualCells[$"B{i+2}"].Style.Numberformat.Format, Is.EqualTo("Text"));
            Assert.That(actualCells[$"C{i+2}"].Style.Numberformat.Format, Is.EqualTo("General"));
        }
    }
}