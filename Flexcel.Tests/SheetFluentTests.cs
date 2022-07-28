using System.Linq.Expressions;
using Flexcel.Fluent;
using Flexcel.Tests.Infrastructure;
using NUnit.Framework;
using AutoFixture;
using Flexcel.Extensions;
using Flexcel.Internal;

namespace Flexcel.Tests;

public class SheetFluentTests : TestBaseSimple
{
    public override void SetUp()
    {
        base.SetUp();
        sut = ExcelFilesCreator.NewExcel().AddSheetAndSelect<TestExportClass>();
    }

    private IExcelSheetFluent<TestExportClass> sut;

    [Test]
    public void AddSheetAndSelect_EmptyName_ShouldThrow()
    {
        Assert.Throws<ArgumentException>(() => sut.AddSheetAndSelect<TestExportClass>(""));
    }

    [Test]
    public void AddSheetAndSelect_DefaultName_ShouldGenerateTitle()
    {
        var actual = sut.AddSheetAndSelect<TestExportClass>();
        
        Assert.That(actual.GetTittle(), Is.EqualTo("Sheet 2"));
    }

    [Test]
    public void AddSheetAndSelect_Always_ShouldGenerateSheetId()
    {
        var sheetTitle = fixture.Create<string>();
        var actual = sut.AddSheetAndSelect<TestExportClass>(sheetTitle, out var sheetId);
        
        Assert.That(sheetId, Is.Not.EqualTo(Guid.Empty));
        Assert.That(sheetId, Is.EqualTo(actual.GetId()));
        Assert.That(actual.GetTittle(), Is.EqualTo(sheetTitle));
    }
    
    [Test]
    public void AddSheetAndSelect_FromOtherSheet_ShouldGenerateSheetId()
    {
        var sheetTitle = fixture.Create<string>();
        var actual = sut.AddSheetAndSelect<TestExportClass>(fixture.Create<string>(), out var firstSheetId)
            .AddSheetAndSelect<TestExportClass>(sheetTitle, out var secondSheetId);
        
        Assert.That(firstSheetId, Is.Not.EqualTo(Guid.Empty));
        Assert.That(secondSheetId, Is.Not.EqualTo(Guid.Empty));
        Assert.That(firstSheetId, Is.Not.EqualTo(secondSheetId));
        Assert.That(actual.GetTittle(), Is.EqualTo(sheetTitle));
        Assert.That(actual.GetId(), Is.EqualTo(secondSheetId));
    }

    [Test]
    [TestCaseSource(nameof(AvailableFieldsWithCaption))]
    public void AddColumn_WithoutTitle_ShouldSetTitleByName(Expression<Func<TestExportClass, object>> selector, string expectedName)
    {
        var actual = sut.AddColumn(selector)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells[AddressHelper.GetCellAddress(0, 0)]
            .Value;
        
        Assert.That(actual, Is.EqualTo(expectedName));
    }
    
    [Test]
    [TestCaseSource(nameof(AvailableFields))]
    public void AddColumn_WithTitle_ShouldSetTitle(Expression<Func<TestExportClass, object>> selector)
    {
        var columnName = fixture.Create<string>();
        var actual = sut.AddColumn(selector, columnName)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells[AddressHelper.GetCellAddress(0, 0)]
            .Value;
        
        Assert.That(actual, Is.EqualTo(columnName));
    }

    [Test]
    public void AddColumns_Always_ShouldUseParameterName()
    {
        var availableFields = AvailableFields().ToArray();
        var actualCells = sut.AddColumns(availableFields)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells;

        for (var i = 0; i < availableFields.Length; ++i)
        {
            var expectedColumnName = availableFields[i].GetMappedParameterName();
            var actualColumnName = actualCells[AddressHelper.GetCellAddress(i, 0)].Value;
            Assert.That(actualColumnName, Is.EqualTo(expectedColumnName));
        }
    }
    
    [Test]
    [TestCaseSource(nameof(AvailableFields))]
    public void AddRow_ColumnAlreadyAdded_ShouldMapValue(Expression<Func<TestExportClass, object>> selector)
    {
        var testItem = fixture.Create<TestExportClass>();
        var compiledSelector = selector.Compile();
        
        var actual = sut.AddColumn(selector)
            .AddRow(testItem)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells[AddressHelper.GetCellAddress(0, 1)]
            .Value;

        Assert.That(actual, Is.EqualTo(compiledSelector(testItem)));
    }
    
    [Test]
    [TestCaseSource(nameof(AvailableFields))]
    public void AddRow_ColumnAlreadyAddedAndRowIsNull_ShouldMapValue(Expression<Func<TestExportClass, object>> selector)
    {
        var actual = sut.AddColumn(selector)
            .AddRow(null)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells[AddressHelper.GetCellAddress(0, 1)]
            .Value;

        Assert.IsNull(actual);
    }
    
    [Test]
    [TestCaseSource(nameof(AvailableFields))]
    public void AddRows_ColumnAlreadyAdded_ShouldMapValue(Expression<Func<TestExportClass, object>> selector)
    {
        var testItems = new []
        {
            fixture.Create<TestExportClass>(),
            fixture.Create<TestExportClass>(),
            null,
            fixture.Create<TestExportClass>(),
            null
        };
        var compiledSelector = selector.Compile();
        
        var actualCells = sut.AddColumn(selector)
            .AddRows(testItems)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells;

        for (var i = 0; i < testItems.Length; ++i)
        {
            var actual = actualCells[AddressHelper.GetCellAddress(0, i + 1)].Value;
            if (testItems[i] == null)
                Assert.IsNull(actual);
            else
                Assert.That(actual, Is.EqualTo(compiledSelector(testItems[i])));
        }
    }
    
    [Test]
    public void AddRow_WithoutAddedColumnsAndRowIsNull_ShouldMapAllExistingFields()
    {
        var selectors = AvailableFields().Select(f => f.Compile()).ToArray();

        var actualCells = sut.AddRow(null)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells;

        for (var i = 0; i < selectors.Length; ++i)
        {
            var actual = actualCells[AddressHelper.GetCellAddress(i, 1)].Value;
            Assert.IsNull(actual);
        }
    }
    
    [Test]
    public void AddRow_WithoutAddedColumns_ShouldMapAllExistingFields()
    {
        var testItem = fixture.Create<TestExportClass>();
        var selectors = AvailableFields().Select(f => f.Compile()).ToArray();

        var actualCells = sut.AddRow(testItem)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells;

        for (var i = 0; i < selectors.Length; ++i)
        {
            var actual = actualCells[AddressHelper.GetCellAddress(i, 1)].Value;
            var expected = selectors[i](testItem);
            Assert.That(actual, Is.EqualTo(expected));
        }
    }
    
    [Test]
    public void AddRows_WithoutAddedColumns_ShouldMapAllExistingFields()
    {
        var testItems = new []
        {
            fixture.Create<TestExportClass>(),
            null,
            null,
            fixture.Create<TestExportClass>(),
            null
        };
        var selectors = AvailableFields().Select(f => f.Compile()).ToArray();

        var actualCells = sut.AddRows(testItems)
            .GetPackage()
            .Workbook
            .Worksheets[0]
            .Cells;

        for (var rowNumber = 0; rowNumber < testItems.Length; ++rowNumber)
        {
            var testItem = testItems[rowNumber];
            for (var columnNumber = 0; columnNumber < selectors.Length; ++columnNumber)
            {
                var actual = actualCells[AddressHelper.GetCellAddress(columnNumber, rowNumber + 1)].Value;
                var expected = testItem == null ? null : selectors[columnNumber](testItem);
                Assert.That(actual, Is.EqualTo(expected));
            }
        }
    }

    private static IEnumerable<Expression<Func<TestExportClass, object>>> AvailableFields()
    {
        yield return d => d.LambdaField;
        yield return d => d.SomeString;
        yield return d => d.SomeBool;
        yield return d => d.DateTime;
        yield return d => d.DateTimeOffset;
        yield return d => d.Constant;
        yield return _ => TestExportClass.StaticConstant;
    }

    private static IEnumerable<TestCaseData> AvailableFieldsWithCaption()
    {
        return AvailableFields().Select(f => new TestCaseData(f, f.GetMappedParameterName()));
    }
}