using OfficeOpenXml;

namespace Flexcel.Fluent;

internal class ExcelFluent : IExcelFluent
{
    private const string DefaultTitle = "Sheet";
    private readonly ExcelPackage currentXLSX;
    private readonly Dictionary<Guid, IExcelSheetFluent> sheets;

    public ExcelFluent()
    {
        currentXLSX = new ExcelPackage();
        sheets = new Dictionary<Guid, IExcelSheetFluent>();
    }
    
    public IExcelSheetFluent<TRowDocument> AddSheetAndSelect<TRowDocument>(string? sheetTitle = null)
    {
        var sheetId = Guid.NewGuid();
        var sheet = new ExcelSheetFluent<TRowDocument>(this, sheetTitle ?? $"{DefaultTitle} {sheets.Count + 1}", sheetId);
        sheets.Add(sheetId, sheet);
        return sheet;
    }

    public IExcelSheetFluent<TRowDocument> AddSheetAndSelect<TRowDocument>(string sheetTitle, out Guid sheetId)
    {
        var sheet = AddSheetAndSelect<TRowDocument>(sheetTitle);
        sheetId = sheet.GetId();
        return sheet;
    }

    public IExcelSheetFluent? GetSheet(Guid id) => sheets.ContainsKey(id) ? sheets[id] : null;
    
    public IExcelSheetFluent<TRowDocument>? GetSheet<TRowDocument>(Guid id) 
        => GetSheet(id) as IExcelSheetFluent<TRowDocument>;

    public ExcelPackage GetPackage()
    {
        foreach (var sheet in sheets.Values)
        {
            sheet.InitColumns();
            sheet.ApplySettings();
        }

        return currentXLSX;   
    }
}