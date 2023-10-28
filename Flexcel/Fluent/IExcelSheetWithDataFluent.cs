namespace Flexcel.Fluent;

public interface IExcelSheetWithDataFluent<TRowDocument> : IExcelSheetFluent, IExcelExportable
{
    IExcelSheetFluent<TNewRowDocument> AddSheetAndSelect<TNewRowDocument>(string? sheetTitle = null);
    IExcelSheetFluent<TNewRowDocument> AddSheetAndSelect<TNewRowDocument>(string sheetTitle, out Guid sheetId);
    IExcelSheetWithDataFluent<TRowDocument> AddRow(TRowDocument row);
    IExcelSheetWithDataFluent<TRowDocument> AddRows(IEnumerable<TRowDocument> rows);
    IExcelSheetWithDataFluent<TRowDocument> AutoFitAllColumnsInSheet();
}