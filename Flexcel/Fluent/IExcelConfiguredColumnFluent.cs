namespace Flexcel.Fluent;

public interface IExcelConfiguredColumnFluent<TRowDocument> : IExcelSheetFluent<TRowDocument>
{
    IExcelConfiguredColumnFluent<TRowDocument> SetWidth(int minWidth);
    IExcelConfiguredColumnFluent<TRowDocument> SetAutoFit(bool autoFit = true);
    IExcelConfiguredColumnFluent<TRowDocument> SetNumberFormat(string format);
}