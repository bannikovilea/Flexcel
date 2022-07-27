namespace Flexcel.Fluent;

public interface IExcelFluent : IExcelExportable
{
    /// <summary>
    /// Добавляет новый лист
    /// </summary>
    /// <param name="sheetTitle">Название листа</param>
    /// <typeparam name="TRowDocument">Модель данных</typeparam>
    /// <returns></returns>
    IExcelSheetFluent<TRowDocument> AddSheetAndSelect<TRowDocument>(string? sheetTitle = null);
    /// <summary>
    /// Добавляет новый лист
    /// </summary>
    /// <param name="sheetTitle">Название листа</param>
    /// <param name="sheetId">ИД листа, если потом понадобится к нему доступ</param>
    /// <typeparam name="TRowDocument">Модель данных</typeparam>
    /// <returns></returns>
    IExcelSheetFluent<TRowDocument> AddSheetAndSelect<TRowDocument>(string sheetTitle, out Guid sheetId);
    
    /// <summary>
    /// Получает лист по ИД
    /// </summary>
    /// <param name="id"></param>
    /// <returns></returns>
    IExcelSheetFluent? GetSheet(Guid id);
    /// <summary>
    /// Получает лист по id и пробует привести его к указанному генерику
    /// </summary>
    /// <param name="id"></param>
    /// <typeparam name="TRowDocument"></typeparam>
    /// <returns></returns>
    IExcelSheetFluent<TRowDocument>? GetSheet<TRowDocument>(Guid id);
}