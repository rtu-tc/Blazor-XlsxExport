using ClosedXML.Excel;

namespace RtuTc.BlazorXlsxExport;
/// <summary>
/// Интерфейс листа таблицы, который можно заполнить данными. 
/// Необходим, так как теоретически страницы могут содержать разеородные данные и лист не должен быть generic
/// Интерфейс объявлен как internal для сокрытия используемой библиотеки для работы с excel
/// </summary>
internal interface IXlsxSheet
{
    Task AddSheetToWorkbook(XLWorkbook xLWorkbook);
}

internal class IXlsxWrapper : IXlsxSheet
{
    private readonly Func<XLWorkbook, Task> _action;

    public IXlsxWrapper(Func<XLWorkbook, Task> action)
    {
        _action = action;
    }

    public async Task AddSheetToWorkbook(XLWorkbook xLWorkbook)
    {
        await _action(xLWorkbook);
    }
}
