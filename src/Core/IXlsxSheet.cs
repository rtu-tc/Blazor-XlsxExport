using ClosedXML.Excel;

namespace RtuTc.BlazorXlsxExport;
/// <summary>
/// Интерфейс листа таблицы, который можно заполнить данными. 
/// Необходим, так как теоретически страницы могут содержать разеородные данные и лист не должен быть generic
/// Интерфейс объявлен как internal для сокрытия используемой библиотеки для работы с excel
/// </summary>
internal interface IXlsxSheet
{
    string Name { get; }
    Task AddSheetToWorkbook(XLWorkbook xLWorkbook);
}

