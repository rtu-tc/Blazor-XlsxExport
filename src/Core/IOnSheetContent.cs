using ClosedXML.Excel;

namespace RtuTc.BlazorXlsxExport;
internal interface IOnSheetContent
{
    /// <summary>
    /// Renders content on sheet
    /// </summary>
    /// <param name="worksheet">target sheet</param>
    /// <param name="columnIndexStart">column start index</param>
    /// <param name="rowIndexStart">row start index</param>
    /// <returns>Coordinates of next cell to render</returns>
    Task<(int rowIndex, int columnIndex)> RenderContent(IXLWorksheet worksheet, int rowIndexStart, int columnIndexStart);
}
