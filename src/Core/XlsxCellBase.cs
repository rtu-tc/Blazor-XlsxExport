using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Components;

namespace RtuTc.BlazorXlsxExport;
public abstract class XlsxCellBase : ComponentBase, IOnSheetContent, IDisposable
{
    [Parameter]
    public ContentDirection Next { get; set; }
    /// <summary>
    /// Count of cells to merge. 1 is nothing to merge
    /// </summary>
    [Parameter]
    public int MergeRight { get; set; }
    /// <summary>
    /// Count of cells to merge. 1 is nothing to merge
    /// </summary>
    [Parameter]
    public int MergeDown { get; set; }

    [Parameter]
    public bool Bold { get; set; }
    [Parameter]
    public bool Italic { get; set; }
    [Parameter]
    public double FontSize { get; set; } = 11;

    [Parameter]
    public XLAlignmentHorizontalValues HorizontalAlign { get; set; } = XLAlignmentHorizontalValues.Left;

    [Parameter]
    public XLAlignmentVerticalValues VerticalAlign { get; set; } = XLAlignmentVerticalValues.Bottom;

    [Parameter]
    public bool WrapText { get; set; }

    [Parameter]
    public int TextRotation { get; set; }

    [CascadingParameter]
    public XlsxSheet? Sheet { get; set; }
    [CascadingParameter]
    internal ReportProgressCounter? ReportProgressCounter { get; set; }

    private int MergeDownOffset => Math.Max(MergeDown - 1, 0);
    private int MergeRightOffset => Math.Max(MergeRight - 1, 0);

    async Task<(int rowIndex, int columnIndex)> IOnSheetContent.RenderContent(IXLWorksheet worksheet, int rowIndexStart, int columnIndexStart)
    {
        if (Sheet is null)
        {
            throw new InvalidOperationException($"{nameof(XlsxRichTextCell)} can be used only in child content of {nameof(XlsxSheet)}");
        }

        if (MergeRight > 1 || MergeDown > 1)
        {
            worksheet.Range(rowIndexStart, columnIndexStart, rowIndexStart + MergeDownOffset, columnIndexStart + MergeRightOffset).Merge();
        }

        var cell = worksheet.Cell(rowIndexStart, columnIndexStart);

        cell.Style
            .Font.SetBold(Bold)
            .Font.SetItalic(Italic)
            .Font.SetFontSize(FontSize)
            .Alignment.SetHorizontal(HorizontalAlign)
            .Alignment.SetVertical(VerticalAlign)
            .Alignment.SetWrapText(WrapText)
            .Alignment.SetTextRotation(TextRotation)
            ;

        await PlaceCellContent(cell);
        ReportProgressCounter?.ElementDone();
        return Next switch
        {
            ContentDirection.Down => (rowIndexStart + 1 + MergeDownOffset, columnIndexStart),
            ContentDirection.Right => (rowIndexStart, columnIndexStart + 1 + MergeRightOffset),
            _ => throw new ArgumentException($"Incorrect value {Next}", nameof(Next))
        };
    }

    protected override void OnInitialized()
    {
        base.OnInitialized();
        Sheet?.AddContent(this);
        ReportProgressCounter?.AddNewElement();

    }

    public void Dispose()
    {
        Sheet?.RemoveContent(this);
        ReportProgressCounter?.RemoveElement();
        GC.SuppressFinalize(this);
    }

    protected abstract ValueTask PlaceCellContent(IXLCell cell);

    public enum ContentDirection
    {
        /// <summary>
        /// ---
        /// -.-
        /// -+-
        /// </summary>
        Down,
        /// <summary>
        /// ---
        /// -.+
        /// ---
        /// </summary>
        Right,
    }
}
