using ClosedXML.Excel;

using Microsoft.AspNetCore.Components;

namespace RtuTc.BlazorXlsxExport.Options;

public abstract class XlsxSheetOption : ComponentBase, IDisposable
{
    [CascadingParameter]
    public XlsxSheet? Sheet { get; set; }

    internal abstract void ApplyOption(IXLWorksheet worksheet);

    protected override void OnInitialized()
    {
        base.OnInitialized();
        if (Sheet is null)
        {
            throw new InvalidOperationException($"{nameof(XlsxSheetOption)} can be used only in child content of {nameof(XlsxSheet)}");
        }
        Sheet.AddOption(this);
    }

    public void Dispose()
    {
        Sheet?.RemoveOption(this);
    }
}
