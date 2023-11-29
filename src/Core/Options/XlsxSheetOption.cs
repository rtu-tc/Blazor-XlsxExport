using ClosedXML.Excel;

using Microsoft.AspNetCore.Components;

namespace RtuTc.BlazorXlsxExport.Options;
[CascadingTypeParameter(nameof(TItem))]
public abstract class XlsxSheetOption<TItem> : ComponentBase, IDisposable
{
    [CascadingParameter]
    public XlsxSheet<TItem>? Sheet { get; set; }

    internal abstract void ApplyOption(IXLWorksheet worksheet);

    protected override void OnInitialized()
    {
        base.OnInitialized();
        if (Sheet is null)
        {
            throw new InvalidOperationException($"{nameof(XlsxSheetOption<TItem>)} can be used only in child content of {nameof(XlsxSheet<TItem>)}");
        }
        Sheet.AddOption(this);
    }

    public void Dispose()
    {
        Sheet?.RemoveOption(this);
    }
}
