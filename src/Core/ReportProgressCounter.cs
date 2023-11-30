namespace RtuTc.BlazorXlsxExport;
internal class ReportProgressCounter
{
    private int elementsToHandleCount;
    private int done;

    public int ProgressPercent => elementsToHandleCount == 0 
        ? 0 
        : done * 100 / elementsToHandleCount;

    public void AddNewElement()
    {
        elementsToHandleCount++;
        Reset();
    }
    public void RemoveElement()
    {
        elementsToHandleCount--;
        Reset();
    }
    public void ElementDone()
    {
        done++;
    }
    public void Reset()
    {
        done = 0;
    }
}
