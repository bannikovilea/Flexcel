namespace Flexcel.Internal;

public class ColumnSettings
{
    public int? Width { get; private set; }
    
    public bool? AutoFit { get; private set; }
    
    public string? CellFormat { get; set; }

    public void SetWidth(int minWidth)
    {
        if (minWidth < 0)
            throw new ArgumentException("Width should be grater than zero");
        
        Width = minWidth;
    }

    public void SetAutoFit(bool autoFit)
    {
        AutoFit = autoFit;
    }

    public void SetFormat(string format)
    {
        CellFormat = format;
    }
}