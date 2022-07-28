namespace Flexcel.Tests;

public class TestExportClass
{
    public static int StaticConstant = 40;
    public int Constant = 20;
    public int LambdaField => Constant;

    public string SomeString { get; set; }
    
    public bool SomeBool { get; set; }
    
    public DateTime DateTime { get; set; }
    
    public DateTimeOffset DateTimeOffset { get; set; }
}