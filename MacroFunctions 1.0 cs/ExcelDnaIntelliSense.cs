using ExcelDna.Integration;
using ExcelDna.IntelliSense;

public class ExcelDnaIntelliSense : IExcelAddIn
{
    public void AutoOpen() => IntelliSenseServer.Install();
    public void AutoClose() => IntelliSenseServer.Uninstall();
}