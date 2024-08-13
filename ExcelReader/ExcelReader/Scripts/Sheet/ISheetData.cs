namespace ExcelReader;

public interface ISheetData
{
    public string ExcelName { get; }

    public string SheetName { get; }

    public string FieldName { get; }
}