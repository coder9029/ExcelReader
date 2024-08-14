using System;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class SheetEnum : ISheetData
{
    public const string Prefix = "E";

    public readonly ISheet SheetData;
    public string ExcelName { get; }
    public string SheetName { get; }
    public string FieldName { get; }

    public SheetEnum(ISheet sheetData, string excelName)
    {
        var sheetName = sheetData.SheetName;
        if (sheetName.Contains('@') || sheetName.Contains('&'))
        {
            throw new Exception($"Excel[{excelName}] Sheet[{sheetName}]: Enum sheet name cannot contain '@' or '&'.");
        }

        SheetData = sheetData;
        ExcelName = excelName;
        SheetName = sheetName;
        FieldName = $"{Prefix}{excelName}{sheetName.ClearRemark()}";
    }
}