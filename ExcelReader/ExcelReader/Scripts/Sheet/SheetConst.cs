using System;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class SheetConst : ISheetData
{
    public const string Suffix = "Const";

    public readonly ISheet SheetData;
    public string ExcelName { get; }
    public string SheetName { get; }
    public string FieldName { get; }

    public SheetConst(ISheet sheetData, string excelExcel)
    {
        var sheetName = sheetData.SheetName;
        if (sheetName.Contains('@') || sheetName.Contains('&'))
        {
            throw new Exception($"Excel[{excelExcel}] Sheet[{sheetName}]: Const sheet name cannot contain '@' or '&'.");
        }

        SheetData = sheetData;
        ExcelName = excelExcel;
        SheetName = sheetName;
        FieldName = $"{excelExcel}{sheetName.ClearRemark()}{Suffix}";
    }
}