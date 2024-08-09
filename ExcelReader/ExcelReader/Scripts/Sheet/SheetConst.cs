using System;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class SheetConst : ISheetData
{
    public readonly ISheet SheetData;
    public readonly string TableName;
    public string ExcelExcel { get; }
    public string SheetName { get; }

    public SheetConst(ISheet sheetData, string excelExcel)
    {
        var sheetName = sheetData.SheetName;
        if (sheetName.Contains('@') || sheetName.Contains('&'))
        {
            throw new Exception($"Excel[{excelExcel}] Sheet[{sheetName}]: Const sheet name cannot contain '@' or '&'.");
        }

        SheetData = sheetData;
        ExcelExcel = excelExcel;
        SheetName = sheetName;
        TableName = sheetName.ClearRemark();
    }
}