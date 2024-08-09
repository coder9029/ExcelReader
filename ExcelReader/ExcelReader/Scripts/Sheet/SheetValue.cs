using System;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class SheetValue : ISheetData
{
    public readonly ISheet SheetData;
    public readonly string TableName;
    public readonly string FieldName;
    public string ExcelExcel { get; }
    public string SheetName { get; }

    public SheetValue(ISheet sheetData, string excelExcel)
    {
        var sheetName = sheetData.SheetName;

        SheetData = sheetData;
        ExcelExcel = excelExcel;
        SheetName = sheetName;

        sheetName = sheetName.ClearRemark();
        if (sheetName.Contains('&'))
        {
            var renameArray = sheetName.Split('&');
            FieldName = renameArray[1];
            TableName = renameArray[0];
        }
        else
        {
            FieldName = sheetName;
            TableName = sheetName;
        }
    }
}