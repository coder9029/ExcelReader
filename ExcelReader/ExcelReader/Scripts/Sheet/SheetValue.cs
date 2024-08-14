using System;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class SheetValue : ISheetData
{
    public const string Suffix = "Table";
    public readonly ISheet SheetData;
    public readonly string TableName;
    public string ExcelName { get; }
    public string SheetName { get; }
    public string FieldName { get; }

    public SheetValue(ISheet sheetData, string excelExcel)
    {
        var sheetName = sheetData.SheetName;

        SheetData = sheetData;
        ExcelName = excelExcel;
        SheetName = sheetName;

        sheetName = sheetName.ClearRemark();
        if (sheetName.Contains('&'))
        {
            var renameArray = sheetName.Split('&');
            FieldName = $"{excelExcel}{renameArray[1]}{Suffix}";
            TableName = renameArray[0];
        }
        else
        {
            FieldName = $"{excelExcel}{sheetName}{Suffix}";
            TableName = sheetName;
        }
    }
}