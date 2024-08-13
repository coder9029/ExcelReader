using System;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class SheetMerge : ISheetData
{
    public readonly ISheet SheetData;
    public readonly string MergeExcel;
    public readonly string TableName;
    public string ExcelName { get; }
    public string SheetName { get; }
    public string FieldName { get; }

    public SheetMerge(ISheet sheetData, string excelExcel)
    {
        var sheetName = sheetData.SheetName;

        SheetData = sheetData;
        ExcelName = excelExcel;
        SheetName = sheetName;
        
        if (sheetName.Contains('&'))
        {
            throw new Exception($"Excel[{excelExcel}] Sheet[{sheetName}]: Const sheet name cannot contain '&'.");
        }

        var mergeArray = sheetName.ClearRemark().Split('@');
        var mergeString = mergeArray[1];
        TableName = mergeArray[0];

        if (!mergeString.Contains('.'))
        {
            FieldName = mergeString;
            MergeExcel = excelExcel;
        }
        else
        {
            var pointArray = mergeString.Split('.');
            FieldName = pointArray[1];
            MergeExcel = pointArray[0];
        }
    }
}