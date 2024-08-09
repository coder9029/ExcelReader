using System;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class SheetMerge : ISheetData
{
    public readonly ISheet SheetData;
    public readonly string MergeExcel;
    public readonly string MergeField;
    public readonly string TableName;
    public string ExcelExcel { get; }
    public string SheetName { get; }

    public SheetMerge(ISheet sheetData, string excelExcel)
    {
        var sheetName = sheetData.SheetName;

        SheetData = sheetData;
        ExcelExcel = excelExcel;
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
            MergeField = mergeString;
            MergeExcel = excelExcel;
        }
        else
        {
            var pointArray = mergeString.Split('.');
            MergeField = pointArray[1];
            MergeExcel = pointArray[0];
        }
    }
}