using System;
using NPOI.SS.UserModel;

namespace Config
{
    public interface ISheetData
    {
        public string ExcelName { get; }

        public string SheetName { get; }

        public string FieldName { get; }
    }

    public class SheetEnum : ISheetData
    {
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
            FieldName = sheetName;
        }
    }

    public class SheetConst : ISheetData
    {
        public readonly ISheet SheetData;
        public string ExcelName { get; }
        public string SheetName { get; }
        public string FieldName { get; }

        public SheetConst(ISheet sheetData, string excelName)
        {
            var sheetName = sheetData.SheetName;
            if (sheetName.Contains('@') || sheetName.Contains('&'))
            {
                throw new Exception($"Excel[{excelName}] Sheet[{sheetName}]: Const sheet name cannot contain '@' or '&'.");
            }

            SheetData = sheetData;
            ExcelName = excelName;
            SheetName = sheetName;
            FieldName = sheetName;
        }
    }

    public class SheetValue : ISheetData
    {
        public readonly ISheet SheetData;
        public readonly string ScriptName;
        public string ExcelName { get; }
        public string SheetName { get; }
        public string FieldName { get; }

        public SheetValue(ISheet sheetData, string excelName)
        {
            var sheetName = sheetData.SheetName;

            SheetData = sheetData;
            ExcelName = excelName;
            SheetName = sheetName;

            if (sheetName.Contains('&'))
            {
                var renameArray = sheetName.Split('&');
                FieldName = renameArray[1];
                ScriptName = renameArray[0];
            }
            else
            {
                FieldName = sheetName;
                ScriptName = sheetName;
            }
        }
    }

    public class SheetMerge : ISheetData
    {
        public readonly ISheet SheetData;
        public readonly string MergeExcel;
        public readonly string ScriptName;
        public string ExcelName { get; }
        public string SheetName { get; }
        public string FieldName { get; }

        public SheetMerge(ISheet sheetData, string excelName)
        {
            var sheetName = sheetData.SheetName;

            SheetData = sheetData;
            ExcelName = excelName;
            SheetName = sheetName;

            var indexMerge = sheetName?.IndexOf('@') ?? -1;
            var indexRename = sheetName?.IndexOf('&') ?? -1;

            if (indexRename != -1)
            {
                if (indexMerge > indexRename)
                {
                    sheetName = sheetName.Split('&')[0];
                }
                else
                {
                    var renameArray = sheetName.Split('&');
                    sheetName = $"{renameArray[0]}@{renameArray[1].Split('@')[1]}";
                }
            }

            var mergeArray = sheetName.Split('@');
            var mergeString = mergeArray[1];
            ScriptName = mergeArray[0];

            if (!mergeString.Contains('.'))
            {
                FieldName = mergeString;
                MergeExcel = excelName;
            }
            else
            {
                var pointArray = mergeString.Split('.');
                FieldName = pointArray[1];
                MergeExcel = pointArray[0];
            }
        }
    }
}