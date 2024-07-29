using System;
using System.Collections;
using System.Collections.Generic;
using NPOI.SS.Formula.Functions;
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

    public class SheetCollector : IEnumerable, IEnumerable<>
    {
        private readonly ESheetType sheetType;
        private readonly int columnIndex;
        private readonly object collector;

        public SheetCollector(Dictionary<int, SheetHeader> header, int index)
        {
            if (!header.TryGetValue(index, out var headerData))
            {
                throw new Exception("This index data is not included");
            }
            
            sheetType = headerData.ESheetType;
            columnIndex = index;

            switch (type)
            {
                case ESheetType.List:
                {
                    collector = new List<IRow>();
                    break;
                }

                case ESheetType.Dict:
                {
                    
                }
            }

            foreach (var VARIABLE in this)
            {
                
            }
        }
        
        public IEnumerator GetEnumerator()
        {
            if (sheetType == ESheetType.List && collector is List<IRow> list)
            {
                return list;
            }
            
            if (sheetType == ESheetType.Dict)
            {
                var key = rowData.GetCell(columnIndex).ToString();
                if (string.IsNullOrEmpty(key))
                {
                    throw new Exception();
                }
                
                if (collector is Dictionary<string, IRow> dictRow)
                {
                    return dictRow;
                }
                
                if (collector is SheetCollector dictCollector)
                {
                    if (dictCollector.TryAdd(rowData))
                    {
                        return true;
                    }
                    
                    throw new Exception("");
                }
            }
        }
        
        public bool TryAdd(IRow rowData)
        {
            if (sheetType == ESheetType.List && collector is List<IRow> list)
            {
                list.Add(rowData);
                return true;
            }

            if (sheetType == ESheetType.Dict)
            {
                var key = rowData.GetCell(columnIndex).ToString();
                if (string.IsNullOrEmpty(key))
                {
                    throw new Exception();
                }
                
                if (collector is Dictionary<string, IRow> dictRow)
                {
                    if (dictRow.TryAdd(key, rowData))
                    {
                        return true;
                    }
                    
                    throw new Exception("");
                }
                
                if (collector is SheetCollector dictCollector)
                {
                    if (dictCollector.TryAdd(rowData))
                    {
                        return true;
                    }
                    
                    throw new Exception("");
                }
            }

            return false;
        }

        public void Add(string key,  IRow value)
        {
            if (collector is SheetDict dict)
            {
                dict.Add(key, value);
            }

            if (collector is SheetList list)
            {
                list.Add(value);
            }
        }
    }

    public struct SheetHeader
    {
        public string FieldName;
        public string FieldType;
        public ESheetType ESheetType;
    }

    public enum ESheetType
    {
        None,
        List,
        Dict,
    }

    public class SheetDict(ESheetType type, int index) : Dictionary<string, object>
    {
        public ESheetType RowType = type;
        public int RowIndex = index;

        public void Add(IRow rowData)
        {
            if (RowType == ESheetType.List)
            {
                base.Add();
            }
        }
    }

    public class SheetList : List<object>
    {
        
    }
}