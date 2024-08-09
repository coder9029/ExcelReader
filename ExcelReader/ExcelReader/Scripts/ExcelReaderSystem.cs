using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelReader
{
    public static class ExcelReaderSystem
    {
        private const string NameSpace = "Config";

        private const string SheetSuffix = "Table";

        private const string ClassName = "TableComponent";

        private static readonly Dictionary<string, List<object>> _excelCollector = new();

        public static readonly StringBuilder StringBuilder = new();

        private static readonly Dictionary<int, (string, string)> _tempSheetColumn = new();

        public static void Program(string dirPath, string outPath, bool isFormat = false)
        {
            if (string.IsNullOrEmpty(dirPath) || !Directory.Exists(dirPath))
            {
                throw new Exception($"DirPath[{dirPath}]: The address is invalid.");
            }

            if (string.IsNullOrEmpty(outPath))
            {
                throw new Exception($"OutPath[{outPath}]: The address is invalid.");
            }

            var isExists = Directory.Exists(outPath);
            if (isExists && isFormat)
            {
                Directory.Delete(outPath, true);
            }

            if (!isExists || isFormat)
            {
                Directory.CreateDirectory(outPath);
            }

            foreach (var file in Directory.GetFiles(dirPath, string.Empty, SearchOption.AllDirectories))
            {
                var fileName = Path.GetFileNameWithoutExtension(file);
                if (fileName.StartsWith('~') || fileName.StartsWith('#'))
                {
                    continue;
                }

                var fileExt = Path.GetExtension(file).ToLower();
                if (fileExt != ".xls" && fileExt != ".xlsx")
                {
                    continue;
                }

                CollectExcel(file);
                GenerateScript(fileName, outPath);
            }

            Console.WriteLine("Excel data is generated successfully.");
        }

        private static void CollectExcel(string file)
        {
            using var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            IWorkbook workbook = null;
            var fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xls")
            {
                workbook = new HSSFWorkbook(fileStream);
            }

            if (fileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook(fileStream);
            }

            if (workbook == null)
            {
                throw new Exception($"File[{file}]: Only support xls/xlsx files.");
            }

            var fileName = Path.GetFileNameWithoutExtension(file);

            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                try
                {
                    if (sheet == null || sheet.SheetName.StartsWith('#'))
                    {
                        continue;
                    }

                    var sheetData = sheet.GetSheetParam(fileName);
                    CollectEnum(sheetData);
                    CollectConst(sheetData);
                    CollectValue(sheetData);
                    CollectMerge(sheetData);
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception);
                    throw;
                }
            }
        }

        private static void CollectEnum(ISheetData sheetData)
        {
            if (sheetData is not SheetEnum sheetEnum)
            {
                return;
            }

            if (!_excelCollector.TryGetValue(sheetEnum.ExcelExcel, out var sheetItems))
            {
                sheetItems = [];
                _excelCollector[sheetEnum.ExcelExcel] = sheetItems;
            }

            foreach (var objects in sheetItems)
            {
                if (objects is ISheetData tempData)
                {
                    if (tempData.FieldName != sheetEnum.FieldName)
                    {
                        continue;
                    }

                    ThrowException(tempData, sheetEnum);
                }

                if (objects is List<ISheetData> tempList)
                {
                    var tempFirst = tempList.First();
                    if (tempFirst.FieldName != sheetEnum.FieldName)
                    {
                        continue;
                    }

                    ThrowException(tempFirst, sheetEnum);
                }
            }

            sheetItems.Add(sheetEnum);
        }

        private static void CollectConst(ISheetData sheetData)
        {
            if (sheetData is not SheetConst sheetConst)
            {
                return;
            }

            if (!_excelCollector.TryGetValue(sheetConst.ExcelExcel, out var sheetItems))
            {
                sheetItems = [];
                _excelCollector[sheetConst.ExcelExcel] = sheetItems;
            }

            foreach (var objects in sheetItems)
            {
                if (objects is ISheetData tempData)
                {
                    if (tempData.FieldName != sheetConst.TableName)
                    {
                        continue;
                    }

                    ThrowException(tempData, sheetConst);
                }

                if (objects is List<ISheetData> tempList)
                {
                    var tempFirst = tempList.First();
                    if (tempFirst.FieldName != sheetConst.TableName)
                    {
                        continue;
                    }

                    ThrowException(tempFirst, sheetConst);
                }
            }

            sheetItems.Add(sheetConst);
        }

        private static void CollectValue(ISheetData sheetData)
        {
            if (sheetData is not SheetValue sheetValue)
            {
                return;
            }

            if (!_excelCollector.TryGetValue(sheetValue.ExcelExcel, out var sheetItems))
            {
                sheetItems = [];
                _excelCollector[sheetValue.ExcelExcel] = sheetItems;
            }

            List<ISheetData> sheetList = null;

            foreach (var objects in sheetItems)
            {
                if (objects is ISheetData tempData)
                {
                    if (tempData.FieldName != sheetValue.FieldName)
                    {
                        continue;
                    }

                    ThrowException(tempData, sheetValue);
                }

                if (objects is List<ISheetData> tempList)
                {
                    var tempFirst = tempList.First();
                    if (tempFirst.FieldName != sheetValue.FieldName)
                    {
                        continue;
                    }

                    sheetList = tempList;

                    if (tempFirst is SheetValue tempValue && tempValue.TableName == sheetValue.TableName)
                    {
                        continue;
                    }

                    if (tempFirst is SheetMerge tempMerge && tempMerge.TableName == sheetValue.TableName)
                    {
                        continue;
                    }

                    ThrowException(tempFirst, sheetValue);
                }
            }

            if (sheetList != null)
            {
                sheetList.Add(sheetValue);
            }
            else
            {
                sheetList = [sheetValue];
                sheetItems.Add(sheetList);
            }
        }

        private static void CollectMerge(ISheetData sheetData)
        {
            if (sheetData is not SheetMerge sheetMerge)
            {
                return;
            }

            if (!_excelCollector.TryGetValue(sheetMerge.MergeExcel, out var sheetItems))
            {
                sheetItems = [];
                _excelCollector[sheetMerge.MergeExcel] = sheetItems;
            }

            List<ISheetData> sheetList = null;

            foreach (var objects in sheetItems)
            {
                if (objects is ISheetData tempData)
                {
                    if (tempData.FieldName != sheetMerge.FieldName)
                    {
                        continue;
                    }

                    ThrowException(tempData, sheetMerge);
                }

                if (objects is List<ISheetData> tempList)
                {
                    var tempFirst = tempList.First();
                    if (tempFirst.FieldName != sheetMerge.FieldName)
                    {
                        continue;
                    }

                    sheetList = tempList;

                    if (tempFirst is SheetValue tempValue && tempValue.TableName == sheetMerge.TableName)
                    {
                        continue;
                    }

                    if (tempFirst is SheetMerge tempMerge && tempMerge.TableName == sheetMerge.TableName)
                    {
                        continue;
                    }

                    ThrowException(tempFirst, sheetMerge);
                }
            }

            if (sheetList != null)
            {
                sheetList.Add(sheetMerge);
            }
            else
            {
                sheetList = [sheetMerge];
                sheetItems.Add(sheetList);
            }
        }

        private static void CollectMerge1(ISheetData sheetData)
        {
            if (sheetData is not SheetMerge mergeData)
            {
                return;
            }

            if (_excelCollector.TryGetValue(mergeData.ExcelExcel, out var sheetItem))
            {
                foreach (var (fieldName, sheetType, sheetList) in sheetItem)
                {
                    if (fieldName != mergeData.MergeField)
                    {
                        continue;
                    }

                    if (sheetType == ESheetCollector.Value)
                    {
                        continue;
                    }

                    var recordSheetData = sheetList.First();
                    var recordExcelName = recordSheetData.ExcelName;
                    var recordSheetName = recordSheetData.SheetName;

                    var targetExcelName = mergeData.ExcelExcel;
                    var targetSheetName = mergeData.SheetName;

                    var sheetA = $"ExcelA[{recordExcelName}] SheetA[{recordSheetName}] ;";
                    var sheetB = $"ExcelB[{targetExcelName}] SheetB[{targetSheetName}] ;";
                    throw new Exception($"{sheetA} {sheetB} : A and B conflict.");
                }
            }
            else
            {
                sheetItem = [];
                _excelCollector[mergeData.ExcelExcel] = sheetItem;
            }

            sheetItem.Add((mergeData.FieldName, ESheetCollector.Value, [mergeData]));


            // var classParam = sheet.GetSheetParam();
            // if (_valueCollector.ContainsKey(classParam.ScriptName))
            // {
            //     return;
            // }

            // if (sheetData.LastRowNum < 4)
            // {
            //     return;
            // }
            //
            // _tempStrBuilder.Clear();
            // var rowField = sheetData.GetRow(0);
            // var rowType = sheetData.GetRow(1);
            //
            // // if (classParam.SheetType is SheetType.List)
            // {
            //     _tempSheetColumn.Clear();
            //     for (var i = 0; i < rowField.LastCellNum; i++)
            //     {
            //         var field = rowField.GetCell(i)?.ToString();
            //         if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
            //         {
            //             continue;
            //         }
            //
            //         var type = rowType.GetCell(i)?.ToString();
            //         if (string.IsNullOrEmpty(type) || !type.GetFieldType(out _))
            //         {
            //             throw new Exception($"Sheet[{sheetData.SheetName}] - Row[2] - Col[{i + 1}] - Value[{type}]: The value is empty or invalid.");
            //         }
            //
            //         _tempSheetColumn.Add(i, (field, type));
            //     }
            //
            //     if (_tempSheetColumn.Count == 0)
            //     {
            //         return;
            //     }
            //
            //     for (var i = 3; i <= sheetData.LastRowNum; i++)
            //     {
            //         var rowTemp = sheetData.GetRow(i);
            //
            //         var keyValueTemp = rowTemp.GetCell(0)?.ToString();
            //         if (string.IsNullOrEmpty(keyValueTemp) || keyValueTemp.StartsWith('#'))
            //         {
            //             continue;
            //         }
            //
            //         var contentString = _tempStrBuilder.ToString();
            //         _tempStrBuilder.Clear();
            //
            //         for (var j = 0; j < rowTemp.LastCellNum; j++)
            //         {
            //             if (!_tempSheetColumn.TryGetValue(j, out (string field, string type) colTemp))
            //             {
            //                 continue;
            //             }
            //
            //             var value = rowTemp.GetCell(j)?.ToString();
            //             if (string.IsNullOrEmpty(value))
            //             {
            //                 continue;
            //             }
            //
            //             var realValue = value.GetFiledValue(colTemp.type);
            //             _tempStrBuilder.AppendLineWithTab($"{colTemp.field} = {realValue},", 4);
            //         }
            //
            //         var valueString = _tempStrBuilder.ToString();
            //         _tempStrBuilder.Clear();
            //         _tempStrBuilder.Append(contentString);
            //
            //         if (string.IsNullOrEmpty(valueString))
            //         {
            //             continue;
            //         }
            //
            //         // _tempStrBuilder.AppendLineWithTab($"new {classParam.ScriptName}", 3);
            //         _tempStrBuilder.AppendLineWithTab("{", 3);
            //         _tempStrBuilder.Append(valueString);
            //         _tempStrBuilder.AppendLineWithTab("},", 3);
            //     }
            //
            //     if (_tempStrBuilder.Length == 0)
            //     {
            //         return;
            //     }
            //
            //     var listString = _tempStrBuilder.ToString();
            //     _tempStrBuilder.Clear();
            //     // _tempStrBuilder.AppendLineWithTab($"public static IReadOnlyList<{classParam.ScriptName}> {sheet.SheetName} = new List<{classParam.ScriptName}>()", 2);
            //     _tempStrBuilder.AppendLineWithTab("{", 2);
            //     _tempStrBuilder.Append(listString);
            //     _tempStrBuilder.AppendLineWithTab("};", 2);
            // }
            //
            // // if (classParam.SheetType is SheetType.Dict)
            // {
            //     var keyType = rowType.GetCell(0)?.ToString();
            //     if (string.IsNullOrEmpty(keyType) || !keyType.GetFieldType(out var keyTypeString))
            //     {
            //         throw new Exception($"Sheet[{sheetData.SheetName}] - Row[{2}] - Col[{1}] - Value[{keyType}]: The value is empty or invalid.");
            //     }
            //
            //     if (!keyType.IsDictKeyValid())
            //     {
            //         throw new Exception($"Sheet[{sheetData.SheetName}] - Row[{2}] - Col[{1}] - Value[{keyType}]: The dictionary key is invalid.");
            //     }
            //
            //     _tempSheetColumn.Clear();
            //     for (var i = 0; i < rowField.LastCellNum; i++)
            //     {
            //         var field = rowField.GetCell(i)?.ToString();
            //         if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
            //         {
            //             continue;
            //         }
            //
            //         var type = rowType.GetCell(i)?.ToString();
            //         if (string.IsNullOrEmpty(type) || !type.GetFieldType(out _))
            //         {
            //             throw new Exception($"Sheet[{sheetData.SheetName}] - Row[2] - Col[{i + 1}] - Value[{type}]: The value is empty or invalid.");
            //         }
            //
            //         _tempSheetColumn.Add(i, (field, type));
            //     }
            //
            //     if (_tempSheetColumn.Count == 0)
            //     {
            //         return;
            //     }
            //
            //     for (var i = 3; i <= sheetData.LastRowNum; i++)
            //     {
            //         var rowTemp = sheetData.GetRow(i);
            //
            //         var keyValueTemp = rowTemp.GetCell(0)?.ToString();
            //         if (string.IsNullOrEmpty(keyValueTemp) || keyValueTemp.StartsWith('#'))
            //         {
            //             continue;
            //         }
            //
            //         var contentString = _tempStrBuilder.ToString();
            //         _tempStrBuilder.Clear();
            //
            //         for (var j = 0; j < rowTemp.LastCellNum; j++)
            //         {
            //             if (!_tempSheetColumn.TryGetValue(j, out (string field, string type) colTemp))
            //             {
            //                 continue;
            //             }
            //
            //             var value = rowTemp.GetCell(j)?.ToString();
            //             if (string.IsNullOrEmpty(value))
            //             {
            //                 continue;
            //             }
            //
            //             var realValue = value.GetFiledValue(colTemp.type);
            //             _tempStrBuilder.AppendLineWithTab($"{colTemp.field} = {realValue},", 7);
            //         }
            //
            //         var valueString = _tempStrBuilder.ToString();
            //         _tempStrBuilder.Clear();
            //         _tempStrBuilder.Append(contentString);
            //
            //         if (string.IsNullOrEmpty(valueString))
            //         {
            //             continue;
            //         }
            //
            //         var keyValueString = keyValueTemp.GetFiledValue(keyType);
            //         _tempStrBuilder.AppendLineWithTab($"case {keyValueString}:", 4);
            //         _tempStrBuilder.AppendLineWithTab("{", 4);
            //         _tempStrBuilder.AppendLineWithTab($"if (!_m{sheetData.SheetName}.ContainsKey({keyValueString}))", 5);
            //         _tempStrBuilder.AppendLineWithTab("{", 5);
            //         // _tempStrBuilder.AppendLineWithTab($"_m{sheet.SheetName}[{keyValueString}] = new {classParam.ScriptName}", 6);
            //         _tempStrBuilder.AppendLineWithTab("{", 6);
            //         _tempStrBuilder.Append(valueString);
            //         _tempStrBuilder.AppendLineWithTab("};", 6);
            //         _tempStrBuilder.AppendLineWithTab("}", 5);
            //         _tempStrBuilder.AppendLine();
            //         _tempStrBuilder.AppendLineWithTab($"return _m{sheetData.SheetName}[{keyValueString}];", 5);
            //         _tempStrBuilder.AppendLineWithTab("}", 4);
            //     }
            //
            //     if (_tempStrBuilder.Length == 0)
            //     {
            //         return;
            //     }
            //
            //     var dictString = _tempStrBuilder.ToString();
            //     _tempStrBuilder.Clear();
            //     // _tempStrBuilder.AppendLineWithTab($"private static readonly Dictionary<{keyTypeString}, {classParam.ScriptName}> _m{sheet.SheetName} = new();", 2);
            //     // _tempStrBuilder.AppendLineWithTab($"public static {classParam.ScriptName} {sheet.SheetName}({keyTypeString} key)", 2);
            //     _tempStrBuilder.AppendLineWithTab("{", 2);
            //     _tempStrBuilder.AppendLineWithTab("switch (key)", 3);
            //     _tempStrBuilder.AppendLineWithTab("{", 3);
            //     _tempStrBuilder.Append(dictString);
            //     _tempStrBuilder.AppendLineWithTab("}", 3);
            //     _tempStrBuilder.AppendLine();
            //     _tempStrBuilder.AppendLineWithTab("return null;", 3);
            //     _tempStrBuilder.AppendLineWithTab("}", 2);
            // }

            // _valueCollector.Add(classParam.ScriptName, _tempStrBuilder.ToString());
        }

        private static void GenerateScript(string path)
        {
            foreach (var sheetItems in _excelCollector.Values)
            {
                foreach (var sheetData in sheetItems)
                {
                    var content = GenerateSheet(sheetData, out var fileName);
                    if (string.IsNullOrEmpty(content) || string.IsNullOrEmpty(fileName))
                    {
                        continue;
                    }

                    StringBuilder.AppendLine("//This file is automatically generated, please do not modify it manually");
                    StringBuilder.AppendLine();
                    StringBuilder.AppendLine("using System.Collections.Generic;");
                    StringBuilder.AppendLine();
                    StringBuilder.AppendLine($"namespace {NameSpace}");
                    StringBuilder.AppendLine("{");
                    StringBuilder.Append(content);
                    StringBuilder.AppendLine("}");

                    File.WriteAllText(Path.Combine(path, $"{fileName}.cs"), StringBuilder.ToString());
                }
            }
        }

        private static string GenerateSheet(object sheetObject, out string fileName)
        {
            if (sheetObject is SheetEnum sheetEnum)
            {
                return GenerateEnum(sheetEnum, out fileName);
            }

            if (sheetObject is SheetConst sheetConst)
            {
                return  GenerateConst(sheetConst, out fileName);
            }

            if (sheetObject is List<ISheetData> sheetList)
            {
                return GenerateList(sheetList, out fileName);
            }

            fileName = string.Empty;
            return string.Empty;
        }

        private static string GenerateEnum(SheetEnum sheetEnum, out string fileName)
        {
            fileName = $"{sheetEnum.ExcelExcel}{sheetEnum.FieldName}{SheetSuffix}";

            StringBuilder.Clear();

            const int fieldColumn = 0;
            const int summaryColumn = 1;
            const int valueColumn = 2;

            var sheetData = sheetEnum.SheetData;
            var isFirstWrite = false;
            for (var i = 1; i <= sheetData.LastRowNum; i++)
            {
                var rowItem = sheetData.GetRow(i);

                var field = rowItem?.GetCell(fieldColumn)?.ToString();
                if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
                {
                    continue;
                }

                if (!isFirstWrite)
                {
                    isFirstWrite = true;
                }
                else
                {
                    StringBuilder.AppendLine();
                }

                var summary = rowItem.GetCell(summaryColumn)?.ToString();
                if (!string.IsNullOrEmpty(summary))
                {
                    StringBuilder.AppendTab($"/// <summary>{summary}</summary>", 2);
                }

                var value = rowItem.GetCell(valueColumn)?.ToString();
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception($"Excel[{sheetEnum.ExcelExcel}] Sheet[{sheetEnum.SheetName}] - Row[{i + 1}] - Col[{valueColumn + 1}] - Value[{value}]: The value is empty.");
                }

                StringBuilder.AppendTab($"{field} = {value},", 2);
            }

            if (StringBuilder.Length == 0)
            {
                return string.Empty;
            }

            var contentString = StringBuilder.ToString();
            StringBuilder.Clear();

            StringBuilder.AppendTab($"public class {fileName}");
            StringBuilder.AppendTab("{");
            StringBuilder.Append(contentString);
            StringBuilder.AppendTab("}");
            return StringBuilder.ToString();
        }

        private static string GenerateConst(SheetConst sheetConst, out string fileName)
        {
            fileName = $"{sheetConst.ExcelExcel}{sheetConst.TableName}{SheetSuffix}";

            StringBuilder.Clear();

            const int fieldColumn = 0;
            const int typeColumn = 1;
            const int summaryColumn = 2;
            const int valueColumn = 3;

            var sheetData = sheetConst.SheetData;
            var isFirstWrite = false;
            for (var i = 1; i <= sheetData.LastRowNum; i++)
            {
                var rowItem = sheetData.GetRow(i);

                var field = rowItem?.GetCell(fieldColumn)?.ToString();
                if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
                {
                    continue;
                }

                if (!isFirstWrite)
                {
                    isFirstWrite = true;
                }
                else
                {
                    StringBuilder.AppendLine();
                }

                var typeString = rowItem.GetCell(typeColumn)?.ToString();
                if (string.IsNullOrEmpty(typeString) || !typeString.GetFieldType(out var type))
                {
                    throw new Exception($"Excel[{sheetConst.ExcelExcel}] Sheet[{sheetConst.SheetName}] - Row[{i + 1}] - Col[{typeColumn + 1}] - Value[{typeString}]: The value is empty or invalid.");
                }

                var summary = rowItem.GetCell(summaryColumn)?.ToString();
                if (!string.IsNullOrEmpty(summary))
                {
                    StringBuilder.AppendTab($"/// <summary>{summary}</summary>", 2);
                }

                var valueString = rowItem.GetCell(valueColumn)?.ToString();
                if (string.IsNullOrEmpty(valueString) || !valueString.GetFiledValue(typeString, out var value))
                {
                    throw new Exception($"Excel[{sheetConst.ExcelExcel}] Sheet[{sheetConst.SheetName}] - Row[{i + 1}] - Col[{valueColumn + 1}] - Value[{valueString}]: The value is empty or invalid.");
                }

                StringBuilder.AppendTab($"public {type} {field} = {value};", 2);
            }

            if (StringBuilder.Length == 0)
            {
                return string.Empty;
            }

            var contentString = StringBuilder.ToString();
            StringBuilder.Clear();

            StringBuilder.AppendTab($"public class {fileName}");
            StringBuilder.AppendTab("{");
            StringBuilder.Append(contentString);
            StringBuilder.AppendTab("}");
            return StringBuilder.ToString();
        }

        private static string GenerateList(List<ISheetData> sheetList, out string fileName)
        {
            fileName = $"{sheetList.ExcelName}{sheetList.FieldName}{SheetSuffix}";

            StringBuilder.Clear();

            const int fieldColumn = 0;
            const int typeColumn = 1;
            const int summaryColumn = 2;
            const int valueColumn = 3;

            var sheetData = sheetList.SheetData;
            var isFirstWrite = false;
            for (var i = 1; i <= sheetData.LastRowNum; i++)
            {
                var rowItem = sheetData.GetRow(i);

                var field = rowItem?.GetCell(fieldColumn)?.ToString();
                if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
                {
                    continue;
                }

                if (!isFirstWrite)
                {
                    isFirstWrite = true;
                }
                else
                {
                    StringBuilder.AppendLine();
                }

                var typeString = rowItem.GetCell(typeColumn)?.ToString();
                if (string.IsNullOrEmpty(typeString) || !typeString.GetFieldType(out var type))
                {
                    throw new Exception($"Excel[{sheetList.ExcelName}] Sheet[{sheetList.SheetName}] - Row[{i + 1}] - Col[{typeColumn + 1}] - Value[{typeString}]: The value is empty or invalid.");
                }

                var summary = rowItem.GetCell(summaryColumn)?.ToString();
                if (!string.IsNullOrEmpty(summary))
                {
                    StringBuilder.AppendTab($"/// <summary>{summary}</summary>", 2);
                }

                var valueString = rowItem.GetCell(valueColumn)?.ToString();
                if (string.IsNullOrEmpty(valueString) || !valueString.GetFiledValue(typeString, out var value))
                {
                    throw new Exception($"Excel[{sheetList.ExcelName}] Sheet[{sheetList.SheetName}] - Row[{i + 1}] - Col[{valueColumn + 1}] - Value[{valueString}]: The value is empty or invalid.");
                }

                StringBuilder.AppendTab($"public {type} {field} = {value};", 2);
            }

            if (StringBuilder.Length == 0)
            {
                return string.Empty;
            }

            var contentString = StringBuilder.ToString();
            StringBuilder.Clear();

            StringBuilder.AppendTab($"public class {fileName}");
            StringBuilder.AppendTab("{");
            StringBuilder.Append(contentString);
            StringBuilder.AppendTab("}");
            return StringBuilder.ToString();
        }

        private static void ThrowException(ISheetData sheetDataA, ISheetData sheetDataB)
        {
            var aExcelName = sheetDataA.ExcelExcel;
            var aSheetName = sheetDataA.SheetName;

            var bExcelName = sheetDataB.ExcelExcel;
            var bSheetName = sheetDataB.SheetName;

            var sheetA = $"ExcelA[{aExcelName}] SheetA[{aSheetName}] ;";
            var sheetB = $"ExcelB[{bExcelName}] SheetB[{bSheetName}] ;";
            throw new Exception($"{sheetA} {sheetB} : A and B conflict.");
        }
    }
}