﻿using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Config
{
    public static class ExcelReaderSystem
    {
        private const string NameSpace = "Config";

        private const string SheetSuffix = "Table";

        private const string ClassName = "TableComponent";

        private static readonly Dictionary<string, List<object>> _excelCollector = new();

        private static readonly StringBuilder _tempStrBuilder = new();

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

            if (!_excelCollector.TryGetValue(sheetEnum.ExcelName, out var sheetItems))
            {
                sheetItems = [];
                _excelCollector[sheetEnum.ExcelName] = sheetItems;
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

            ////////////////////////////////////////////////////
            // if (sheet.LastRowNum < 1)
            // {
            //     throw new Exception($"Sheet[{sheet.SheetName}]: The sheet rows is invalid.");
            // }
            //
            // _tempStrBuilder.Clear();
            //
            // var isFirstWrite = false;
            // for (var i = 1; i <= sheet.LastRowNum; i++)
            // {
            //     var rowItem = sheet.GetRow(i);
            //
            //     var field = rowItem?.GetCell(0)?.ToString();
            //     if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
            //     {
            //         continue;
            //     }
            //
            //     if (!isFirstWrite)
            //     {
            //         isFirstWrite = true;
            //     }
            //     else
            //     {
            //         _tempStrBuilder.AppendLine();
            //     }
            //
            //     var summary = rowItem.GetCell(1)?.ToString();
            //     if (!string.IsNullOrEmpty(summary))
            //     {
            //         _tempStrBuilder.AppendLineWithTab($"/// <summary>{summary}</summary>", 2);
            //     }
            //
            //     _tempStrBuilder.AppendLineWithTab($"{field},", 2);
            // }
            //
            // if (_tempStrBuilder.Length == 0)
            // {
            //     return;
            // }
            //
            // var contentString = _tempStrBuilder.ToString();
            // _tempStrBuilder.Clear();
            // _tempStrBuilder.AppendLineWithTab($"public enum {sheetParam.ScriptName}");
            // _tempStrBuilder.AppendLineWithTab("{");
            // _tempStrBuilder.Append(contentString);
            // _tempStrBuilder.AppendLineWithTab("}");
            //
            // _constCollector.Add(sheetParam.ScriptName, _tempStrBuilder.ToString());
        }

        private static void CollectConst(ISheetData sheetData)
        {
            if (sheetData is not SheetConst sheetConst)
            {
                return;
            }

            if (!_excelCollector.TryGetValue(sheetConst.ExcelName, out var sheetItems))
            {
                sheetItems = [];
                _excelCollector[sheetConst.ExcelName] = sheetItems;
            }

            foreach (var objects in sheetItems)
            {
                if (objects is ISheetData tempData)
                {
                    if (tempData.FieldName != sheetConst.FieldName)
                    {
                        continue;
                    }

                    ThrowException(tempData, sheetConst);
                }

                if (objects is List<ISheetData> tempList)
                {
                    var tempFirst = tempList.First();
                    if (tempFirst.FieldName != sheetConst.FieldName)
                    {
                        continue;
                    }

                    ThrowException(tempFirst, sheetConst);
                }
            }

            sheetItems.Add(sheetConst);


            // if (_constCollector.ContainsKey(sheet.SheetName))
            // {
            //     return;
            // }
            //
            // _tempStrBuilder.Clear();
            //
            // var isFirstWrite = false;
            // for (var i = 1; i <= sheet.LastRowNum; i++)
            // {
            //     var rowItem = sheet.GetRow(i);
            //
            //     var field = rowItem?.GetCell(0)?.ToString();
            //     if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
            //     {
            //         continue;
            //     }
            //
            //     if (!isFirstWrite)
            //     {
            //         isFirstWrite = true;
            //     }
            //     else
            //     {
            //         _tempStrBuilder.AppendLine();
            //     }
            //
            //     var summary = rowItem.GetCell(2)?.ToString();
            //     if (!string.IsNullOrEmpty(summary))
            //     {
            //         _tempStrBuilder.AppendLineWithTab($"/// <summary>{summary}</summary>", 2);
            //     }
            //
            //     var type = rowItem.GetCell(1)?.ToString();
            //     if (string.IsNullOrEmpty(type) || !type.GetFieldType(out var typeString))
            //     {
            //         throw new Exception($"Sheet[{sheet.SheetName}] - Row[{i + 1}] - Col[2] - Value[{type}]: The value is empty or invalid.");
            //     }
            //
            //     var value = rowItem.GetCell(3)?.ToString();
            //     if (string.IsNullOrEmpty(value))
            //     {
            //         throw new Exception($"Sheet[{sheet.SheetName}] - Row[{i + 1}] - Col[3] - Value[{value}]: The value is empty.");
            //     }
            //
            //     var valueString = value.GetFiledValue(type);
            //     _tempStrBuilder.AppendLineWithTab($"public {typeString} {field} = {valueString};", 2);
            // }
            //
            // if (_tempStrBuilder.Length == 0)
            // {
            //     return;
            // }
            //
            // _constCollector.Add(sheet.SheetName, _tempStrBuilder.ToString());
        }

        private static void CollectValue(ISheetData sheetData)
        {
            if (sheetData is not SheetValue sheetValue)
            {
                return;
            }

            if (!_excelCollector.TryGetValue(sheetValue.ExcelName, out var sheetItems))
            {
                sheetItems = [];
                _excelCollector[sheetValue.ExcelName] = sheetItems;
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

                    if (tempFirst is SheetValue tempValue && tempValue.ScriptName == sheetValue.ScriptName)
                    {
                        continue;
                    }

                    if (tempFirst is SheetMerge tempMerge && tempMerge.ScriptName == sheetValue.ScriptName)
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


            ////////////////////////////////////////////////////////

            // if (sheet.LastRowNum < 3)
            // {
            //     throw new Exception($"Sheet[{sheet.SheetName}]: The sheet rows is invalid.");
            // }
            //
            // _tempStrBuilder.Clear();
            // var rowField = sheet.GetRow(0);
            // var rowType = sheet.GetRow(1);
            // var rowSummary = sheet.GetRow(2);

            // var isFirstWrite = false;
            // for (var i = 0; i < rowField.LastCellNum; i++)
            // {
            //     var field = rowField.GetCell(i)?.ToString();
            //     if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
            //     {
            //         continue;
            //     }
            //
            //     if (!isFirstWrite)
            //     {
            //         isFirstWrite = true;
            //     }
            //     else
            //     {
            //         _tempStrBuilder.AppendLine();
            //     }
            //
            //     var summary = rowSummary.GetCell(i)?.ToString();
            //     if (!string.IsNullOrEmpty(summary))
            //     {
            //         _tempStrBuilder.AppendLineWithTab($"/// <summary>{summary}</summary>", 2);
            //     }
            //
            //     var type = rowType.GetCell(i)?.ToString();
            //     if (string.IsNullOrEmpty(type) || !type.GetFieldType(out var typeString))
            //     {
            //         throw new Exception($"Sheet[{sheet.SheetName}] - Row[3] - Col[{i + 1}] - Value[{type}]: The value is empty or invalid.");
            //     }
            //
            //     _tempStrBuilder.AppendLineWithTab($"public {typeString} {field};", 2);
            // }
            //
            // if (_tempStrBuilder.Length == 0)
            // {
            //     return;
            // }
            //
            // var contentString = _tempStrBuilder.ToString();
            // _tempStrBuilder.Clear();
            // _tempStrBuilder.AppendLine("//This file is automatically generated, please do not modify it manually");
            // _tempStrBuilder.AppendLine();
            // _tempStrBuilder.AppendLine("using System.Collections.Generic;");
            // _tempStrBuilder.AppendLine();
            // _tempStrBuilder.AppendLine($"namespace {NameSpace}");
            // _tempStrBuilder.AppendLine("{");
            // _tempStrBuilder.AppendLineWithTab($"public partial class {classParam.ScriptName}");
            // _tempStrBuilder.AppendLineWithTab("{");
            // _tempStrBuilder.Append(contentString);
            // _tempStrBuilder.AppendLineWithTab("}");
            // _tempStrBuilder.AppendLine("}");
            // _classCollector.Add(classParam.ScriptName, _tempStrBuilder.ToString());
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

                    if (tempFirst is SheetValue tempValue && tempValue.ScriptName == sheetMerge.ScriptName)
                    {
                        continue;
                    }

                    if (tempFirst is SheetMerge tempMerge && tempMerge.ScriptName == sheetMerge.ScriptName)
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

            if (_excelCollector.TryGetValue(mergeData.ExcelName, out var sheetItem))
            {
                foreach (var (fieldName, sheetType, sheetList) in sheetItem)
                {
                    if (fieldName != mergeData.MergeField)
                    {
                        continue;
                    }

                    if (sheetType == SheetType.Value)
                    {
                        continue;
                    }

                    var recordSheetData = sheetList.First();
                    var recordExcelName = recordSheetData.ExcelName;
                    var recordSheetName = recordSheetData.SheetName;

                    var targetExcelName = mergeData.ExcelName;
                    var targetSheetName = mergeData.SheetName;

                    var sheetA = $"ExcelA[{recordExcelName}] SheetA[{recordSheetName}] ;";
                    var sheetB = $"ExcelB[{targetExcelName}] SheetB[{targetSheetName}] ;";
                    throw new Exception($"{sheetA} {sheetB} : A and B conflict.");
                }
            }
            else
            {
                sheetItem = [];
                _excelCollector[mergeData.ExcelName] = sheetItem;
            }

            sheetItem.Add((mergeData.FieldName, SheetType.Value, [mergeData]));


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
                    var content = GenerateSheet(path, sheetData);
                    if (string.IsNullOrEmpty(content))
                    {
                        continue;
                    }

                    _tempStrBuilder.AppendLine("//This file is automatically generated, please do not modify it manually");
                    _tempStrBuilder.AppendLine();
                    _tempStrBuilder.AppendLine("using System.Collections.Generic;");
                    _tempStrBuilder.AppendLine();
                    _tempStrBuilder.AppendLine($"namespace {NameSpace}");
                    _tempStrBuilder.AppendLine("{");
                    _tempStrBuilder.AppendLineWithTab($"public partial class {ClassName}");
                    _tempStrBuilder.AppendLineWithTab("{");
                    _tempStrBuilder.Append(content);
                    _tempStrBuilder.AppendLineWithTab("}");
                    _tempStrBuilder.AppendLine("}");

                    File.WriteAllText(Path.Combine(path, $"{excelName}{fieldName}Data.cs"), _tempStrBuilder.ToString());
                }
            }
        }

        private static string GenerateSheet(string path, object sheetObject)
        {
            if (sheetObject is SheetEnum sheetEnum)
            {
                return GenerateEnum(path, sheetEnum);
            }

            return string.Empty;
        }

        private static string GenerateEnum(string path, SheetEnum sheetData)
        {
            var tableName = $"{sheetData.ExcelName}{sheetData.FieldName}{SheetSuffix}";
            _tempStrBuilder.AppendLine("//This file is automatically generated, please do not modify it manually");
            _tempStrBuilder.AppendLine();
            _tempStrBuilder.AppendLine("using System.Collections.Generic;");
            _tempStrBuilder.AppendLine();
            _tempStrBuilder.AppendLine($"namespace {NameSpace}");
            _tempStrBuilder.AppendLine("{");
            _tempStrBuilder.AppendLineWithTab($"public partial class {tableName}");
            _tempStrBuilder.AppendLineWithTab("{");
            _tempStrBuilder.Append(content);
            _tempStrBuilder.AppendLineWithTab("}");
            _tempStrBuilder.AppendLine("}");

            File.WriteAllText(Path.Combine(path, $"{tableName}.cs"), _tempStrBuilder.ToString());
        }

        private static void ThrowException(ISheetData sheetDataA, ISheetData sheetDataB)
        {
            var aExcelName = sheetDataA.ExcelName;
            var aSheetName = sheetDataA.SheetName;

            var bExcelName = sheetDataB.ExcelName;
            var bSheetName = sheetDataB.SheetName;

            var sheetA = $"ExcelA[{aExcelName}] SheetA[{aSheetName}] ;";
            var sheetB = $"ExcelB[{bExcelName}] SheetB[{bSheetName}] ;";
            throw new Exception($"{sheetA} {sheetB} : A and B conflict.");
        }
    }
}