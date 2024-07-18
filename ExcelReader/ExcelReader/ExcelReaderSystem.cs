using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Config
{
    public static class ExcelReaderSystem
    {
        private const string NameSpace = "Config";

        private static Dictionary<string, string> _constCollector = new();

        private static Dictionary<string, string> _classCollector = new();

        private static Dictionary<string, string> _valueCollector = new();

        private static StringBuilder _tempStrBuilder = new();

        private static Dictionary<int, (string, string)> _tempSheetColumn = new();

        public static void Program(string dirPath, string outPath)
        {
            var files = Directory.GetFiles(dirPath, string.Empty, SearchOption.AllDirectories);
            foreach (var file in files)
            {
                var fileName = Path.GetFileNameWithoutExtension(file);
                if (fileName.StartsWith('~'))
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

            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                try
                {
                    if (sheet == null || sheet.SheetName.StartsWith('#'))
                    {
                        continue;
                    }

                    if (sheet.SheetName.StartsWith('@'))
                    {
                        CollectConst(sheet);
                    }
                    else
                    {
                        var cellComment = sheet.GetRow(0)?.GetCell(0)?.CellComment?.ToString();
                        if (string.IsNullOrEmpty(cellComment))
                        {
                            throw new Exception($"Sheet[{sheet.SheetName}]: The cell comment is null or empty.");
                        }
                        
                        CollectClass(sheet);
                        CollectValue(sheet);
                    }
                }
                catch (Exception exception)
                {
                    throw new Exception($"Excel[{file}]: There are some errors thrown.\n" + exception);
                }
            }
        }

        private static void CollectConst(ISheet sheet)
        {
            if (_constCollector.ContainsKey(sheet.SheetName))
            {
                return;
            }

            _tempStrBuilder.Clear();

            var isFirstWrite = false;
            for (var i = 0; i <= sheet.LastRowNum; i++)
            {
                var rowItem = sheet.GetRow(i);

                var field = rowItem?.GetCell(0)?.ToString();
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
                    _tempStrBuilder.AppendLine();
                }

                var summary = rowItem.GetCell(2)?.ToString();
                if (!string.IsNullOrEmpty(summary))
                {
                    _tempStrBuilder.AppendLineWithTab($"/// <summary>{summary}</summary>", 2);
                }

                var type = rowItem.GetCell(1)?.ToString();
                if (string.IsNullOrEmpty(type) || !type.GetFieldType(out var typeString))
                {
                    throw new Exception($"Sheet[{sheet.SheetName}] - Row[{i + 1}] - Col[2] - Value[{type}]: The value is empty or invalid.");
                }

                var value = rowItem.GetCell(3)?.ToString();
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception($"Sheet[{sheet.SheetName}] - Row[{i + 1}] - Col[3] - Value[{value}]: The value is empty.");
                }

                var valueString = value.GetFiledValue(type);
                _tempStrBuilder.AppendLineWithTab($"public {typeString} {field} = {valueString};", 2);
            }

            if (_tempStrBuilder.Length == 0)
            {
                return;
            }

            _constCollector.Add(sheet.SheetName, _tempStrBuilder.ToString());
        }

        private static void CollectClass(ISheet sheet)
        {
            var classParam = sheet.GetClassParam();
            if (_classCollector.ContainsKey(classParam.Name))
            {
                return;
            }

            if (sheet.LastRowNum < 3)
            {
                throw new Exception($"Sheet[{sheet.SheetName}]: The sheet is invalid.");
            }

            _tempStrBuilder.Clear();
            var rowField = sheet.GetRow(0);
            var rowType = sheet.GetRow(1);
            var rowSummary = sheet.GetRow(2);

            var isFirstWrite = false;
            for (var i = 0; i < rowField.LastCellNum; i++)
            {
                var field = rowField.GetCell(i)?.ToString();
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
                    _tempStrBuilder.AppendLine();
                }

                var summary = rowSummary.GetCell(i)?.ToString();
                if (!string.IsNullOrEmpty(summary))
                {
                    _tempStrBuilder.AppendLineWithTab($"/// <summary>{summary}</summary>", 2);
                }

                var type = rowType.GetCell(i)?.ToString();
                if (string.IsNullOrEmpty(type) || !type.GetFieldType(out var typeString))
                {
                    throw new Exception($"Sheet[{sheet.SheetName}] - Row[3] - Col[{i + 1}] - Value[{type}]: The value is empty or invalid.");
                }

                _tempStrBuilder.AppendLineWithTab($"public {typeString} {field};", 2);
            }

            if (_tempStrBuilder.Length == 0)
            {
                return;
            }

            var contentString = _tempStrBuilder.ToString();
            _tempStrBuilder.Clear();
            _tempStrBuilder.AppendLine("//This file is automatically generated, please do not modify it manually");
            _tempStrBuilder.AppendLine();
            _tempStrBuilder.AppendLine("using System.Collections.Generic;");
            _tempStrBuilder.AppendLine();
            _tempStrBuilder.AppendLine($"namespace {NameSpace}");
            _tempStrBuilder.AppendLine("{");
            _tempStrBuilder.AppendLineWithTab($"public partial class {classParam.Name}");
            _tempStrBuilder.AppendLineWithTab("{");
            _tempStrBuilder.Append(contentString);
            _tempStrBuilder.AppendLineWithTab("}");
            _tempStrBuilder.AppendLine("}");
            _classCollector.Add(classParam.Name, _tempStrBuilder.ToString());
        }

        private static void CollectValue(ISheet sheet)
        {
            var classParam = sheet.GetClassParam();
            if (_valueCollector.ContainsKey(classParam.Name))
            {
                return;
            }

            if (sheet.LastRowNum < 4)
            {
                return;
            }

            _tempStrBuilder.Clear();
            var rowField = sheet.GetRow(0);
            var rowType = sheet.GetRow(1);

            if (classParam.Type is ClassType.List)
            {
                _tempSheetColumn.Clear();
                for (var i = 0; i < rowField.LastCellNum; i++)
                {
                    var field = rowField.GetCell(i)?.ToString();
                    if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
                    {
                        continue;
                    }

                    var type = rowType.GetCell(i)?.ToString();
                    if (string.IsNullOrEmpty(type) || !type.GetFieldType(out _))
                    {
                        throw new Exception($"Sheet[{sheet.SheetName}] - Row[2] - Col[{i + 1}] - Value[{type}]: The value is empty or invalid.");
                    }

                    _tempSheetColumn.Add(i, (field, type));
                }

                if (_tempSheetColumn.Count == 0)
                {
                    return;
                }

                for (var i = 3; i <= sheet.LastRowNum; i++)
                {
                    var rowTemp = sheet.GetRow(i);

                    var keyValueTemp = rowTemp.GetCell(0)?.ToString();
                    if (string.IsNullOrEmpty(keyValueTemp) || keyValueTemp.StartsWith('#'))
                    {
                        continue;
                    }

                    var contentString = _tempStrBuilder.ToString();
                    _tempStrBuilder.Clear();

                    for (var j = 0; j < rowTemp.LastCellNum; j++)
                    {
                        if (!_tempSheetColumn.TryGetValue(j, out (string field, string type) colTemp))
                        {
                            continue;
                        }

                        var value = rowTemp.GetCell(j)?.ToString();
                        if (string.IsNullOrEmpty(value))
                        {
                            continue;
                        }

                        var realValue = value.GetFiledValue(colTemp.type);
                        _tempStrBuilder.AppendLineWithTab($"{colTemp.field} = {realValue},", 4);
                    }

                    var valueString = _tempStrBuilder.ToString();
                    _tempStrBuilder.Clear();
                    _tempStrBuilder.Append(contentString);

                    if (string.IsNullOrEmpty(valueString))
                    {
                        continue;
                    }

                    _tempStrBuilder.AppendLineWithTab($"new {classParam.Name}", 3);
                    _tempStrBuilder.AppendLineWithTab("{", 3);
                    _tempStrBuilder.Append(valueString);
                    _tempStrBuilder.AppendLineWithTab("},", 3);
                }

                if (_tempStrBuilder.Length == 0)
                {
                    return;
                }

                var listString = _tempStrBuilder.ToString();
                _tempStrBuilder.Clear();
                _tempStrBuilder.AppendLineWithTab($"public static IReadOnlyList<{classParam.Name}> {sheet.SheetName} = new List<{classParam.Name}>()", 2);
                _tempStrBuilder.AppendLineWithTab("{", 2);
                _tempStrBuilder.Append(listString);
                _tempStrBuilder.AppendLineWithTab("};", 2);
            }

            if (classParam.Type is ClassType.Dict)
            {
                var keyType = rowType.GetCell(0)?.ToString();
                if (string.IsNullOrEmpty(keyType) || !keyType.GetFieldType(out var keyTypeString))
                {
                    throw new Exception($"Sheet[{sheet.SheetName}] - Row[{2}] - Col[{1}] - Value[{keyType}]: The value is empty or invalid.");
                }

                if (!keyType.IsDictKeyValid())
                {
                    throw new Exception($"Sheet[{sheet.SheetName}] - Row[{2}] - Col[{1}] - Value[{keyType}]: The dictionary key is invalid.");
                }

                _tempSheetColumn.Clear();
                for (var i = 0; i < rowField.LastCellNum; i++)
                {
                    var field = rowField.GetCell(i)?.ToString();
                    if (string.IsNullOrEmpty(field) || field.StartsWith('#'))
                    {
                        continue;
                    }

                    var type = rowType.GetCell(i)?.ToString();
                    if (string.IsNullOrEmpty(type) || !type.GetFieldType(out _))
                    {
                        throw new Exception($"Sheet[{sheet.SheetName}] - Row[2] - Col[{i + 1}] - Value[{type}]: The value is empty or invalid.");
                    }

                    _tempSheetColumn.Add(i, (field, type));
                }

                if (_tempSheetColumn.Count == 0)
                {
                    return;
                }

                for (var i = 3; i <= sheet.LastRowNum; i++)
                {
                    var rowTemp = sheet.GetRow(i);

                    var keyValueTemp = rowTemp.GetCell(0)?.ToString();
                    if (string.IsNullOrEmpty(keyValueTemp) || keyValueTemp.StartsWith('#'))
                    {
                        continue;
                    }

                    var contentString = _tempStrBuilder.ToString();
                    _tempStrBuilder.Clear();

                    for (var j = 0; j < rowTemp.LastCellNum; j++)
                    {
                        if (!_tempSheetColumn.TryGetValue(j, out (string field, string type) colTemp))
                        {
                            continue;
                        }

                        var value = rowTemp.GetCell(j)?.ToString();
                        if (string.IsNullOrEmpty(value))
                        {
                            continue;
                        }

                        var realValue = value.GetFiledValue(colTemp.type);
                        _tempStrBuilder.AppendLineWithTab($"{colTemp.field} = {realValue},", 7);
                    }

                    var valueString = _tempStrBuilder.ToString();
                    _tempStrBuilder.Clear();
                    _tempStrBuilder.Append(contentString);

                    if (string.IsNullOrEmpty(valueString))
                    {
                        continue;
                    }

                    var keyValueString = keyValueTemp.GetFiledValue(keyType);
                    _tempStrBuilder.AppendLineWithTab($"case {keyValueString}:", 4);
                    _tempStrBuilder.AppendLineWithTab("{", 4);
                    _tempStrBuilder.AppendLineWithTab($"if (!_m{sheet.SheetName}.ContainsKey({keyValueString}))", 5);
                    _tempStrBuilder.AppendLineWithTab("{", 5);
                    _tempStrBuilder.AppendLineWithTab($"var data = new {classParam.Name}", 6);
                    _tempStrBuilder.AppendLineWithTab("{", 6);
                    _tempStrBuilder.Append(valueString);
                    _tempStrBuilder.AppendLineWithTab("};", 6);
                    _tempStrBuilder.AppendLineWithTab($"_m{sheet.SheetName}[{keyValueString}] = data;", 6);
                    _tempStrBuilder.AppendLineWithTab("}", 5);
                    _tempStrBuilder.AppendLine();
                    _tempStrBuilder.AppendLineWithTab($"return _m{sheet.SheetName}[{keyValueString}];", 5);
                    _tempStrBuilder.AppendLineWithTab("}", 4);
                }

                if (_tempStrBuilder.Length == 0)
                {
                    return;
                }

                var dictString = _tempStrBuilder.ToString();
                _tempStrBuilder.Clear();
                _tempStrBuilder.AppendLineWithTab($"private static Dictionary<{keyTypeString}, {classParam.Name}> _m{sheet.SheetName} = new();", 2);
                _tempStrBuilder.AppendLineWithTab($"public static {classParam.Name} {sheet.SheetName}({keyTypeString} key)", 2);
                _tempStrBuilder.AppendLineWithTab("{", 2);
                _tempStrBuilder.AppendLineWithTab("switch (key)", 3);
                _tempStrBuilder.AppendLineWithTab("{", 3);
                _tempStrBuilder.Append(dictString);
                _tempStrBuilder.AppendLineWithTab("}", 3);
                _tempStrBuilder.AppendLine();
                _tempStrBuilder.AppendLineWithTab("return null;", 3);
                _tempStrBuilder.AppendLineWithTab("}", 2);
            }

            _valueCollector.Add(classParam.Name, _tempStrBuilder.ToString());
        }

        private static void GenerateScript(string name, string path)
        {
            foreach (var (table, value) in _classCollector)
            {
                File.WriteAllText(Path.Combine(path, $"{table}.cs"), value);
            }

            _tempStrBuilder.Clear();

            foreach (var (_, value) in _constCollector)
            {
                _tempStrBuilder.Append(value);
            }

            foreach (var (_, value) in _valueCollector)
            {
                if (_tempStrBuilder.Length != 0)
                {
                    _tempStrBuilder.AppendLine();
                }

                _tempStrBuilder.Append(value);
            }

            if (_tempStrBuilder.Length != 0)
            {
                var contentString = _tempStrBuilder.ToString();
                _tempStrBuilder.Clear();

                _tempStrBuilder.AppendLine("//This file is automatically generated, please do not modify it manually");
                _tempStrBuilder.AppendLine();
                _tempStrBuilder.AppendLine("using System.Collections.Generic;");
                _tempStrBuilder.AppendLine();
                _tempStrBuilder.AppendLine($"namespace {NameSpace}");
                _tempStrBuilder.AppendLine("{");
                _tempStrBuilder.AppendLineWithTab($"public partial class {name}");
                _tempStrBuilder.AppendLineWithTab("{");
                _tempStrBuilder.Append(contentString);
                _tempStrBuilder.AppendLineWithTab("}");
                _tempStrBuilder.AppendLine("}");

                File.WriteAllText(Path.Combine(path, $"{name}.cs"), _tempStrBuilder.ToString());
            }

            _classCollector.Clear();
            _constCollector.Clear();
            _valueCollector.Clear();
        }
    }
}