﻿using System;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace Config
{
    public static class ExcelReaderHelper
    {
        public static StringBuilder AppendLineWithTab(this StringBuilder strBuilder, string value, int count = 1)
        {
            var str = string.Empty;
            for (var i = 0; i < count; i++)
            {
                str += '\t';
            }

            strBuilder.AppendLine($"{str}" + value);
            return strBuilder;
        }

        public static ISheetData GetSheetParam(this ISheet sheetData, string excelName)
        {
            var firstArray = sheetData.GetRow(0)?.GetCell(0)?.ToString()?.Split(' ');
            if (firstArray.Length < 1)
            {
                throw new Exception($"Sheet[{sheetData.SheetName}]: The sheet first cell does not comply with the rule.");
            }

            var sheetName = sheetData.SheetName;
            var sheetType = firstArray[1];

            if (sheetType == "Enum")
            {
                return new SheetEnum(sheetData, excelName);
            }

            if (sheetType == "Const")
            {
                return new SheetConst(sheetData, excelName);
            }

            if (sheetName.Contains('@'))
            {
                return new SheetMerge(sheetData, excelName);
            }
            else
            {
                return new SheetValue(sheetData, excelName);
            }
        }

        public static bool GetFieldType(this string str, out string type)
        {
            type = str;
            return true;
        }

        public static string GetFiledValue(this string str, string type)
        {
            if (type == "bool")
            {
                return str.ToLower();
            }

            if (type == "float")
            {
                return $"{str}f";
            }

            if (type == "string")
            {
                return str.StartsWith("$\"") && str.EndsWith('"') ? str : $"\"{str}\"";
            }

            if (type.EndsWith("[]") && !type.EndsWith("[][]"))
            {
                var keyType = type[..^"[]".Length];
                var keyValue = string.Empty;

                var tempValue = str.Split(',');
                foreach (var tempItem in tempValue)
                {
                    if (string.IsNullOrEmpty(tempItem))
                    {
                        continue;
                    }

                    var realValue = tempItem.GetFiledValue(keyType);

                    if (string.IsNullOrEmpty(keyValue))
                    {
                        keyValue = realValue;
                    }
                    else
                    {
                        keyValue += $", {realValue}";
                    }
                }

                return $"new[] {{ {keyValue} }}";
            }

            if (type.EndsWith("[][]"))
            {
                var keyType = type[..^"[][]".Length];
                var keyValue = string.Empty;

                var tempValue = str.Split("),");
                foreach (var tempItem in tempValue)
                {
                    var tempArray = tempItem;
                    if (tempItem.StartsWith('('))
                    {
                        tempArray = tempItem[1..];
                    }

                    if (string.IsNullOrEmpty(tempArray))
                    {
                        continue;
                    }

                    var realValue = tempArray.GetFiledValue($"{keyType}[]");

                    if (string.IsNullOrEmpty(keyValue))
                    {
                        keyValue = realValue;
                    }
                    else
                    {
                        keyValue += $", {realValue}";
                    }
                }

                return $"new[] {{ {keyValue} }}";
            }

            return str;
        }

        public static bool IsDictKeyValid(this string str)
        {
            if (str is "int" or "long" or "string")
            {
                return true;
            }

            //枚举值，先暂定
            // if (str.StartsWith("E"))
            // {
            //     return true;
            // }

            return false;
        }
    }
}