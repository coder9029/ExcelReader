using System;
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

        public static ClassParam GetClassParam(this ISheet sheet)
        {
            var cellComment = sheet.GetRow(0)?.GetCell(0)?.CellComment.String.ToString();
            if (string.IsNullOrEmpty(cellComment))
            {
                throw new Exception($"单[{sheet.SheetName}]: The sheet first cell comment is null.");
            }

            string className = null;
            string classType = null;

            var arrayComment = cellComment.Split('\n');
            foreach (var item in arrayComment)
            {
                var arrayTemp = item.Split(':');
                if (arrayTemp.Length < 2)
                {
                    continue;
                }

                var label = arrayTemp[0];
                var value = arrayTemp[1];

                if (label == "Name")
                {
                    className = value;
                    continue;
                }

                if (label == "Type")
                {
                    classType = value;
                    continue;
                }
            }

            if (string.IsNullOrEmpty(className) || string.IsNullOrEmpty(classType) || !Enum.TryParse(classType, out ClassType eClassType))
            {
                throw new Exception($"Sheet[{sheet.SheetName}] - Comment[{cellComment}]: The sheet first cell comment value is empty or invalid.");
            }

            return new ClassParam(className, eClassType);
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