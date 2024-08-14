using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class TableEnum(SheetEnum sheetEnum) : ITableData
{
    private readonly SheetEnum _sheetEnum = sheetEnum;

    public string GenerateEnum()
    {
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
                throw new Exception($"Excel[{sheetEnum.ExcelName}] Sheet[{sheetEnum.SheetName}] - Row[{i + 1}] - Col[{valueColumn + 1}] - Value[{value}]: The value is empty.");
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
}