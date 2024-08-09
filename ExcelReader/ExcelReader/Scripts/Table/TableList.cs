using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class TableList : ITableCollector
{
    private readonly List<IRow> iRowData;

    public TableList(Dictionary<int, SheetHeader> header, int index)
    {
    }

    public void AddRow(IRow iRow)
    {
        iRowData.Add(iRow);
    }

    private string GetContent(int retract)
    {
        var stringBuilder = ExcelReaderSystem.StringBuilder;
        stringBuilder.Clear();

        foreach (var iRowTemp in iRowData)
        {
            var keyValueTemp = iRowTemp.GetCell(0)?.ToString();
            if (string.IsNullOrEmpty(keyValueTemp) || keyValueTemp.StartsWith('#'))
            {
                continue;
            }


            var contentString = stringBuilder.ToString();
            stringBuilder.Clear();

            for (var j = 0; j < iRowTemp.LastCellNum; j++)
            {
                if (!_tempSheetColumn.TryGetValue(j, out (string field, string type) colTemp))
                {
                    continue;
                }

                var value = iRowTemp.GetCell(j)?.ToString();
                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }

                var realValue = value.GetFiledValue(colTemp.type);
                stringBuilder.AppendTab($"{colTemp.field} = {realValue},", retract + 1);
            }

            var valueString = stringBuilder.ToString();
            stringBuilder.Clear();
            stringBuilder.Append(contentString);

            if (string.IsNullOrEmpty(valueString))
            {
                continue;
            }

            stringBuilder.AppendTab($"new {classParam.ScriptName}", retract);
            stringBuilder.AppendTab("{", retract);
            stringBuilder.Append(valueString);
            stringBuilder.AppendTab("},", retract);
        }

        return stringBuilder.ToString();
    }
}