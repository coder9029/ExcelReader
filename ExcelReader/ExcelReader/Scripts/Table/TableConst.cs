using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace ExcelReader;

public class TableConst : ITableData
{
    private readonly List<IRow> iRowData;

    public TableConst(Dictionary<int, SheetHeader> header, int index)
    {
    }
}