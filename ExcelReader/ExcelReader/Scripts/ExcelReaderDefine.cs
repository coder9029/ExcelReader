using System;
using System.Collections;
using System.Collections.Generic;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace ExcelReader
{

    public struct SheetHeader
    {
        public string FieldName;
        public string FieldType;
        public ESheetCollector ESheetCollector;
    }

    public enum ESheetCollector
    {
        None,
        List,
        Dict,
    }
}