using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelMapping.Extenstions
{
    public static class ExcelWorkSheetExtenstion
    {
        public static List<string> GettingHeaders(this ExcelWorksheet workSheet)
        {
            return workSheet.Cells[1, 1, 1, workSheet.Dimension.Columns].Select(a => a.Text.Trim()).ToList();
        }
    }
}
