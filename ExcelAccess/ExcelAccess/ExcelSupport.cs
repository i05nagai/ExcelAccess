using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices.Marshal;

namespace check01
{
    class ExcelSupport
    {
        public static bool OperateSheets(
            Excel.Workbook workbook,
            Func<Excel.Sheets, bool> handler)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Sheets;

                //delegate
                handler(sheets);
            }
            finally
            {
                Marshal.ReleaseComObject(sheets);
                sheets = null;
                GC.Collect();
            }
            return true;
        }

        public static bool OperateWorksheet(
            Excel.Workbook workbook,
            String sheetName,
            Func<Excel.Worksheet, bool> handler)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Sheets;

                //delegate
                handler(sheets[sheetName]);
            }
            finally
            {
                Marshal.ReleaseComObject(sheets);
                sheets = null;
                GC.Collect();
            }
            return true;
        }

        public static bool OperateWorksheet(
            Excel.Workbook workbook,
            int sheetIndex,
            Func<Excel.Worksheet, bool> handler)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Sheets;

                //delegate
                handler(sheets[sheetIndex]);
            }
            finally
            {
                Marshal.ReleaseComObject(sheets);
                sheets = null;
                GC.Collect();
            }
            return true;
        }

        public static bool OperateCells(
            Excel.Workbook workbook,
            int sheetIndex,
            Func<Excel.Range, bool> handler)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet worksheet = null;
            Excel.Range cells = null;
            try
            {
                sheets = workbook.Sheets;
                worksheet = sheets[sheetIndex];
                cells = worksheet.Cells;

                //delegate
                handler(cells);
            }
            finally
            {
                Marshal.ReleaseComObject(cells);
                Marshal.ReleaseComObject(sheets);
                cells = null;
                sheets = null;
                GC.Collect();
            }
            return true;
        }
    }
}
