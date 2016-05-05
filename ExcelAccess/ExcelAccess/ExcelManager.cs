using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices.Marshal;

namespace check01
{
    class ExcelManager
    {
        public static void CreateExcelFile(
            String path,
            Func<Excel.Workbook, bool> handler)
        {
            //Excelオブジェクトの初期化
            Excel.Application excel = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            try
            {
                //Excelシートのインスタンスを作る
                excel = new Excel.Application();
                excel.Visible = false;
                workbooks = excel.Workbooks;
                workbook = workbooks.Add();

                //delate
                handler(workbook);

                workbook.SaveAs(path);
                workbook.Close(false);
                excel.Quit();
            }
            finally
            {
                //to release excel process
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(excel);
                workbook = null;
                workbooks = null;
                excel = null;
                GC.Collect();
            }
        }

        public static void OpenWorkbook(
            String path,
            Func<Excel.Workbook, bool> handler)
        {
            //Excelオブジェクトの初期化
            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            try
            {
                excel = new Excel.Application();
                excel.Visible = false;
                workbook = excel.Workbooks.Open(path,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                //delegate
                handler(workbook);

                workbook.Save();
                excel.Quit();
            }
            finally
            {
                //to release excel process
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excel);
                workbook = null;
                excel = null;
                GC.Collect();
            }
        }
    }
}
