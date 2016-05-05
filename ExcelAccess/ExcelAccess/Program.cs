using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices.Marshal;

namespace check01
{
    class Program
    {
        static void Main(string[] args)
        {
            String path = System.IO.Path.GetFullPath(@".\sample.xlsx");

            //create new file
            ExcelManager.CreateExcelFile(
                path,
                (Excel.Workbook workbook) =>
                {
                    ExcelSupport.OperateCells(workbook, 1,
                        (Excel.Range cells) =>
                        {
                            for (int i = 1; i < 10; i++)
                            {
                                cells[i, 1].Value2 = "hoge";
                            }

                            return true;
                        });

                    return true;
                });

            //open file
            ExcelManager.OpenWorkbook(path, (Excel.Workbook workbook) =>
            {
                ExcelSupport.OperateCells(workbook, 1, (Excel.Range cells) =>
                {
                    for (int i = 1; i < 10; i++)
                    {
                        Console.WriteLine(cells[i, 1].Value2);
                    }

                    return true;
                });
                return true;
            });
        }
    }
}
