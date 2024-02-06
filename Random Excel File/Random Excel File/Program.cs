using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Random_Excel_File
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Create folder in the app path for creating Excel file inside 
            var systemPath = System.Environment. GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            var complete = Path.Combine(systemPath, "RandomExcel");

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook;
            Worksheet worksheet;

            workbook = excel.Workbooks.Open(complete);
            worksheet = workbook.Worksheets[0];

            Range cellRange = worksheet.Range["A1:D1"];
            string[] strings = new[] {"test1", "testDwa", "testyTrzy","to powinno być cztery"};

            cellRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, strings);

            workbook.SaveAs("RandomExcelFile");
            workbook.Close();

            Process.Start("RandomExcelFIle");
        }
    }
}
