using System;
using Microsoft.Office.Interop.Excel;
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
            //Creating folder in the app path for creating Excel file inside 
            string directoryPath = AppDomain.CurrentDomain.BaseDirectory;
            string folderName = "RandomExcelFolder";
            string folderPath = Path.Combine(directoryPath, folderName);

            try
            {
                //Attempt to create the directory
                Directory.CreateDirectory(folderPath);
                Console.WriteLine("Folder created successfully." + folderPath);
            }
            catch (Exception ex)
            {
                //Handling any exceptions that may occur
                Console.WriteLine("Error creating folder in directory: " + folderPath + " " + ex.Message);
            }

            //Creating instance of Excel application
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook;
            Worksheet worksheet;

            //Creating worksheet in Excel application 
            workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            worksheet = (Worksheet)workbook.Worksheets[1];

            worksheet.Name = "RandomExcelFile";

            //Setting range for cells to fill
            Range headlineCellRange = worksheet.Range["A1:F1"];
            //Array with column headlines
            string[] headlines = new[] {"Name", "Surname", "ID", "Amount", "Cost", "Unit price"};

            //Filling cells with values from array and setting font to bold to show that those are headlines
            headlineCellRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, headlines);
            headlineCellRange.Style.Font.Bold = true;

            //Creating path for folder, saving and closing app
            string excelFilePath = Path.Combine(folderPath, "RandomExcelFile.xlsx");

            workbook.SaveAs(excelFilePath);
            workbook.Close();

        }
    }
}
