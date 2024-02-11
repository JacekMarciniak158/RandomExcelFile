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
            //Create folder in the app path for creating Excel file inside / folder is created in .exe directory 
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
                //Handle any exceptions that may occur
                Console.WriteLine("Error creating folder in directory: " + folderPath + " " + ex.Message);
            }

            try
            {
                //Create instance of Excel application
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook;
                Worksheet worksheet;

                //Create worksheet in Excel application 
                workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                worksheet = (Worksheet)workbook.Worksheets[1];

                worksheet.Name = "RandomExcelFile";

                //Set range for cells to fill
                Range headlineCellRange = worksheet.Range["A1", "F1"];
                //Array with column headlines
                string[] headlines = new[] { "ID", "Name", "Surname", "Amount", "Unit price", "Cost" };

                //Fillcells with values from array and setting font to bold to show that those are headlines
                headlineCellRange.Font.Bold = true;
                headlineCellRange.Value = headlines;

                //Create range for ids and filling it with samples
                Range idCellRange = worksheet.Range["A2", "A20"];
                int[,] ids = new int[19, 1];
                for (int i = 0; i < 19; i++)
                {
                    ids[i, 0] = i;
                }
                idCellRange.Value = ids;

                //Create range for names and filling it with samples
                Range namesCellRange = worksheet.Range["B2", "B20"];
                string[,] names = new string[19, 1];
                for (int i = 0; i < 19; i++)
                {
                    names[i, 0] = "Name " + i;
                }
                namesCellRange.Value = names;

                //Create range for surnames and filling it with samples
                Range surnameCellRange = worksheet.Range["C2", "C20"];
                string[,] surnames = new string[19, 1];
                for (int i = 0; i < 19; i++)
                {
                    surnames[i, 0] = "Surname " + i;
                }
                surnameCellRange.Value = surnames;

                //Create range for amount and fill it with random intigers
                Range amountCellRange = worksheet.Range["D2", "D20"];
                int[,] amounts = new int[19, 1];
                Random amountRandom = new Random();
                for (int i = 0; i < 19; i++)
                {
                    amounts[i, 0] = amountRandom.Next(100, 1000);
                }
                amountCellRange.Value = amounts;

                //Create range for unit price and fill it with random to two decimal place numbers
                Range unitPriceCellRange = worksheet.Range["E2", "E20"];
                double[,] unitPrices = new double[19, 1];
                int unitPriceRange = 60;
                Random costRandom = new Random();
                for (int i = 0; i < 19; i++)
                {
                    unitPrices[i, 0] += costRandom.NextDouble() * unitPriceRange;
                    unitPrices[i, 0] = Math.Round(unitPrices[i, 0], 2);
                }
                unitPriceCellRange.Value = unitPrices;

                //Create range for cost and fill it with corresponding values
                Range costCellRange = worksheet.Range["F2", "F20"];
                double[,] costs = new double[19, 1];
                for (int i = 0; i < 19; i++)
                {
                    costs[i, 0] = amounts[i, 0] * unitPrices[i, 0];
                }
                costCellRange.Value = costs;

                //Fit column width to values
                worksheet.Columns.AutoFit();

                //Create path for folder, save and close app
                string excelFilePath = Path.Combine(folderPath, "RandomExcelFile.xlsx");
                workbook.SaveAs(excelFilePath);
                workbook.Close();
                excel.Quit();
            }
            catch
            {
                //If excelfile oppened write to console and wait for anything to be clicked
                Console.WriteLine("Please close opened RandomExcel file");
                Console.ReadKey();
            }

        }
    }
}
