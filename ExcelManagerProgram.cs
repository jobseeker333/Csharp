//Wrapper class to simplify the creation of Excel files in C# 4.0
//This class uses the improved COM Interop in C# 4.0 and makes it simple to create Excel files from data
//in your application. Note that you have to add a reference to the Microsoft Excel 14.0 Object Library in your
//Visual Studio 6.0 project.    

using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelManager em = new ExcelManager(@"c:\\Test.xlsx");
            em.FillCell(1, 2, "Jan");
            em.FillCell(1, 3, "Feb");
            em.FillCell(1, 4, "Mar");
            em.FillCell(2, 1, "Boston");
            em.FillCell(3, 1, "New York");
            em.Save();
        }
    }

    public class ExcelManager
    {
        Application excelApp;
        Workbook _wb;
        Worksheet _ws;
        _range _range;
        string _filePath;

        public ExcelManager(string filePath, string sheetName = "Data")
        {
            _filePath = filePath;
            excelApp = new Application();
            if (excelApp == null)
                throw new Exception("Excel could not be started.  Check your Office installation.");
            //excelApp.Visible = true; //uncomment to see it being created
            _wb = excelApp.Workbooks.Add(Xl_wbATemplate.xl_wbATWorksheet);
            _ws = (Worksheet)_wb.Worksheets[1];
            if (_ws == null)
                throw new Exception("The workbook could not be created.  Check your Office installation.");
            _ws.Name = sheetName;
            _range = _ws.Used_range;
        }

        public void FillCell(int row, int column, string value)
        {
            _ws.Cells[row, column] = value;
        }

        public void Save()
        {
            _wb.SaveAs(filePath);
            _wb.Close();
        }
    }
}
