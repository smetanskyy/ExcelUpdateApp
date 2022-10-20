using System;

namespace ExcelUpdateApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Process starts!");
            ExcelHelper excel = new ExcelHelper();
            try
            {
                //excel.CreateExcel("test");
                //excel.FillCell(1, 1, "TEST1");
                //excel.AddColumn(new string[] { "TEST2", "TEST3" }, 2);
                //excel.AddRow(3, new string[] { "TEST4", "TEST5" });
                //excel.SaveExcel();
                excel.OpenExcel("test");
                //excel.FillCell(1, 1, "TEST6");
                //excel.SaveExcel();
                var cell = excel.GetCellValue(1, 1);
                var cells_col_1 = excel.GetColumn(1);
                var cells_col_2 = excel.GetColumn(2);
                var cells_row = excel.GetRow(1);
                var row_number = excel.GetRowNumberByValue("TEST7", 1);
                var row_number_not_match = excel.GetRowNumberByValue("for", 1, false);

                Console.WriteLine(row_number);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                excel.Quit();
                Console.WriteLine("Process ends!");
            }
            // var wb = xlApp.Workbooks.Open(fn, ReadOnly: false);
            // wb.Close(SaveChanges: true);
            // xlApp.Quit();
        }
    }
}
