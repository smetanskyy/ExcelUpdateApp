using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelUpdateApp
{
    class ExcelHelper
    {
        Application _xlSamp;
        Workbook _xlWorkBook;
        Worksheet _xlWorkSheet;
        object _misValue;
        string _excelName;

        public ExcelHelper() { }

        public void CreateExcel(string excelName)
        {
            _excelName = AddExcelExtension(excelName);
            //Create excel app object
            _xlSamp = new Application();
            if (_xlSamp == null)
            {
                Console.WriteLine("Excel is not Insatalled");
                Console.ReadKey();
                return;
            }

            _misValue = System.Reflection.Missing.Value;

            //Create a new excel book and sheet
            _xlWorkBook = _xlSamp.Workbooks.Add(_misValue);
            _xlWorkSheet = (Worksheet)_xlWorkBook.Worksheets.get_Item(1);
        }

        private string AddExcelExtension(string excelName)
        {
            if (string.IsNullOrWhiteSpace(excelName))
            {
                return "book.xls";
            }
            var ext = Path.GetExtension(excelName);
            if (ext == "xls" || ext == "xlsx")
            {
                return excelName;
            }
            else
            {
                return excelName + ".xls";
            }
        }

        public void OpenExcel(string excelName)
        {
            ReleaseExcelFile();
            _excelName = AddExcelExtension(excelName);
            var pathApp = Environment.CurrentDirectory;
            var pathExcel = Path.Combine(pathApp, _excelName);

            _xlSamp = new Application();
            if (_xlSamp == null)
            {
                Console.WriteLine("Excel is not Insatalled");
                Console.ReadKey();
                return;
            }

            try
            {
                _xlWorkBook = _xlSamp.Workbooks.Open(pathExcel);
                _xlWorkSheet = (Worksheet)_xlWorkBook.Worksheets.get_Item(1);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void AddRow(int row, string[] columns)
        {
            for (int i = 0; i < columns.Length; i++)
            {
                _xlWorkSheet.Cells[row, i + 1] = columns[i];
            }
        }

        public void AddColumn(string[] rows, int column)
        {
            for (int i = 0; i < rows.Length; i++)
            {
                _xlWorkSheet.Cells[i + 1, column] = rows[i];
            }
        }

        public void FillCell(int row, int column, string value)
        {
            _xlWorkSheet.Cells[row, column] = value;
        }

        public string GetCellValue(int row, int column)
        {
            return (string)(_xlWorkSheet.Cells[row, column] as Microsoft.Office.Interop.Excel.Range).Value;
        }

        private string CheckValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value;
        }

        public string[] GetColumn(int column)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Columns[column];
            var count = range.Rows.Count;
            var arr = range.Value;
            var list = new List<string>();
            for (int i = 1; i <= count; i++)
            {
                var cellValue = arr[i, 1] as string;
                list.Add(CheckValue(cellValue));
            }
            return list.ToArray();
        }

        public string[] GetRow(int row)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Rows[row];
            var count = range.Columns.Count;
            var arr = range.Value;
            var list = new List<string>();
            for (int i = 1; i <= count; i++)
            {
                var cellValue = arr[1, i] as string;
                list.Add(CheckValue(cellValue));
            }
            return list.ToArray();
        }

        public int GetRowNumberByValue(string value, int column, bool match = true)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Columns[column];
            var count = range.Rows.Count;
            var arr = range.Value;
            for (int i = 1; i <= count; i++)
            {
                var cellValue = arr[i, 1] as string;
                if (cellValue == value && match)
                {
                    return i;
                }
                else if (cellValue != null && cellValue.Contains(value))
                {
                    return i;
                }
            }
            return -1;
        }

        public void SaveExcel(string excelName)
        {
            _excelName = AddExcelExtension(excelName);
            SaveExcel();
        }

        public void SaveExcel()
        {
            if (_xlSamp == null || _xlWorkBook == null)
            {
                return;
            }

            //Save the opened excel book to custom location
            var pathApp = Environment.CurrentDirectory;
            var excelNameWithoutExt = Path.GetFileNameWithoutExtension(_excelName);
            var ext = Path.GetExtension(_excelName);
            var pathExcel = Path.Combine(pathApp, excelNameWithoutExt + ext);

            int tryName = 0;
            string pathExcelTry = pathExcel;
            while (File.Exists(pathExcelTry))
            {
                tryName++;
                pathExcelTry = Path.Combine(pathApp, excelNameWithoutExt + tryName.ToString() + ext);
            }

            pathExcel = pathExcelTry;
            _xlWorkBook.SaveAs(pathExcel, XlFileFormat.xlWorkbookNormal, _misValue, _misValue, _misValue, _misValue,
                XlSaveAsAccessMode.xlExclusive, _misValue, _misValue, _misValue, _misValue, _misValue);

            Quit();
        }
        public void Quit()
        {
            _xlWorkBook.Close(true, _misValue, _misValue);
            _xlSamp.Quit();
            ReleaseExcelFile();
        }

        private void ReleaseExcelFile()
        {
            //release Excel Object 
            try
            {
                if (_xlSamp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSamp);
                    _xlSamp = null;
                }
            }
            catch (Exception ex)
            {
                _xlSamp = null;
                Console.Write("Error " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
