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
        string _excelName;

        public ExcelHelper() { }

        public void CreateExcel(string excelName)
        {
            _excelName = AddExcelExtension(excelName);
            //Create excel app object
            _xlSamp = new Application();
            if (_xlSamp == null)
            {
                throw new Exception("Excel is not Insatalled");
            }

            object _misValue = System.Reflection.Missing.Value;

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
            if (ext == ".xls" || ext == ".xlsx")
            {
                return excelName;
            }
            else
            {
                return excelName + ".xls";
            }
        }

        public void OpenExcel(string excelName, string worksheetName = null)
        {
            ReleaseExcelFile();
            _excelName = AddExcelExtension(excelName);
            var pathApp = Environment.CurrentDirectory;
            var pathExcel = Path.Combine(pathApp, _excelName);

            _xlSamp = new Application();
            if (_xlSamp == null)
            {
                throw new Exception("Excel is not Insatalled");
            }

            try
            {
                _xlWorkBook = _xlSamp.Workbooks.Open(pathExcel);
                if (string.IsNullOrWhiteSpace(worksheetName))
                {
                    _xlWorkSheet = (Worksheet)_xlWorkBook.Worksheets.get_Item(1);
                }
                else
                {
                    _xlWorkSheet = (Worksheet)_xlWorkBook.Worksheets[worksheetName];
                }
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
            return Convert.ToString((_xlWorkSheet.Cells[row, column] as Microsoft.Office.Interop.Excel.Range).Value);
        }

        private string CheckValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        public string[] GetColumn(int column, bool removeEmpty = false, bool removeHeader = false)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Columns[column];
            var count = range.Rows.Count;
            var arr = range.Value;
            var list = new List<string>();
            for (int i = removeHeader ? 2 : 1; i <= count; i++)
            {
                string cellValue = Convert.ToString(arr[i, 1]);

                if (!removeEmpty)
                {
                    list.Add(CheckValue(cellValue));
                }
                else if (removeEmpty && !string.IsNullOrWhiteSpace(cellValue))
                {
                    list.Add(CheckValue(cellValue));
                }
            }
            return list.ToArray();
        }

        public Dictionary<int, string> GetColumnKeyValue(int column, bool removeEmpty = false, bool removeHeader = false)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Columns[column];
            var count = range.Rows.Count;
            var arr = range.Value;
            var dic = new Dictionary<int, string>();
            for (int i = removeHeader ? 2 : 1; i <= count; i++)
            {
                string cellValue = Convert.ToString(arr[i, 1]);

                if (!removeEmpty)
                {
                    dic[i] = CheckValue(cellValue);
                }
                else if (removeEmpty && !string.IsNullOrWhiteSpace(cellValue))
                {
                    dic[i] = cellValue;
                }
            }
            return dic;
        }

        public string[] GetRow(int row, bool removeEmpty = false)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Rows[row];
            var count = range.Columns.Count;
            var arr = range.Value;
            var list = new List<string>();
            for (int i = 1; i <= count; i++)
            {
                string cellValue = Convert.ToString(arr[1, i]);

                if (!removeEmpty)
                {
                    list.Add(CheckValue(cellValue));
                }
                else if (removeEmpty && !string.IsNullOrWhiteSpace(cellValue))
                {
                    list.Add(CheckValue(cellValue));
                }
            }
            return list.ToArray();
        }

        public Dictionary<int, string> GetRowKeyValue(int row, bool removeEmpty = false)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Rows[row];
            var count = range.Columns.Count;
            var arr = range.Value;
            var dic = new Dictionary<int, string>();
            for (int i = 1; i <= count; i++)
            {
                string cellValue = Convert.ToString(arr[1, i]);
                if (!removeEmpty)
                {
                    dic[i] = CheckValue(cellValue);
                }
                else if (removeEmpty && !string.IsNullOrWhiteSpace(cellValue))
                {
                    dic[i] = cellValue;
                }
            }
            return dic;
        }

        public int GetRowNumberByValue(string value, int column, bool match = true)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Columns[column];
            var count = range.Rows.Count;
            var arr = range.Value;
            for (int i = 1; i <= count; i++)
            {
                string cellValue = Convert.ToString(arr[i, 1]);
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

        public int GetColumnNumberByValue(string value, int row, bool matchAll = true)
        {
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)_xlWorkSheet.UsedRange.Rows[row];
            var count = range.Columns.Count;
            var arr = range.Value;
            for (int i = 1; i <= count; i++)
            {
                string cellValue = Convert.ToString(arr[1, i]);
                if (cellValue == value && matchAll)
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

            //_xlWorkBook.SaveAs(pathExcel, XlFileFormat.xlWorkbookNormal, _misValue, _misValue, _misValue, _misValue,
            //    XlSaveAsAccessMode.xlExclusive, _misValue, _misValue, _misValue, _misValue, _misValue);
            
            _xlWorkBook.SaveAs(pathExcel);
        }
        public void Quit()
        {
            //_xlWorkBook?.Close(true, _misValue, _misValue);
            _xlWorkBook?.Close();
            _xlSamp?.Quit();
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
