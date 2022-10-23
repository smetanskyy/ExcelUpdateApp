using System;
using System.Linq;

namespace ExcelUpdateApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Process has started \n");
            string srcFilename;
            string srcSheet;
            string srcMatch;
            string destFilename;
            string destSheet;
            string destMatch;

            try
            {
                var src = FileHelper.Read("_Source.txt");
                srcFilename = src[0].Split("-->", StringSplitOptions.RemoveEmptyEntries).Last().Trim();
                var srcSecondLine = src[1].Split("-->", StringSplitOptions.RemoveEmptyEntries);
                srcSheet = srcSecondLine.Length == 1 ? string.Empty : srcSecondLine.Last().Trim();
                srcMatch = src[2].Split("-->", StringSplitOptions.RemoveEmptyEntries).Last().Trim();

                var dest = FileHelper.Read("_Destination.txt");
                destFilename = dest[0].Split("-->", StringSplitOptions.RemoveEmptyEntries).Last().Trim();
                var destSecondLine = dest[1].Split("-->", StringSplitOptions.RemoveEmptyEntries);
                destSheet = destSecondLine.Length == 1 ? string.Empty : destSecondLine.Last().Trim();
                destMatch = dest[2].Split("-->", StringSplitOptions.RemoveEmptyEntries).Last().Trim();

                var display = string.IsNullOrWhiteSpace(srcSheet) ? "1" : srcSheet;
                Console.WriteLine($"SRC: EXCEL -> \"{srcFilename}\" SHEET -> \"{display}\" MATCH -> \"{srcMatch}\"");
                display = string.IsNullOrWhiteSpace(destSheet) ? "1" : destSheet;
                Console.WriteLine($"DEST: EXCEL -> \"{destFilename}\" SHEET -> \"{display}\" MATCH -> \"{destMatch}\"");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                return;
            }

            ExcelHelper source = new ExcelHelper();
            ExcelHelper destination = new ExcelHelper();

            try
            {
                Console.WriteLine($"Loading \"{srcFilename}\" ... ");

                source.OpenExcel(srcFilename, srcSheet);
                var srcHeaders = source.GetRowKeyValue(1, true);
                var srcColNumber = srcHeaders.FirstOrDefault(x => x.Value == srcMatch).Key;
                var keys = source.GetColumnKeyValue(srcColNumber, true, true);

                Console.WriteLine();
                Console.WriteLine($"Loading \"{destFilename}\" ... \n");

                destination.OpenExcel(destFilename, destSheet);
                var destHeaders = destination.GetRowKeyValue(1, true);
                var destColNumber = destHeaders.FirstOrDefault(x => x.Value == destMatch).Key;

                foreach (var key in keys)
                {
                    var curRowNumberSrc = key.Key;
                    var curRowNumberDest = destination.GetRowNumberByValue(key.Value, destColNumber, false);
                    var display = string.Empty;
                    foreach (var header in srcHeaders)
                    {
                        var curColumnNumberSrc = header.Key;
                        var valueCellSrc = source.GetCellValue(curRowNumberSrc, curColumnNumberSrc);
                        display += $"{valueCellSrc} ";
                        if (curRowNumberDest < 1)
                        {
                            continue;
                        }
                        var curHeaderDest = destHeaders.Where(x => string.Equals(x.Value, header.Value, StringComparison.OrdinalIgnoreCase));
                        if (curHeaderDest == null || curHeaderDest.Count() > 1)
                        {
                            continue;
                        }
                        destination.FillCell(curRowNumberDest, curHeaderDest.FirstOrDefault().Key, valueCellSrc);
                    }
                    if (curRowNumberDest < 1)
                    {
                        Console.WriteLine($"NOT FOUND: {display}");
                    }
                    else
                    {
                        Console.WriteLine($"ADDED: {display}");
                    }
                }

                Console.WriteLine();
                Console.WriteLine($"Saving ...");
                destination.SaveExcel();
                Console.WriteLine($"Excel file has been saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
            }
            finally
            {
                source.Quit();
                destination.Quit();
                Console.WriteLine();
                Console.WriteLine("Process has ended!");
                Console.WriteLine("Enter any key for exit ... ");
            }
            // var wb = xlApp.Workbooks.Open(fn, ReadOnly: false);
            // wb.Close(SaveChanges: true);
            // xlApp.Quit();
        }
    }
}
