using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelUpdateApp
{
    class FileHelper
    {
        public static string[] Read(string file)
        {
            string pathApp = Environment.CurrentDirectory;
            var pathFile = Path.Combine(pathApp, file);

            if (string.IsNullOrWhiteSpace(file) || !File.Exists(pathFile))
            {
                return null;
            }
            // Read a text file line by line.  
            return File.ReadAllLines(file);
        }
    }
}
