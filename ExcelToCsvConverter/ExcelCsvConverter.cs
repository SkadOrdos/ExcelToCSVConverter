using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelToCsvConverter
{
    class ExcelCsvConverter : IDisposable
    {
        private readonly Settings Settings;
        private readonly string SheetName = "sheet";


        public ExcelCsvConverter(Settings sett)
        {
            Settings = sett;
        }


        #region// ExcelToCSV

        public void ExcelToCsv()
        {
            var excelFiles = Settings.WorkFiles.ToList();
            Console.WriteLine("Excel files found: \t" + excelFiles.Count);

            int sheetsProcessed = 0;
            foreach (var file in excelFiles)
            {
                var excel = new ExcelPackage(new FileInfo(file));
                var workSheets = excel.Workbook.Worksheets.ToList();
                int sheetCount = workSheets.Count();

                if (Settings.ProcessedSheetCount > 0) sheetCount = Math.Min(sheetCount, Settings.ProcessedSheetCount);

                int bookSheetNum = 0;
                for (int s = 0; s < sheetCount; s++)
                {
                    var sheet = workSheets[s];

                    int columnsCount = sheet.Dimension.Columns;
                    // Skip first column with keys
                    for (int col = 2; col <= columnsCount; col++)
                    {
                        bookSheetNum++;
                        bool hasValues = false;

                        int rowsCount = sheet.Dimension.Rows;
                        List<string> lines = new List<string>();
                        for (int row = 1; row <= rowsCount; row++)
                        {
                            var cellKey = sheet.Cells[row, 1]?.Text;
                            if (!String.IsNullOrEmpty(cellKey))
                            {
                                var cellValue = sheet.Cells[row, col]?.Value;
                                string valueText = cellValue != null ? cellValue.ToString().Trim() : String.Empty;

                                if (!String.IsNullOrEmpty(valueText))
                                {
                                    hasValues = true;
                                    lines.Add(cellKey.Trim() + Settings.OutFileSeparator + valueText);
                                }
                            }
                        }

                        if (!hasValues) continue;
                        sheetsProcessed++;

                        lines.Sort();
                        string newFile = SheetName + bookSheetNum + Settings.GetFormatExtension;
                        File.WriteAllLines(newFile, lines.ToArray(), Encoding.UTF8);
                    }
                }

                // Close package
                excel.Dispose();
            }

            Console.WriteLine("Excel sheets processed: \t" + sheetsProcessed);
        }

        #endregion

        #region// CSVToExcel

        public void CSVToExcel()
        {
            ReadCSVLocales(Directory.GetFiles(Environment.CurrentDirectory, "*." + Settings.OutFileExtension));
        }

        private void ReadCSVLocales(IEnumerable<string> files)
        {
            var csvFiles = files.ToList();
            Console.WriteLine("CSV files found: \t" + csvFiles.Count);

            var localeTable = new Dictionary<string, IDictionary<string, string>>();
            foreach (var rf in csvFiles)
            {
                localeTable[rf] = ParseCSVFile(rf.Trim());
            }

            // Sort by rows count
            var sortedTable = localeTable.OrderByDescending(lt => lt.Value.Keys.Count);
            if (!sortedTable.Any())
                return;

            // Create Lists by files
            var localesDic = new Dictionary<string, IDictionary<string, string>>();
            foreach (var lst in sortedTable)
                localesDic.Add(lst.Key, lst.Value);

            // Create values table [Rows, Columns]
            string[] keys = localesDic.Values.SelectMany(kl => kl.Keys).Distinct().OrderBy(key => key).ToArray();
            string[,] valuesTable = new string[keys.Length, localesDic.Count + 1];

            for (int fi = 0; fi < localesDic.Count; fi++)
            {
                var localeFile = localesDic.ElementAt(fi);
                var locales = localeFile.Value;

                string value;
                int valueColumn = fi + 1;
                for (int ri = 0; ri < keys.Length; ri++)
                {
                    if (String.IsNullOrEmpty(valuesTable[ri, 0])) valuesTable[ri, 0] = keys[ri];

                    value = null;
                    string key = valuesTable[ri, 0];
                    if (locales.TryGetValue(key, out value))
                    {
                        valuesTable[ri, valueColumn] = value;
                    }
                    else
                    {
                        value = key;
                        Console.WriteLine("Error: Key {0} not found in {1}!", key, Path.GetFileName(localeFile.Key));
                    }
                }
            }

            Console.WriteLine("CSV files processed: \t" + localeTable.Count);
            SaveToExcel(valuesTable);
        }

        private IDictionary<string, string> ParseCSVFile(string file)
        {
            var codes = new Dictionary<string, string>();
            try
            {
                string[] codeLines = File.ReadAllLines(file);
                Array.Sort(codeLines);

                foreach (string line in codeLines)
                {
                    string[] split = line.Trim().Split(new string[] { Settings.OutFileSeparator }, StringSplitOptions.None);
                    if (split.Length >= 2)
                    {
                        string key = split[0].Trim();
                        string splitValue = split[1];
                        // Если в значении есть знак разделителя его нужно сохранить
                        if (split.Length > 2)
                            splitValue = String.Join(Settings.OutFileSeparator, split.Skip(1));

                        codes[key] = splitValue.Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return codes;
        }


        private void SaveToExcel(string[,] values)
        {
            string file = Path.Combine(Environment.CurrentDirectory, Settings.ExportBookFile);
            string oldFile = file + ".old";
            // Save previous file
            try
            {
                if (File.Exists(file))
                {
                    File.Copy(file, oldFile, true);
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error create backup file: " + ex.Message);
            }

            // Save sheet to new file
            try
            {
                using (var excel = new ExcelPackage())
                {
                    var xSheet = excel.Workbook.Worksheets.Add(SheetName);
                    var cellsRange = xSheet.Cells;

                    int columnsCount = values.GetLength(1);
                    int rowsCount = values.GetLength(0);
                    for (int ci = 0; ci < columnsCount; ci++)
                    {
                        xSheet.Column(ci + 1).Width = 50;

                        for (int ri = 0; ri < rowsCount; ri++)
                        {
                            string cellValue = values[ri, ci];
                            xSheet.SetValue(ri + 1, ci + 1, cellValue);
                        }
                    }

                    excel.SaveAs(new FileInfo(file));
                }
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error write to excel file: " + exc.Message);
            }
            finally
            {
                //CloseExcel();
            }
        }


        [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        #endregion

        public void Dispose()
        {

        }
    }
}
