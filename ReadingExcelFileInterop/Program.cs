using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadingExcelFileInterop
{
    public class Program
    {
        static StreamWriter logWriter;
        static String logfilePath;
        static StreamWriter writer;

        public static void Main(string[] args)
        {
            var inicio = DateTime.Now;
            var produtosDataTable = ReadExcelFile();
            var final = DateTime.Now;
        }

        static private System.Drawing.Color GetBackgroundColor(
        Microsoft.Office.Interop.Excel.Workbook WorkbookObject,
        string SheetName,
        uint LocationX, uint LocationY            //  Location.Item1 is LocationX, and Location.Item2 is LocationY
            )
        {
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorksheet = WorkbookObject.Worksheets[SheetName];
            return System.Drawing.ColorTranslator.FromOle(
                System.Convert.ToInt32(ExcelWorksheet.Cells[LocationX + 1, LocationY + 1].Interior.Color)
                );
        }

        public static System.Data.DataTable ReadExcelFile()
        {
            //string filename = CheckFile();
            string filename = "D:\\Temp\\2.xlsx";
            logfilePath = Path.GetDirectoryName(filename) + "\\watch.log";
            var watcher = new FileSystemWatcher(Path.GetDirectoryName(filename));

            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;

            watcher.Changed += OnChanged;
            watcher.Created += OnCreated;
            watcher.Deleted += OnDeleted;
            watcher.Renamed += OnRenamed;
            watcher.Error += OnError;

            watcher.Filter = "*.csv";
            watcher.IncludeSubdirectories = true;
            watcher.EnableRaisingEvents = true;

            
            try
            {
                if (filename != null)
                {
                    
                    string opfilename = filename.Substring(0, (filename.IndexOf(".xls"))) + ".csv";
                    writer = new StreamWriter(opfilename);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nAn exception has occured.");
                Console.WriteLine(e.ToString());
            }

            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filename);

            Worksheet template = (Worksheet)workbook.Sheets["2021"];

            string s = "";

            for (uint LoopNumberX = 0; LoopNumberX < template.UsedRange.Columns.Count; LoopNumberX++)
            {
                s = "";
                uint color = 0;
                for (uint LoopNumberY = 0; LoopNumberY < template.UsedRange.Rows.Count; LoopNumberY++)
                {
                    writer.AutoFlush = true;
                    System.Drawing.Color SystemDrawingColor = GetBackgroundColor(workbook, "2021", LoopNumberX, LoopNumberY);
                    
                    Console.WriteLine("Location: (" + LoopNumberX.ToString() + ", " + LoopNumberY.ToString() + ")\tColor=" + SystemDrawingColor.ToString());
                    if(SystemDrawingColor == Color.White)
                    {
                        if(color < 1)
                        {
                            color = 1;
                        }
                        
                    }
                    if (SystemDrawingColor == Color.Lime)
                    {
                        if (color < 2) {
                            color = 2;
                        }
                    }
                    Range range = template.Cells[LoopNumberX + 1, LoopNumberY + 1];
                    string text = range.Text.ToString();
                    Console.WriteLine(text);
                    s += text + ",";
                }
                if(color == 1)
                {
                    s += "#,";
                }
                if (color == 2)
                {
                    s += "##,";
                }
                s = s.Substring(0, s.Length - 1);
                //Console.WriteLine(s);
                writer.WriteLine(s);
            }

            Console.WriteLine("\nCSV file has been successfully created.");
            if (writer != null)
                writer.Close();
            return null;

        }

        private static string CheckFile()
        {
            Console.Write("\nEnter \\path\\to\\filename: ");
            string fileName = Console.ReadLine();
            fileName = fileName.Replace(@"\", @"\\");
            fileName = fileName.Replace(@"/", @"\\");
            // Check if file exists and file type is supported
            if (!File.Exists(fileName) || (Path.GetExtension(fileName) != ".xls" && Path.GetExtension(fileName) != ".xlsx"))
            {
                Console.WriteLine("\nInvalid file path or extension.");
                return null;
            }
            else
                return fileName;
        }
        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            if (e.ChangeType != WatcherChangeTypes.Changed)
            {
                return;
            }
            logWriter = new StreamWriter(logfilePath, true);

            logWriter.WriteLine($"{DateTime.UtcNow.ToString()} Changed: {e.FullPath}");

            Console.WriteLine($"{DateTime.UtcNow.ToString()} Changed: {e.FullPath}");

            if (logWriter != null)
                logWriter.Close();
        }

        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            string value = $"{DateTime.UtcNow.ToString()} Created: {e.FullPath}";
            logWriter = new StreamWriter(logfilePath, true);
            logWriter.WriteLine(value);

            Console.WriteLine(value);

            if (logWriter != null)
                logWriter.Close();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {
            logWriter = new StreamWriter(logfilePath, true);
            logWriter.WriteLine($"{DateTime.UtcNow.ToString()} Deleted: {e.FullPath}");

            Console.WriteLine($"{DateTime.UtcNow.ToString()} Deleted: {e.FullPath}");

            if (logWriter != null)
                logWriter.Close();
        }
        private static void OnRenamed(object sender, RenamedEventArgs e)
        {
            logWriter = new StreamWriter(logfilePath, true);
            logWriter.WriteLine($"{DateTime.UtcNow.ToString()} Renamed:");
            logWriter.WriteLine($"    Old: {e.OldFullPath}");
            logWriter.WriteLine($"    New: {e.FullPath}");

            Console.WriteLine($"{DateTime.UtcNow.ToString()} Renamed:");
            Console.WriteLine($"    Old: {e.OldFullPath}");
            Console.WriteLine($"    New: {e.FullPath}");

            if (logWriter != null)
                logWriter.Close();
        }

        private static void OnError(object sender, ErrorEventArgs e) =>
            PrintException(e.GetException());

        private static void PrintException(Exception ex)
        {
            if (ex != null)
            {
                logWriter = new StreamWriter(logfilePath, true);
                logWriter.WriteLine($"{DateTime.UtcNow.ToString()} Message: {ex.Message}");
                logWriter.WriteLine("Stacktrace:");
                logWriter.WriteLine(ex.StackTrace);
                logWriter.WriteLine();

                if (logWriter != null)
                    logWriter.Close();
                Console.WriteLine($"{DateTime.UtcNow.ToString()} Message: {ex.Message}");
                Console.WriteLine("Stacktrace:");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();
                PrintException(ex.InnerException);
            }
        }
    }
}
