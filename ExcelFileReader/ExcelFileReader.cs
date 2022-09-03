using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileReader {
    public class ExcelFile {
        public string Filename { get; }
        public ExcelFile(String filename) {
            Filename = filename;
        }
        public object[,] ReadAll(String table) {
            return ReadData(table);
        }
        public object[,] ReadData(string table, string? range = null) {
            var path = string.Format($"{Directory.GetCurrentDirectory()}\\{Filename}");
            Console.WriteLine(path);

            Excel.Application excel = new();
            Excel.Workbook wb = excel.Workbooks.Open(path);
            Excel.Worksheet ws = (Excel.Worksheet) wb.Worksheets[table];

            Excel.Range cells = range == null ? ws.UsedRange : ws.Range[range];
            object[,] arr = (object[,]) cells.Value;

            wb.Close();
            excel.Quit();
            return arr;
        }
    }
}