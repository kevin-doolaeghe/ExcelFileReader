using ExcelFileReader;
using System;

namespace ExcelFileReaderTest {
    public class Program {
        static void Main(string[] args) {
            MaterialDelayExcelFile file = new("File01.xlsx");
            List<Product> products = file.ReadData();

            Console.WriteLine(products.Count + " products found.");
            Product p = products.OrderBy(x => x.ActualDelay).Last();
            Console.WriteLine(p.ToString());
            /*
            foreach (Product product in products) {
                Console.WriteLine(product.ToString());
            }
            */
        }
    }
}
