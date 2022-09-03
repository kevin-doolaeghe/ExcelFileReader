using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelFileReader {
    public class MaterialDelayExcelFile {
        public string Filename { get; }
        public MaterialDelayExcelFile(string filename) {
            Filename = filename;
        }
        public List<Product> ReadData() {
            ExcelFile file = new(Filename);
            object[,] data = file.ReadData("Sheet1");
            Console.WriteLine($"{data.GetLength(0)} records found.");

            List<Product> products = new();
            try {
                for (int x = 2; x <= data.GetLength(0); x++) {
                    if (data[x, 70] != null) {
                        Product p = new();
                        p.Reference = data[x, 4].ToString();
                        p.Name = data[x, 3].ToString();
                        p.Description = data[x, 6].ToString();
                        p.ProductFamily = data[x, 15].ToString();
                        p.ActualDelay = (double) data[x, 9];
                        p.PreviousDelay = (double) data[x, 10];
                        products.Add(p);
                    }
                }
            } catch (Exception e) {
                Console.WriteLine(e);
            }
            return products;
        }
    }
    public class Product {
        public string? Reference { get; set; }
        public string? Name { get; set; }
        public string? Description { get; set; }
        public string? ProductFamily { get; set; }
        public double ActualDelay { get; set; }
        public double PreviousDelay { get; set; }
        public override string ToString() {
            return JsonSerializer.Serialize(this);
        }
    }
}
