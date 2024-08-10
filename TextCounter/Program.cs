using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PdfWordCounter
{
    class Program
    {
        static readonly string SourceDirectory = Path.Combine(Directory.GetCurrentDirectory(), "kaynak");
        static readonly string OutputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "çıktı");

        static void Main(string[] args)
        {
            try
            {
                SetupDirectories();

                Console.WriteLine("SB PDF Kelime sayacına hoş geldiniz...");
                Console.WriteLine();
                Console.WriteLine("//////////////////////////////////////////////");
                Console.WriteLine();
                Console.WriteLine("Klasör içinde bulunan kaynak klasörüne PDF dosyanızı ekleyip programı tekrar çalıştırınız.");
                Console.WriteLine();
                Console.WriteLine("Lütfen PDF dosyanızın adını (uzantısız) giriniz:");
                Console.WriteLine();
                string pdfFileName = Console.ReadLine();
                string pdfPath = GetFilePath(SourceDirectory, pdfFileName, "pdf");
                if (pdfPath == null) return;

                Console.WriteLine("Lütfen Excel dosyanızın adını (uzantısız) giriniz:");
                Console.WriteLine();
                string excelFileName = Console.ReadLine();
                string excelPath = Path.Combine(OutputDirectory, excelFileName + ".xlsx");

                Console.WriteLine("\nİşlem başlıyor...");
                string text = ExtractTextFromPdf(pdfPath);
                Console.WriteLine("PDF başarıyla açıldı ve içindeki tüm metinler alındı.");

                Console.WriteLine("\nMetinler parçalanıp sayıma başlanıyor...");
                var wordFrequencies = CountWordFrequencies(text);
                Console.WriteLine("Metinler başarıyla parçalandı.");

                Console.WriteLine("\nExcel dosyası oluşturuluyor...");
                ExportToExcel(wordFrequencies, excelPath);
                Console.WriteLine($"Excel Dosyası '{excelFileName}.xlsx' adıyla 'çıktı' klasörüne oluşturuldu.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Bir hata oluştu: {ex.Message}");
            }
            finally
            {
                Console.WriteLine("Devam etmek için bir tuşa basın...");
                Console.ReadKey();
            }
        }

        static void SetupDirectories()
        {
            Directory.CreateDirectory(SourceDirectory);
            Directory.CreateDirectory(OutputDirectory);
        }

        static string GetFilePath(string directory, string fileName, string extension)
        {
            string filePath = Path.Combine(directory, $"{fileName}.{extension}");
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Hata: Belirtilen {extension.ToUpper()} dosyası bulunamadı.");
                return null;
            }
            return filePath;
        }

        static string ExtractTextFromPdf(string pdfPath)
        {
            try
            {
                using (var pdf = PdfDocument.Open(pdfPath))
                {
                    return pdf.GetPages().Aggregate("", (text, page) => text + page.Text);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("PDF dosyası okunurken bir hata oluştu.", ex);
            }
        }

        static Dictionary<string, int> CountWordFrequencies(string text)
        {
            var words = Regex.Matches(text.ToLower(), @"\b[\wçğıöşüâ]+\b")
                             .Cast<Match>()
                             .Select(m => m.Value);

            return words.GroupBy(w => w)
                        .ToDictionary(g => g.Key, g => g.Count());
        }

        static void ExportToExcel(Dictionary<string, int> wordFrequencies, string excelPath)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Kelime Sayımları");

                    worksheet.Cell(1, 1).Value = "Kelime";
                    worksheet.Cell(1, 2).Value = "Sayı";

                    int row = 2;
                    foreach (var word in wordFrequencies.OrderByDescending(w => w.Value))
                    {
                        worksheet.Cell(row, 1).Value = word.Key;
                        worksheet.Cell(row, 2).Value = word.Value;
                        row++;
                    }

                    workbook.SaveAs(excelPath);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Excel dosyası oluşturulurken bir hata oluştu.", ex);
            }
        }
    }
}
