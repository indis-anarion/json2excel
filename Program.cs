using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;

namespace Json2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string jsonFilePath = args.Length > 0 ? args[0] : "docs/FORD-KS-SingleSheet.json";
                string excelFilePath = args.Length > 1 ? args[1] : "output.xlsx";
                
                Console.WriteLine($"JSON dosyası okunuyor: {jsonFilePath}");
                var jsonContent = File.ReadAllText(jsonFilePath);
                var jsonObject = JObject.Parse(jsonContent);

                using (var workbook = new XLWorkbook())
                {
                    foreach (var property in jsonObject.Properties())
                    {
                        // Her bir JSON dizisi için yeni bir Excel sayfası oluştur
                        var worksheet = workbook.Worksheets.Add(SanitizeSheetName(property.Name));
                        Console.WriteLine($"Sheet oluşturuluyor: {property.Name}");
                        if (property.Value is not JArray dataArray || dataArray.Count == 0) continue;

                        // RowIndex=0 olan satırı başlık olarak bul
                        var headerRow = dataArray.OfType<JObject>().FirstOrDefault(row => row["RowIndex"]?.ToString() == "0");
                        if (headerRow == null) continue;

                        // Başlık satırını al ve hücreleri doldur
                        var headers = headerRow.Properties().Where(p => p.Name != "RowIndex").Select(p => p.Name).ToList();
                        for (int col = 0; col < headers.Count; col++)
                        {
                            worksheet.Cell(1, col + 1).Value = headerRow[headers[col]]?.ToString() ?? headers[col];
                            worksheet.Cell(1, col + 1).Style.Font.Bold = true;
                            worksheet.Cell(1, col + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        }

                        // Veri satırlarını RowIndex'e göre sırala
                        var dataRows = dataArray.OfType<JObject>()
                            .Where(row => row["RowIndex"]?.ToString() != "0")
                            .Select(row => new { Row = row, Index = int.TryParse(row["RowIndex"]?.ToString(), out var idx) ? idx : int.MaxValue })
                            .OrderBy(x => x.Index)
                            .Select(x => x.Row);

                        // Veri satırlarını işleyerek Excel hücrelerine yaz
                        int excelRowIndex = 2;
                        foreach (var dataRow in dataRows)
                        {
                            for (int col = 0; col < headers.Count; col++)
                            {
                                // Hücre değerini ayarla
                                SetCellValue(worksheet.Cell(excelRowIndex, col + 1), dataRow[headers[col]]);
                            }
                            excelRowIndex++;
                        }
                        worksheet.Columns().AdjustToContents();
                    }
                    workbook.SaveAs(excelFilePath);
                    Console.WriteLine($"Excel dosyası oluşturuldu: {excelFilePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hata oluştu: {ex.Message}");
            }
        }

        // JSON veri tipine göre Excel hücresine uygun formatta yaz
        private static void SetCellValue(IXLCell cell, JToken? value)
        {
            if (value is null || value.Type == JTokenType.Null)
            {
                cell.Value = "";
                return;
            }
            switch (value.Type)
            {
                case JTokenType.Integer:
                case JTokenType.Float:
                    var num = value.ToObject<double?>();
                    cell.Value = num.HasValue ? num.Value : value.ToString();
                    break;
                case JTokenType.Date:
                    var date = value.ToObject<DateTime?>();
                    cell.Value = date.HasValue ? date.Value : value.ToString();
                    cell.Style.DateFormat.Format = "dd/MM/yyyy HH:mm:ss";
                    break;
                case JTokenType.Boolean:
                    var b = value.ToObject<bool?>();
                    cell.Value = b.HasValue ? b.Value : value.ToString();
                    break;
                default:
                    cell.Value = value.ToString();
                    break;
            }
        }

        // Excel sheet isimlerinde geçersiz karakterleri temizle
        private static string SanitizeSheetName(string name)
        {
            var invalidChars = new[] { '\\', '/', '?', '*', '[', ']', ':' };
            var sanitized = invalidChars.Aggregate(name, (current, c) => current.Replace(c, '_'));
            return sanitized.Length > 31 ? sanitized.Substring(0, 31) : sanitized;
        }
    }
}
