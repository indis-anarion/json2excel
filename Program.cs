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
                string jsonFilePath = args.Length > 0 ? args[0] : "docs/entek.json";
                string excelFilePath = args.Length > 1 ? args[1] : "entek.xlsx";

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

                        // İlk veri satırından başlık isimlerini al (property adları olacak)
                        var firstDataRow = dataArray.OfType<JObject>().FirstOrDefault();
                        if (firstDataRow == null) continue;

                        // Tüm property adlarını başlık olarak kullan (RowIndex hariç)
                        var headers = firstDataRow.Properties()
                            .Where(p => p.Name != "RowIndex")
                            .Select(p => p.Name)
                            .ToList();

                        // Başlık satırını doldur (property adları başlık olacak)
                        for (int col = 0; col < headers.Count; col++)
                        {
                            var cell = worksheet.Cell(1, col + 1);
                            cell.Value = headers[col];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.LightSkyBlue;
                        }

                        // Tüm veri satırlarını olduğu gibi yaz 
                        var dataRows = dataArray.OfType<JObject>().ToList();

                        // Veri satırlarını işleyerek Excel hücrelerine yaz
                        int excelRowIndex = 2;
                        foreach (var dataRow in dataRows)
                        {
                            for (int col = 0; col < headers.Count; col++)
                            {
                                var cell = worksheet.Cell(excelRowIndex, col + 1);
                                SetCellValue(cell, dataRow[headers[col]]);
                            }
                            excelRowIndex++;
                        }
                        worksheet.Columns().AdjustToContents();

                        // Zebra tasarımı için tablo oluştur
                        if (dataRows.Count > 0)
                        {
                            var dataRange = worksheet.Range(1, 1, dataRows.Count + 1, headers.Count);
                            var table = dataRange.CreateTable();

                            // Zebra efekti için tablo temasını ayarla
                            table.Theme = XLTableTheme.TableStyleLight15;
                        }
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
