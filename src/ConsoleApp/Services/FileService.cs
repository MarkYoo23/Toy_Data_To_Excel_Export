using ConsoleApp.Models;
using CsvHelper;
using System.Globalization;

namespace ConsoleApp.Services
{
    internal class FileService
    {
        public async Task<IEnumerable<Source>> ReadSourceAsync()
        {
            var folder = $"{AppContext.BaseDirectory}\\Resources";
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            var fileName = "source.csv";
            var filePath = $"{folder}\\{fileName}";

            var bytes = await File.ReadAllBytesAsync(filePath);

            using var memoryStream = new MemoryStream();
            await memoryStream.WriteAsync(bytes);
            memoryStream.Position = 0;

            using var streamReader = new StreamReader(memoryStream);
            streamReader.BaseStream.Position = 0;

            using var csvReader = new CsvReader(streamReader, CultureInfo.CurrentCulture);
            var sources = csvReader.GetRecords<Source>().ToArray();

            return sources;
        }

        public async Task<bool> WriteExcelFileAsync(byte[] bytes)
        {
            var folder = $"{AppContext.BaseDirectory}\\Resources";
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            var fileName = $"result_{DateTime.Now:yyyyMMddhhmmss}.xlsx";
            var filePath = $"{folder}\\{fileName}";

            await File.WriteAllBytesAsync(filePath, bytes);
            
            return true;
        }
    }
}
