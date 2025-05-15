using System.Data;
using System.Globalization;
using CsvHelper;
using ExcelDataReader;

namespace FileRowCounterApp
{

    public interface IFileRowCounter
    {
        bool CanHandle(string extension);
        int CountRows(string filePath);
    }

    public class CsvRowCounter : IFileRowCounter
    {
        public bool CanHandle(string extension) => extension.Equals(".csv", StringComparison.OrdinalIgnoreCase);

        public int CountRows(string filePath)
        {
            using var reader = new StreamReader(filePath);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            var records = csv.GetRecords<dynamic>().ToList();
            return records.Count;
        }
    }

    public class ExcelRowCounter : IFileRowCounter
    {
        public bool CanHandle(string extension)
            => extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase);

        public int CountRows(string filePath)
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            var table = result.Tables[0];
            return table.Rows.Count - 1; // Exclude header
        }
    }

    public class FileProcessor
    {
        private readonly List<IFileRowCounter> _counters;

        public FileProcessor()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            _counters = new List<IFileRowCounter>
            {
                new CsvRowCounter(),
                new ExcelRowCounter()
            };
        }

        public void ProcessFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("Folder not found.");
                return;
            }

            var files = Directory.GetFiles(folderPath, "*.*")
                .Where(f => _counters.Any(c => c.CanHandle(Path.GetExtension(f))))
                .ToList();

            Console.WriteLine("{0,-40} | {1,10}", "File", "Count");
            Console.WriteLine(new string('-', 55));

            foreach (var file in files)
            {
                string fileName = Path.GetFileName(file);
                string extension = Path.GetExtension(file);

                try
                {
                    var counter = _counters.FirstOrDefault(c => c.CanHandle(extension));
                    if (counter != null)
                    {
                        int count = counter.CountRows(file);
                        Console.WriteLine("{0,-40} | {1,10}", fileName, count);
                    }
                    else
                    {
                        Console.WriteLine("{0,-40} | {1,10}", fileName, "N/A");
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("{0,-40} | {1,10}", fileName, "ERROR");
                }
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Set Folder location
            string folderPath = @"C:\Temp\LEAP";
            var processor = new FileProcessor();
            processor.ProcessFolder(folderPath);
        }
    }
}
