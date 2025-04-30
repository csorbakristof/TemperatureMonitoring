using System.Text.Json;

namespace DataMergerTool
{
    internal class Program
    {
        static void Main(string[] args)
        {
            const string directoryName = @"e:\temp\allTempLogZips";
            var loader = new TempHumLoader();
            var measurements = loader.Load(directoryName);
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            };
            var json = JsonSerializer.Serialize(measurements, options);
            var jsonFileName = Path.Combine(directoryName, "allTempLog.json");
            File.WriteAllText(jsonFileName, json);
            Console.WriteLine("Measurements collected.");
        }
    }
}
