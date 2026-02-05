namespace CertificationProcessingSystem
{
    internal class Program
    {
        static void Main(string[] args)
        {

            string filePath = "data/Data.csv";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Error: File not found.");
                return;
            }

            var candidates = CsvService.LoadAndCleanData(filePath);

            // הדפסה לבדיקה
            foreach (var c in candidates)
            {
                Console.WriteLine($"Name: {c.FullName}, Dept: {c.Department}, Final Score: {c.FinalScore}");
            }


        }
    }
}
