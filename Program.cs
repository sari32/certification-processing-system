using System.Diagnostics;

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


            string templatePath = Path.GetFullPath("Template.docx"); 
            string outputFolder = Path.GetFullPath("Output");

            if (!File.Exists(templatePath))
            {
                Console.WriteLine("Error: Template file not found!");
                return;
            }

            var generator = new DocumentGenerator();

            Console.WriteLine("Starting report generation...");

            foreach (var candidate in candidates)
            {
                if(candidate.FinalScore>=70)
                generator.GenerateReport(candidate, templatePath, outputFolder);
            }

            Console.WriteLine("Done.");

            Process.Start("explorer.exe", outputFolder);

        }
    }
}
