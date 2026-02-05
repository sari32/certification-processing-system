using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

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

            string templatePath = Path.GetFullPath("Template.docx"); 
            string outputFolder = Path.GetFullPath("Output");

            if (!File.Exists(templatePath))
            {
                Console.WriteLine("Error: Template file not found!");
                return;
            }

            var generator = new DocumentGenerator();
            Application wordApp = new Application(); 

            try
            {
                Console.WriteLine("Starting report generation...");
                foreach (var candidate in candidates)
                {
                    if (candidate.FinalScore >= 70)
                    {
                        generator.GenerateReport(candidate, templatePath, outputFolder, wordApp);
                    }
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    wordApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    Console.WriteLine("Word Application closed safely.");
                }
            }

            Console.WriteLine("Done.");
            Process.Start("explorer.exe", outputFolder);

        }
    }
}
