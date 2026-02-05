using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificationProcessingSystem
{
    public class CsvService
    {
        public static List<Candidate> LoadAndCleanData(string filePath)
        {
            var candidates = new List<Candidate>();
            var lines = File.ReadAllLines(filePath).Skip(1); // דילוג על הכותרת
            foreach (var line in lines)
            {
                var values = line.Split(',');
                var candidate = new Candidate(
                    values[0],
                    values[1],
                    values[2],
                    int.Parse(values[3]),
                    int.Parse(values[4]));
                candidates.Add(candidate);
            }
            //הסרת כפילויות
            var uniqueCandidates = candidates
                .GroupBy(c => new { c.FullName, c.Department })
                .Select(g => g.First())
                .ToList();

            return uniqueCandidates;
        }
    }
}
