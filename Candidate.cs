using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificationProcessingSystem
{
    public class Candidate
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Department { get; set; }
        public int TheoryScore { get; set; }
        public int PracticalScore { get; set; }

        public double FinalScore { get; private set; }
        public string FullName => $"{FirstName} {LastName}";

        //constructor
        public Candidate(string firstName, string lastName, string department, int theoryScore, int practicalScore)
        {
            FirstName = firstName;
            LastName = lastName;
            Department = department;
            TheoryScore = theoryScore;
            PracticalScore = practicalScore;

            FormatNames();
            CalculateFinalScore();
        }


        public void FormatNames()
        {
            FirstName = FixCase(FirstName);
            LastName = FixCase(LastName);
        }

        // פונקציית עזר פנימית לתיקון השמות (dana -> Dana)
        private string FixCase(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            text = text.Trim();
            if (text.Length == 1) return text.ToUpper();
            return char.ToUpper(text[0]) + text.Substring(1).ToLower();
        }
        public void CalculateFinalScore()
        {
            FinalScore = (TheoryScore * 0.4) + (PracticalScore * 0.6);
        }

          




    }
}
