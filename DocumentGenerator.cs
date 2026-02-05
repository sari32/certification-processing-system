using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace CertificationProcessingSystem
{
    public class DocumentGenerator
    {
        public void GenerateReport(Candidate candidate, string templatePath, string outputFolder)
        {
            Application wordApp = new Application();
            Document doc = null;

            try
            {
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                // העתקת התבנית לקובץ זמני ע"מ שלא להרוס את המקור 
                string tempDocPath = Path.Combine(outputFolder, $"{candidate.FullName}_Temp.docx");
                File.Copy(templatePath, tempDocPath, true);

                doc = wordApp.Documents.Open(tempDocPath);

                // הכנת הטקסט המשתנה לפי הציון הסופי 
                string bodyText;

                if (candidate.FinalScore > 90)
                {
                    bodyText = $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {candidate.FinalScore}.\n" +
                           "נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית.";
                }
                else
                {
                    bodyText = "הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.";
                }
                ReplaceMergeField(doc, "FullName", candidate.FullName);
                ReplaceMergeField(doc, "Department", candidate.Department);
                ReplaceMergeField(doc, "Body", bodyText);

                // שמירת המסמך כ-PDF
                string pdfPath = Path.Combine(outputFolder, $"{candidate.FullName}_Report.pdf");
                doc.SaveAs2(pdfPath, WdSaveFormat.wdFormatPDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating report for {candidate.FullName}: {ex.Message}");
            }
            finally
            {
                // סגירה נקייה של וורד
                if (doc != null)
                {
                    doc.Close(false); // סגור בלי לשמור שינויים בקובץ הוורד הזמני
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
                // מחיקת קובץ הוורד הזמני
                string tempDocPath = Path.Combine(outputFolder, $"{candidate.FullName}_Temp.docx");
                if (File.Exists(tempDocPath)) File.Delete(tempDocPath);
            }
        }

            private void ReplaceMergeField(Document doc, string fieldName, string text)
        {
            foreach (Field field in doc.Fields)
            {
                if (field.Code.Text.Contains(fieldName))
                {
                    field.Select();
                    field.Result.Text = text;
                }
            }
        }

    }

}


