# 🎓 Certification Processing System (CPS)

![C#](https://img.shields.io/badge/C%23-239120?style=for-the-badge&logo=c-sharp&logoColor=white)
![.NET](https://img.shields.io/badge/.NET-512BD4?style=for-the-badge&logo=dotnet&logoColor=white)
![Microsoft Word](https://img.shields.io/badge/Microsoft_Word-2B579A?style=for-the-badge&logo=microsoft-word&logoColor=white)

מערכת אוטומטית לעיבוד נתוני הכשרה והפקת מכתבים רשמיים בפורמט PDF. המערכת קוראת נתונים מקובץ CSV, מבצעת ניקוי נתונים (Deduplication), ומפיקה מסמכים מותאמים אישית על בסיס תבנית Word.

## 🚀 תכונות עיקריות (Key Features)

- **Data Cleaning**: הסרת כפילויות אוטומטית לפי מזהה ייחודי (ID).
- **Automated Reporting**: הפקת מכתבים מבוססת תבנית (`.docx`) עם שדות Mail Merge.
- **Dynamic Logic**: 
  - מועמדים עם ציון מתחת ל-70 אינם זכאים למכתב.
  - עבור מועמדים מצטיינים (ציון > 90) המערכת מפיקה מכתב עם ייעוד לתפקיד "מוביל טכנולוגי".
  - עבור מועמדים עם ציון בין 70 ל 90 המערכת מפיקה מכתב עם הודעה שלא נמצא תפקיד מתאים עבורם.
- **PDF Export**: המרה אוטומטית של המכתבים לפורמט PDF .
- **Resource Management**: ניהול זיכרון קפדני וסגירת תהליכי רקע של Office (COM Cleanup).

## 🛠 טכנולוגיות (Tech Stack)

- **Language**: C# (.NET Core/Standard)
- **Library**: `Microsoft.Office.Interop.Word`
- **Source**: CSV Data Parsing
- **Output**: PDF (Portable Document Format)

## 📋 דרישות קדם (Prerequisites)

כדי להריץ את המערכת בהצלחה, יש לוודא:
1. מותקנת חבילת **Microsoft Office** (Word) על המחשב המריץ.
2. קובץ הנתונים נמצא בנתיב: `data/Data.csv`.
3. קובץ התבנית `Template.docx` נמצא בתיקיית המקור (מוגדר כ-`Copy Always`).

## 📁 מבנה הפרויקט (Project Structure)

```text
├── Data/
│   └── Data.csv          # קובץ המקור עם נתוני המתלמדים
├── CsvService.cs         # לוגיקת קריאה וניקוי נתונים
├── DocumentGenerator.cs  # מנוע הפקת המסמכים (Word Interop)
└── Program.cs            # נקודת הכניסה וניהול התהליך
└── Template.docx         # תבנית ה-Word עם שדות המיזוג

## 🚀 שיפורים לייצור (Production Readiness)
**מענה לשאלה: אילו שיפורים הייתי מבצע לפני העלאה לייצור (Production)?**

במידה והמערכת הייתה מיועדת לסביבת Production בעומסים גבוהים, הייתי מבצע את השינויים הארכיטקטוניים הבאים:

### 1. החלפת תשתית העבודה מול Word (Technical Stack)
השימוש הנוכחי ב-`Microsoft.Office.Interop` אינו מומלץ לשרתים (Server-side) כיוון שהוא איטי, דורש התקנת Office, וחשוף לבעיות זיכרון (Memory Leaks).
*   **הפתרון:** מעבר לספריות העובדות ישירות מול ה-XML של הקובץ, כגון **OpenXML SDK** או **DocX**.
*   **היתרון:** ביצועים מהירים פי 100, אי-תלות בהתקנת Word, ויציבות מלאה.

### 2. ולידציה לתבנית (Template Validation)
מכיוון שהמשתמש יכול לערוך את התבנית ידנית, קיים סיכון למחיקת שדות מיזוג (Merge Fields).
*   **הפתרון:** הוספת שלב מקדים הסורק את התבנית ומוודא שכל השדות הנדרשים (`FullName`, `Score`, `Department`) קיימים לפני תחילת העיבוד, כדי למנוע יצירת מסמכים ריקים/שגויים.


### 3. טיפול בשגיאות ולוגים (Error Handling & Logging)
*   שימוש בספריית לוגים (כגון Serilog או NLog) לכתיבת שגיאות לקובץ מסודר ולא רק ל-Console.
*   יצירת מנגנון "Retry" למקרה שהקובץ תפוס ע"י תהליך אחר.
