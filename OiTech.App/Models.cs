using System;
using System.Collections.Generic;

namespace OiTech.App
{
    public class Student
    {
        public int Number { get; set; }

        public string FIO { get; set; } = "";
        public string WordFIO { get; set; } = "";

        public string Speciality { get; set; } = "";
        public string Group { get; set; } = "";
        public string Curator { get; set; } = "";

        public string SpecialityCode { get; set; } = "";
        public string SpecialityName { get; set; } = "";

        public string QualificationCode { get; set; } = "";
        public string QualificationName { get; set; } = "";

        public List<GradeEntry> Grades { get; set; } = new();
    }

    public class GradeEntry
    {
        public string Subject { get; set; } = "";

        public string SubjectKaz { get; set; } = "";
        public string SubjectRus { get; set; } = "";

        // Код раздела из Excel: 5, 6, 7, 7.1, 7.2
        public string SubjectCode { get; set; } = "";

        // Родительский код: для 7.1 будет 7
        public string ParentCode { get; set; } = "";

        // SUBJECT = основной предмет, PART = вложенный раздел
        public string RowType { get; set; } = "SUBJECT";

        // 0 = основной, 1 = вложенный, 2 = ещё глубже
        public int Level { get; set; } = 0;

        public string Hours { get; set; } = "";

        // Кредиты
        public string Credits { get; set; } = "";

        public string LetterGrade { get; set; } = "";
        public string GPA { get; set; } = "";
        public string Percentage { get; set; } = "";
        public string TraditionalGrade { get; set; } = "";

        // Для совместимости со старым кодом
        public string Grade
        {
            get => TraditionalGrade;
            set => TraditionalGrade = value;
        }
    }

    // ВАЖНО:
    // Этот класс нужен для AppData.cs.
    // Он хранит историю действий: импорт, печать и т.д.
    public class ActionEntry
    {
        public DateTime Time { get; set; } = DateTime.Now;
        public string Text { get; set; } = "";

        public ActionEntry()
        {
        }

        public ActionEntry(string text)
        {
            Time = DateTime.Now;
            Text = text;
        }
    }
}