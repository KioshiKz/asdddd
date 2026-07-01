using System;
using System.Collections.Generic;
using System.Linq;

namespace OiTech.App
{
    public static class AppData
    {
        public static List<Student> Students { get; set; } = new();

        public static int PrintedCount { get; set; }

        public static List<ActionEntry> Actions { get; set; } = new();

        public static int StudentsCount
        {
            get => Students?.Count ?? 0;
            set
            {
                // Для совместимости со старым кодом
            }
        }

        public static List<string> SubjectNames
        {
            get
            {
                if (Students == null || Students.Count == 0)
                    return new List<string>();

                return Students
                    .SelectMany(s => s.Grades ?? new List<GradeEntry>())
                    .Select(g => g.Subject)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct()
                    .ToList();
            }
            set
            {
                // Для совместимости со старым кодом
            }
        }

        public static void AddAction(string text)
        {
            Actions.Add(new ActionEntry
            {
                Time = DateTime.Now,
                Text = text
            });
        }
    }
}