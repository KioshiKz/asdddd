using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OiTech.App.Services
{
    public class WordService
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // Пример групп:
        // 22ТП -41р
        // 22ТМ -41к
        // 23МК-31р
        // 25ЭБп-32р пр
        // 24 БР-21р
        private static readonly Regex GroupRx = new(
            @"^\d{2}\s*[А-ЯA-ZҚҮҰӘІҒҢҺӨЁ].*[-–]\s*\d{2}",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Строка специальности обычно начинается с кода:
        // 06130100 Программное обеспечение, квалификация 4S06130103 Разработчик...
        // 07161300 Техническое обслуживание..., квалификация 4S07161304 Техник-механик
        private static readonly Regex SpecRx = new(
            @"^\d{5,}",
            RegexOptions.Compiled);

        private static readonly Regex SkipRx = new(
            @"^(Классный|Аудитор|Аудитория|Смена|Всего|Бюджет|Платно|Итого|Барлығы)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Хвосты после ФИО:
        // платник, платно, д-о, эб, МБ, АС, бр, academ и т.д.
        private static readonly Regex SuffixRx = new(
            @"\s+(платни[кц]|платно|д[-/–]о|д-о|эб|[Мм][Бб]|[Аа][Сс]|пл\b|бр\b|academ|академ)\b.*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public List<Student> Import(string filePath)
        {
            var importedStudents = ImportWord(filePath);

            MergeWithAppData(importedStudents);

            return AppData.Students;
        }

        public List<Student> ImportWord(string filePath)
        {
            var lines = ReadParagraphs(filePath);

            var students = new List<Student>();

            string currentGroup = "";

            string currentSpecialityRaw = "";

            string currentSpecialityCode = "";
            string currentSpecialityName = "";

            string currentQualificationCode = "";
            string currentQualificationName = "";

            int number = 1;

            foreach (var raw in lines)
            {
                string text = CleanText(raw);

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (IsGroup(text))
                {
                    currentGroup = StripGroupSuffix(text);

                    currentSpecialityRaw = "";

                    currentSpecialityCode = "";
                    currentSpecialityName = "";

                    currentQualificationCode = "";
                    currentQualificationName = "";

                    continue;
                }

                if (IsSpeciality(text))
                {
                    currentSpecialityRaw = text;

                    var parsed = ParseSpecialityAndQualification(text);

                    currentSpecialityCode = parsed.SpecialityCode;
                    currentSpecialityName = parsed.SpecialityName;

                    currentQualificationCode = parsed.QualificationCode;
                    currentQualificationName = parsed.QualificationName;

                    continue;
                }

                if (SkipRx.IsMatch(text))
                    continue;

                if (string.IsNullOrWhiteSpace(currentGroup))
                    continue;

                string name = RemoveListNumber(text);
                name = SuffixRx.Replace(name, "").Trim();

                if (!IsStudentName(name))
                    continue;

                students.Add(new Student
                {
                    Number = number++,

                    FIO = name,
                    WordFIO = name,

                    Group = currentGroup,

                    Speciality = currentSpecialityRaw,

                    SpecialityCode = currentSpecialityCode,
                    SpecialityName = currentSpecialityName,

                    QualificationCode = currentQualificationCode,
                    QualificationName = currentQualificationName,

                    Grades = new List<GradeEntry>()
                });
            }

            return students;
        }

        private void MergeWithAppData(List<Student> wordStudents)
        {
            if (wordStudents == null || wordStudents.Count == 0)
                return;

            if (AppData.Students == null)
                AppData.Students = new List<Student>();

            // Если Excel ещё не импортировали — просто ставим студентов из Word.
            // Потом Excel можно импортировать отдельно.
            if (AppData.Students.Count == 0)
            {
                AppData.Students = wordStudents;
                AppData.AddAction("Импорт Word");
                return;
            }

            foreach (var excelStudent in AppData.Students)
            {
                Student? wordStudent = FindWordStudentForExcelStudent(excelStudent, wordStudents);

                if (wordStudent == null)
                {
                    // Если по ФИО не нашли, пробуем по группе.
                    wordStudent = wordStudents.FirstOrDefault(w =>
                        NormalizeGroup(w.Group) == NormalizeGroup(excelStudent.Group));
                }

                if (wordStudent == null)
                {
                    // Если группа в Excel пустая, но Word содержит одну группу,
                    // берём первого студента как источник специальности/квалификации.
                    wordStudent = wordStudents.FirstOrDefault();
                }

                if (wordStudent == null)
                    continue;

                // ФИО из Word ставим только если нашли именно этого студента.
                var matchedByName = IsSameStudentName(excelStudent.FIO, wordStudent.FIO);

                if (matchedByName && !string.IsNullOrWhiteSpace(wordStudent.WordFIO))
                    excelStudent.WordFIO = wordStudent.WordFIO;

                // Специальность и квалификация можно ставить по группе.
                if (!string.IsNullOrWhiteSpace(wordStudent.Speciality))
                    excelStudent.Speciality = wordStudent.Speciality;

                if (!string.IsNullOrWhiteSpace(wordStudent.SpecialityCode))
                    excelStudent.SpecialityCode = wordStudent.SpecialityCode;

                if (!string.IsNullOrWhiteSpace(wordStudent.SpecialityName))
                    excelStudent.SpecialityName = wordStudent.SpecialityName;

                if (!string.IsNullOrWhiteSpace(wordStudent.QualificationCode))
                    excelStudent.QualificationCode = wordStudent.QualificationCode;

                if (!string.IsNullOrWhiteSpace(wordStudent.QualificationName))
                    excelStudent.QualificationName = wordStudent.QualificationName;
            }

            AppData.AddAction("Импорт Word");
        }

        private Student? FindWordStudentForExcelStudent(Student excelStudent, List<Student> wordStudents)
        {
            if (excelStudent == null)
                return null;

            string excelName = NormalizeName(excelStudent.FIO);

            if (string.IsNullOrWhiteSpace(excelName))
                return null;

            Student? exact = wordStudents.FirstOrDefault(w =>
                NormalizeName(w.FIO) == excelName ||
                NormalizeName(w.WordFIO) == excelName);

            if (exact != null)
                return exact;

            var excelParts = SplitName(excelStudent.FIO);

            if (excelParts.Length < 2)
                return null;

            foreach (var w in wordStudents)
            {
                var wordParts = SplitName(w.FIO);

                if (wordParts.Length < 2)
                    continue;

                // Фамилия + имя совпали
                if (SameText(excelParts[0], wordParts[0]) &&
                    SameText(excelParts[1], wordParts[1]))
                {
                    return w;
                }
            }

            return null;
        }

        private List<string> ReadParagraphs(string filePath)
        {
            var result = new List<string>();

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return result;

            string tempFile = Path.Combine(
                Path.GetTempPath(),
                "oitech_word_" + Guid.NewGuid().ToString("N") + ".docx");

            try
            {
                // Самое важное место:
                // читаем файл даже если он открыт в Microsoft Word.
                using (var source = new FileStream(
                    filePath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete))
                using (var destination = new FileStream(
                    tempFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.None))
                {
                    source.CopyTo(destination);
                }

                using var zip = ZipFile.OpenRead(tempFile);

                var entry = zip.GetEntry("word/document.xml")
                            ?? throw new Exception("Файл повреждён или не является DOCX.");

                using var stream = entry.Open();
                var doc = XDocument.Load(stream);

                foreach (var para in doc.Descendants(W + "p"))
                {
                    string text = string.Concat(
                        para.Descendants(W + "t").Select(t => t.Value)).Trim();

                    if (!string.IsNullOrWhiteSpace(text))
                        result.Add(CleanText(text));
                }

                return result;
            }
            finally
            {
                try
                {
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                }
                catch
                {
                }
            }
        }

        private bool IsGroup(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = CleanText(text);

            if (text.Length > 40)
                return false;

            return GroupRx.IsMatch(text);
        }

        private bool IsSpeciality(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = CleanText(text);

            if (!SpecRx.IsMatch(text))
                return false;

            return text.Length > 15;
        }

        private string StripGroupSuffix(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = CleanText(text);

            text = Regex.Replace(text, @"\s+пр\s*$", "", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"\s+", " ");

            return text.Trim();
        }

        private SpecialityParseResult ParseSpecialityAndQualification(string text)
        {
            var result = new SpecialityParseResult();

            string clean = CleanText(text);

            // Берём код специальности в начале строки
            Match main = Regex.Match(
                clean,
                @"^(?<specCode>\d{5,}[A-Za-zА-Яа-яӘІҢҒҮҰҚӨҺӨЁәіңғүұқөһөё0-9\-\.]*)\s*[-–]?\s*(?<rest>.+)$",
                RegexOptions.IgnoreCase);

            if (!main.Success)
                return result;

            result.SpecialityCode = CleanCode(main.Groups["specCode"].Value);

            string rest = CleanText(main.Groups["rest"].Value);

            // Вариант:
            // Программное обеспечение, квалификация 4S06130103 Разработчик программного обеспечения
            Match qual = Regex.Match(
                rest,
                @"(?<specName>.*?)(?:,?\s*(?:квалификация|біліктілік)\s+)(?<qualCode>[0-9A-Za-zА-Яа-яӘІҢҒҮҰҚӨҺӨЁәіңғүұқөһөё\-\.]+)?\s*[-–]?\s*(?<qualName>.*)$",
                RegexOptions.IgnoreCase);

            if (qual.Success)
            {
                result.SpecialityName = CleanName(qual.Groups["specName"].Value);
                result.QualificationCode = CleanCode(qual.Groups["qualCode"].Value);
                result.QualificationName = CleanName(qual.Groups["qualName"].Value);
            }
            else
            {
                // Вариант:
                // 04110100 Учет и аудит 4S04110102 Бухгалтер
                Match noWordQualification = Regex.Match(
                    rest,
                    @"(?<specName>.*?)(?<qualCode>\b[0-9][A-Za-zА-Яа-я0-9]{1,3}[0-9]{5,}\b)\s*(?<qualName>.*)$",
                    RegexOptions.IgnoreCase);

                if (noWordQualification.Success)
                {
                    result.SpecialityName = CleanName(noWordQualification.Groups["specName"].Value);
                    result.QualificationCode = CleanCode(noWordQualification.Groups["qualCode"].Value);
                    result.QualificationName = CleanName(noWordQualification.Groups["qualName"].Value);
                }
                else
                {
                    result.SpecialityName = CleanName(rest);
                }
            }

            // Если название квалификации перенеслось на следующую строку,
            // это уже объединено в Word как абзац обычно. Если нет — останется то, что есть.
            return result;
        }

        private string RemoveListNumber(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = CleanText(text);

            // Удаляет "1. ", "25) ", "12 - "
            text = Regex.Replace(text, @"^\d+\s*[\.\)\-–]\s*", "");

            return text.Trim();
        }

        private string CleanText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = text.Replace("\r", " ").Replace("\n", " ");
            text = Regex.Replace(text, @"\s+", " ");

            return text.Trim(' ', ':', '-', '–');
        }

        private string CleanCode(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            string result = CleanText(text);

            result = result.Trim('.', ',', ';', ':', '-', '–');

            return result;
        }

        private string CleanName(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            string result = CleanText(text);

            result = Regex.Replace(result, @"\s+код.*$", "", RegexOptions.IgnoreCase);
            result = Regex.Replace(result, @"\s+квалификация.*$", "", RegexOptions.IgnoreCase);
            result = Regex.Replace(result, @"\s+біліктілік.*$", "", RegexOptions.IgnoreCase);

            result = result.Trim(' ', '.', ',', ';', ':', '-', '–');

            return result;
        }

        private bool IsStudentName(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = RemoveListNumber(CleanText(text));

            if (text.Length < 5)
                return false;

            if (!HasCyrillic(text))
                return false;

            string lower = text.ToLower();

            if (lower.Contains("квалификация")) return false;
            if (lower.Contains("біліктілік")) return false;
            if (lower.Contains("мамандық")) return false;
            if (lower.Contains("специальность")) return false;
            if (lower.Contains("аудитор")) return false;
            if (lower.Contains("аудитория")) return false;
            if (lower.Contains("смена")) return false;
            if (lower.Contains("бюджет")) return false;
            if (lower.Contains("всего")) return false;
            if (lower.Contains("классный")) return false;
            if (lower.Contains("руководитель")) return false;

            var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            return parts.Length >= 2;
        }

        private static bool HasCyrillic(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            return text.Any(c =>
                (c >= 'А' && c <= 'я') || c == 'ё' || c == 'Ё' ||
                c == 'Қ' || c == 'қ' ||
                c == 'Ү' || c == 'ү' ||
                c == 'Ұ' || c == 'ұ' ||
                c == 'Ә' || c == 'ә' ||
                c == 'І' || c == 'і' ||
                c == 'Ғ' || c == 'ғ' ||
                c == 'Ң' || c == 'ң' ||
                c == 'Һ' || c == 'һ' ||
                c == 'Ө' || c == 'ө');
        }

        private string NormalizeName(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = text.ToUpperInvariant();

            text = Regex.Replace(
                text,
                @"[^А-ЯЁӘІҢҒҮҰҚӨҺӨA-Z]",
                "");

            return text;
        }

        private string NormalizeGroup(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            return Regex.Replace(
                text.ToUpperInvariant(),
                @"[\s\-–_]",
                "");
        }

        private string[] SplitName(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return Array.Empty<string>();

            return RemoveListNumber(CleanText(text))
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.ToUpperInvariant())
                .ToArray();
        }

        private bool SameText(string a, string b)
        {
            return NormalizeName(a) == NormalizeName(b);
        }

        private bool IsSameStudentName(string a, string b)
        {
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
                return false;

            if (NormalizeName(a) == NormalizeName(b))
                return true;

            var aParts = SplitName(a);
            var bParts = SplitName(b);

            if (aParts.Length < 2 || bParts.Length < 2)
                return false;

            return SameText(aParts[0], bParts[0]) &&
                   SameText(aParts[1], bParts[1]);
        }

        private sealed class SpecialityParseResult
        {
            public string SpecialityCode { get; set; } = "";

            public string SpecialityName { get; set; } = "";

            public string QualificationCode { get; set; } = "";

            public string QualificationName { get; set; } = "";
        }
    }
}