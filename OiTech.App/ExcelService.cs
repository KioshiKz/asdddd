using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OiTech.App
{
    public class ExcelService
    {
        public List<Student> Import(string path)
        {
            return ImportStudents(path);
        }

        public List<Student> ImportStudents(string path)
        {
            ExcelPackage.License.SetNonCommercialPersonal("OiTech Diploma System");

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new Exception("Файл Excel не найден.");

            var allStudents = new List<Student>();

            using var package = new ExcelPackage(new FileInfo(path));

            if (package.Workbook.Worksheets.Count == 0)
                throw new Exception("Файл Excel не содержит листов.");

            foreach (var ws in package.Workbook.Worksheets)
            {
                if (ws.Dimension == null)
                    continue;

                var students = ParseWorksheet(ws, path, allStudents.Count);
                allStudents.AddRange(students);
            }

            if (allStudents.Count == 0)
                throw new Exception("Студенты не найдены. Проверьте формат Excel.");

            AppData.Students = allStudents;
            AppData.StudentsCount = allStudents.Count;

            AppData.SubjectNames.Clear();
            foreach (var name in allStudents
                         .SelectMany(s => s.Grades)
                         .Select(g => string.IsNullOrWhiteSpace(g.SubjectKaz) ? g.Subject : g.SubjectKaz)
                         .Where(x => !string.IsNullOrWhiteSpace(x))
                         .Distinct())
            {
                AppData.SubjectNames.Add(name);
            }

            return allStudents;
        }

        private List<Student> ParseWorksheet(ExcelWorksheet ws, string path, int numberOffset)
        {
            var students = new List<Student>();

            int totalRows = ws.Dimension.Rows;
            int totalCols = ws.Dimension.Columns;

            int headerRow = DetectHeaderRow(ws);
            if (headerRow <= 0)
                headerRow = 11;

            int fioCol = DetectFioColumn(ws, headerRow);
            if (fioCol <= 0)
                fioCol = 1;

            int firstStudentRow = DetectFirstStudentRow(ws, headerRow, fioCol);

            string speciality = DetectSpeciality(ws, headerRow);
            string group = DetectGroup(ws, headerRow, path);
            string curator = DetectCurator(ws, group);

            System.Diagnostics.Debug.WriteLine($"========== ПАРСИНГ ЛИСТА: {ws.Name} ==========");
            System.Diagnostics.Debug.WriteLine($"HeaderRow: {headerRow}, FioCol: {fioCol}, FirstStudentRow: {firstStudentRow}");
            System.Diagnostics.Debug.WriteLine($"Group: {group}, Speciality: {speciality}");

            for (int r = firstStudentRow; r <= totalRows; r++)
            {
                string fio = CleanStudentFio(Cell(ws, r, fioCol));

                if (string.IsNullOrWhiteSpace(fio))
                    continue;

                if (IsServiceRow(fio))
                    continue;

                int num = 0;
                if (fioCol > 1)
                    int.TryParse(Cell(ws, r, fioCol - 1), out num);

                if (num <= 0)
                    int.TryParse(Cell(ws, r, 1), out num);

                var student = new Student
                {
                    Number = num > 0 ? num : numberOffset + students.Count + 1,
                    FIO = fio,
                    Speciality = speciality,
                    Group = group,
                    Curator = curator,
                    Grades = new List<GradeEntry>()
                };

                FillGrades(ws, headerRow, r, fioCol, student, path);

                if (student.Grades.Count > 0)
                    students.Add(student);
            }

            return students;
        }

        private void FillGrades(ExcelWorksheet ws, int headerRow, int studentRow, int fioCol, Student student, string path)
        {
            bool isRussian = ws.Name.ToLower().Contains("rus")
              || Path.GetFileNameWithoutExtension(path)
                     .ToLower()
                     .Contains("rus");

            if (ws.Dimension == null)
                return;

            int totalCols = ws.Dimension.Columns;
            var processedSubjects = new HashSet<string>();

            for (int c = fioCol + 1; c <= totalCols; c++)
            {
                string subjectCode = FindSubjectCode(ws, headerRow, c);
                string subjectKaz = FindSubjectNameKaz(ws, headerRow, c);
                string subjectRus = FindSubjectNameRus(ws, headerRow, c);

                string subject = !string.IsNullOrWhiteSpace(subjectKaz)
                    ? subjectKaz
                    : subjectRus;

                // Если это ПМ, но не нашли название - создаём из кода
                if (Regex.IsMatch(subjectCode, @"^ПМ\s*\d+", RegexOptions.IgnoreCase) && string.IsNullOrWhiteSpace(subject))
                {
                    subject = subjectCode;
                }

                // Если название не найдено, но есть код - используем код как название
                if (string.IsNullOrWhiteSpace(subject) && !string.IsNullOrWhiteSpace(subjectCode))
                {
                    subject = subjectCode;
                }

                // Если не нашли название предмета — идём дальше
                if (string.IsNullOrWhiteSpace(subject))
                    continue;

                // Пропускаем служебные заголовки, НО только если это не код модуля
                bool isModuleCode = !string.IsNullOrWhiteSpace(subjectCode) && 
                    Regex.IsMatch(subjectCode, @"^(ПМ|ОН|КМ|ЖММ|ЖБП|ООД|ОГС|ОГСЭ)\s*[\d\.]+", RegexOptions.IgnoreCase);

                // Также проверяем название предмета на наличие кода модуля
                bool isModuleInSubject = Regex.IsMatch(subject, @"^(ПМ|ОН|КМ|ЖММ|ЖБП|ООД|ОГС|ОГСЭ)\s*[\d\.]+", RegexOptions.IgnoreCase);

                if (!isModuleCode && !isModuleInSubject && IsBadSubject(subject))
                {
                    continue;
                }

                // Проверяем, не обработали ли уже этот предмет.
                // Ключ включает саму колонку: в реальных файлах встречаются разные
                // столбцы с дословно одинаковым текстом заголовка (например,
                // скопированные результаты обучения "РО 4.1..." под разными ПМ) —
                // это разные предметы с разными оценками, дедуп по тексту их путал.
                string subjectKey = $"{c}_{(subject ?? "").Trim().ToUpperInvariant()}_{(subjectCode ?? "").Trim().ToUpperInvariant()}";

                if (processedSubjects.Contains(subjectKey))
                {
                    // не перепрыгиваем колонки — просто пропускаем дублирующий заголовок
                    continue;
                }

                // Парсим блок из 4-х колонок с оценками
                GradeEntry? grade = ParseGradeBlock(ws, studentRow, c);

                // Даже если оценок нет, создаём запись (ПМ может быть пустым)
                if (grade == null)
                {
                    grade = new GradeEntry();
                }

                string hours = FindHours(ws, headerRow, c);
                string credits = FindCredits(ws, headerRow, c);

                grade.Subject = subject ?? "";
                grade.SubjectKaz = subjectKaz ?? "";
                grade.SubjectRus = subjectRus ?? "";
                grade.SubjectCode = subjectCode ?? "";
                grade.ParentCode = GetParentCode(subjectCode);
                grade.RowType = GetSubjectRowType(subjectCode);
                grade.Level = GetSubjectLevel(subjectCode);
                grade.Hours = hours ?? "";
                grade.Credits = credits ?? "";

                student.Grades.Add(grade);
                processedSubjects.Add(subjectKey);

                // Перепрыгиваем 3 колонки с оценками (процент, GPA, традиционная)
                c += 3;
            }
        }

        private GradeEntry? ParseGradeBlock(ExcelWorksheet ws, int studentRow, int startCol)
        {
            string letter = CleanCellText(Cell(ws, studentRow, startCol));
            string percentage = CleanCellText(Cell(ws, studentRow, startCol + 1));
            string gpa = CleanCellText(Cell(ws, studentRow, startCol + 2));
            string traditional = CleanCellText(Cell(ws, studentRow, startCol + 3));

            if (string.IsNullOrWhiteSpace(letter) &&
                string.IsNullOrWhiteSpace(gpa) &&
                string.IsNullOrWhiteSpace(percentage) &&
                string.IsNullOrWhiteSpace(traditional))
            {
                return null;
            }

            return new GradeEntry
            {
                LetterGrade = letter ?? "",
                GPA = gpa ?? "",
                Percentage = percentage ?? "",
                TraditionalGrade = traditional ?? ""
            };
        }

        private string FindSubjectCode(ExcelWorksheet ws, int headerRow, int col)
        {
            // Ищем код на 1 строку выше или в той же строке
            string code1 = CleanCellText(Cell(ws, headerRow - 1, col));
            if (LooksLikeSubjectCode(code1)) return code1;

            string code2 = CleanCellText(Cell(ws, headerRow, col));
            if (LooksLikeSubjectCode(code2)) return code2;

            // Также проверяем строку выше на 2 позиции
            string code3 = CleanCellText(Cell(ws, headerRow - 2, col));
            if (LooksLikeSubjectCode(code3)) return code3;

            return "";
        }
   
        private bool LooksLikeSubjectCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return false;

            code = code.Trim().Replace(",", ".");

            // 1. Проверяет чисто числовые коды: "1", "1.2", "1.2.3"
            if (Regex.IsMatch(code, @"^\d+(\.\d+)*$"))
                return true;

            // 2. Распознает стандарты типа "ООД.01", "ПМ.02", "ОГСЭ.04"
            if (Regex.IsMatch(code, @"^[А-Яа-яA-Za-zӘІҢҒҮҰҚӨҺӨЁәіңғүұқөһөё]+\s*[\.\-]?\s*\d+(\.\d+)*$", RegexOptions.IgnoreCase))
                return true;

            // 3. Распознаёт "ПМ 1", "ПМ 01" и т.д. (с пробелом).
            // Обязательно с "$" в конце: без него это совпадало с началом ЛЮБОЙ
            // строки вида "ПМ 2.Ведение бухгалтерского учета..." — то есть с целым
            // предложением-названием модуля, а не только с коротким кодом. Из-за
            // этого такие строки принимались за код с точкой внутри и получали
            // отступ вложенной строки, хотя это обычные предметы верхнего уровня.
            if (Regex.IsMatch(code, @"^ПМ\s+\d+(\.\d+)*$", RegexOptions.IgnoreCase))
                return true;

            return false;
        }

        private string GetSubjectRowType(string? code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return "SUBJECT";

            code = code.Trim().Replace(",", ".");
            if (code.Contains("."))
                return "PART";

            return "SUBJECT";
        }

        private string GetParentCode(string? code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return "";

            code = code.Trim().Replace(",", ".");

            int index = code.LastIndexOf('.');

            if (index <= 0)
                return "";

            return code.Substring(0, index);
        }

        private int GetSubjectLevel(string? code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return 0;

            code = code.Trim().Replace(",", ".");
            return code.Count(ch => ch == '.');
        }

        private string FindSubjectNameKaz(ExcelWorksheet ws, int headerRow, int col)
        {
            string[] candidates =
            {
                Cell(ws, headerRow, col),
                Cell(ws, headerRow + 1, col),
                Cell(ws, headerRow - 1, col),
                Cell(ws, 11, col),
                Cell(ws, 10, col),
                Cell(ws, 9, col),
                Cell(ws, 8, col)
            };

            foreach (var item in candidates)
            {
                string text = CleanCellText(item);

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                // Пропускаем числа (часы, кредиты и т.д.)
                if (Regex.IsMatch(text, @"^\d+([,.]\d+)?$"))
                    continue;

                // Если это "ПМ 1" или подобное - не пропускаем!
                if (Regex.IsMatch(text, @"^ПМ\s*\d+", RegexOptions.IgnoreCase))
                    return text;

                // Если это не служебный заголовок
                if (IsBadSubject(text))
                    continue;

                // Возвращаем любой непустой текст, включая коды типа ПМ.01
                return text;
            }

            return "";
        }

        private string FindSubjectNameRus(ExcelWorksheet ws, int headerRow, int col)
        {
            string[] candidates =
            {
                Cell(ws, headerRow, col),
                Cell(ws, headerRow + 1, col),
                Cell(ws, headerRow - 1, col),
                Cell(ws, 11, col),
                Cell(ws, 10, col),
                Cell(ws, 9, col),
                Cell(ws, 8, col)
            };

            foreach (var item in candidates)
            {
                string text = CleanCellText(item);

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (Regex.IsMatch(text, @"^\d+([,.]\d+)?$"))
                    continue;

                // Ð•ÑÐ»Ð¸ ÑÑ‚Ð¾ "ÐŸÐœ 1" Ð¸Ð»Ð¸ Ð¿Ð¾Ð´Ð¾Ð±Ð½Ð¾Ðµ - Ð½Ðµ Ð¿Ñ€Ð¾Ð¿ÑƒÑÐºÐ°ÐµÐ¼!
                if (Regex.IsMatch(text, @"^ÐŸÐœ\s*\d+", RegexOptions.IgnoreCase))
                    return text;

                // Ð•ÑÐ»Ð¸ ÑÑ‚Ð¾ Ð½Ðµ ÑÐ»ÑƒÐ¶ÐµÐ±Ð½Ñ‹Ð¹ Ð·Ð°Ð³Ð¾Ð»Ð¾Ð²Ð¾Ðº
                if (IsBadSubject(text))
                    continue;

                // Ð’Ð¾Ð·Ð²Ñ€Ð°Ñ‰Ð°ÐµÐ¼ Ð»ÑŽÐ±Ð¾Ð¹ Ð½ÐµÐ¿ÑƒÑÑ‚Ð¾Ð¹ Ñ‚ÐµÐºÑÑ‚, Ð²ÐºÐ»ÑŽÑ‡Ð°Ñ ÐºÐ¾Ð´Ñ‹ Ñ‚Ð¸Ð¿Ð° ÐŸÐœ.01
                return text;
            }

            return "";
        }

        private string FindHours(ExcelWorksheet ws, int headerRow, int col)
        {
            // Ищем в нескольких строках
            string[] candidates =
            {
                Cell(ws, headerRow + 1, col),
                Cell(ws, headerRow + 2, col),
                Cell(ws, headerRow + 3, col),
                Cell(ws, headerRow - 1, col)
            };

            foreach (var item in candidates)
            {
                string text = CleanCellText(item);
                if (Regex.IsMatch(text, @"^\d+([,.]\d+)?$"))
                    return text;
            }

            return "";
        }

        private string FindCredits(ExcelWorksheet ws, int headerRow, int col)
        {
            // Ищем в нескольких строках
            string[] candidates =
            {
                Cell(ws, headerRow + 2, col),
                Cell(ws, headerRow + 3, col),
                Cell(ws, headerRow + 1, col),
                Cell(ws, headerRow - 1, col)
            };

            foreach (var item in candidates)
            {
                string text = CleanCellText(item);
                if (Regex.IsMatch(text, @"^\d+([,.]\d+)?$"))
                    return text;
            }

            return "";
        }

        private int DetectHeaderRow(ExcelWorksheet ws)
        {
            if (ws.Dimension == null)
                return 11;

            int maxRows = Math.Min(ws.Dimension.Rows, 60);
            int maxCols = ws.Dimension.Columns;

            for (int r = 1; r <= maxRows - 4; r++)
            {
                for (int c = 1; c <= Math.Min(10, maxCols); c++)
                {
                    string t = Cell(ws, r, c).ToLower();

                    if (t.Contains("фио") ||
                        t.Contains("таж") ||
                        t.Contains("аты-жөні") ||
                        t.Contains("аты жөні"))
                    {
                        return r;
                    }
                }
            }

            return 11;
        }

        private int DetectFioColumn(ExcelWorksheet ws, int headerRow)
        {
            if (ws.Dimension == null) return 1;

            for (int r = 1; r <= Math.Min(headerRow + 1, ws.Dimension.Rows); r++)
            {
                for (int c = 1; c <= Math.Min(10, ws.Dimension.Columns); c++)
                {
                    string text = Cell(ws, r, c).ToLower();
                    if (text.Contains("фио") || text.Contains("аты-жөні") || text.Contains("аты жөні") ||
                        text.Contains("тегі") || text.Contains("таж") || text.Contains("студент"))
                        return c;
                }
            }
            return 1;
        }

        private int DetectFirstStudentRow(ExcelWorksheet ws, int headerRow, int fioCol)
        {
            int start = headerRow + 3;

            for (int r = headerRow + 1; r <= Math.Min(headerRow + 10, ws.Dimension.Rows); r++)
            {
                string fio = Cell(ws, r, fioCol);
                if (IsProbableFio(fio)) return r;
            }

            return start;
        }

        private bool IsProbableFio(string text)
        {
            text = CleanStudentFio(text);
            if (string.IsNullOrWhiteSpace(text) || IsServiceRow(text)) return false;
            return HasCyrillic(text) && text.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length >= 2;
        }

        private string DetectSpeciality(ExcelWorksheet ws, int headerRow)
        {
            if (ws.Dimension == null) return "";

            for (int r = 1; r <= Math.Min(headerRow, ws.Dimension.Rows); r++)
            {
                for (int c = 1; c <= Math.Min(ws.Dimension.Columns, 20); c++)
                {
                    string text = Cell(ws, r, c);
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    string lower = text.ToLower();
                    if (lower.Contains("мамандық") || lower.Contains("специальность") ||
                        lower.Contains("біліктілік") || lower.Contains("квалификация") ||
                        Regex.IsMatch(text, @"^\d{5,}"))
                    {
                        return CleanCellText(text);
                    }
                }
            }
            return "";
        }

        private string DetectGroup(ExcelWorksheet ws, int headerRow, string path)
        {
            if (ws.Dimension != null)
            {
                for (int r = 1; r <= Math.Min(headerRow + 5, ws.Dimension.Rows); r++)
                {
                    for (int c = 1; c <= Math.Min(ws.Dimension.Columns, 20); c++)
                    {
                        string text = Cell(ws, r, c);
                        Match m = Regex.Match(text, @"\b\d{2}\s*[А-ЯA-ZӘІҢҒҮҰҚӨҺӨЁ]{1,5}\s*[-–]?\s*\d{2}\s*[ккрр]?\b", RegexOptions.IgnoreCase);
                        if (m.Success) return m.Value.Trim();
                    }
                }
            }

            string file = Path.GetFileNameWithoutExtension(path);
            Match fromFile = Regex.Match(file, @"\b\d{2}\s*[А-ЯA-ZӘІҢҒҮҰҚӨҺӨЁ]{1,5}\s*[-–]?\s*\d{2}\s*[ккрр]?\b", RegexOptions.IgnoreCase);
            if (fromFile.Success) return fromFile.Value.Trim();

            return "";
        }

        private string DetectCurator(ExcelWorksheet ws, string group)
        {
            if (ws.Dimension == null) return "";

            int totalRows = ws.Dimension.Rows;
            int totalCols = ws.Dimension.Columns;

            string fromRow61 = DetectCuratorInRow(ws, 61, totalCols, group);
            if (!string.IsNullOrWhiteSpace(fromRow61)) return fromRow61;

            for (int r = 58; r <= Math.Min(65, totalRows); r++)
            {
                string found = DetectCuratorInRow(ws, r, totalCols, group);
                if (!string.IsNullOrWhiteSpace(found)) return found;
            }

            for (int r = 1; r <= totalRows; r++)
            {
                string found = DetectCuratorInRow(ws, r, totalCols, group);
                if (!string.IsNullOrWhiteSpace(found)) return found;
            }

            return "";
        }

        private string DetectCuratorInRow(ExcelWorksheet ws, int row, int totalCols, string group)
        {
            if (row <= 0 || ws.Dimension == null || row > ws.Dimension.Rows) return "";

            for (int c = 1; c <= totalCols; c++)
            {
                string text = Cell(ws, row, c);
                if (string.IsNullOrWhiteSpace(text)) continue;

                string lower = text.ToLower();
                if (!lower.Contains("куратор") && !lower.Contains("жетекші") && !lower.Contains("жетекшісі")) continue;

                string curator = CleanCuratorText(text);
                if (!string.IsNullOrWhiteSpace(curator)) return curator;
            }
            return "";
        }

        private string CleanCuratorText(string text)
        {
            string result = CleanCellText(text);
            result = Regex.Replace(result, @"(?i)^.*?куратор\s+группы[:\-\s]*", "");
            result = Regex.Replace(result, @"(?i)^.*?куратор[:\-\s]*", "");
            result = Regex.Replace(result, @"(?i)^.*?оқу\s+тобының\s+жетекшісі[:\-\s]*", "");
            result = Regex.Replace(result, @"(?i)^.*?жетекшісі[:\-\s]*", "");
            result = Regex.Replace(result, @"(?i)^.*?жетекші[:\-\s]*", "");
            result = Regex.Replace(result, @"^\d{2}\s*[А-ЯA-ZӘІҢҒҮҰҚӨҺӨЁ]{1,5}\s*[-–]?\s*\d{2}\s*[ккрр]?\s*", "", RegexOptions.IgnoreCase);
            return result.Trim(' ', ':', '-', '–');
        }

        private bool IsServiceRow(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return true;

            string lower = text.ToLower();
            return lower.Contains("куратор") || lower.Contains("классный") || lower.Contains("аудитор") ||
                   lower.Contains("итого") || lower.Contains("барлығы") || lower.Contains("всего") ||
                   lower.Contains("подпись") || lower.Contains("заместитель") || lower.Contains("руководитель") ||
                   lower.Contains("м.п") || lower.Contains("м.о");
        }

        private bool IsBadSubject(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return true;

            string lower = text.ToLower();

            // Проверяем, является ли текст названием ПМ
            if (Regex.IsMatch(text, @"^ПМ\s*\d+", RegexOptions.IgnoreCase))
                return false; // НЕ пропускаем ПМ

            // Проверяем, является ли это кодом модуля (ОН, КМ, ООД и т.д.)
            if (Regex.IsMatch(text, @"^(ПМ|ОН|КМ|ЖММ|ЖБП|ООД|ОГС|ОГСЭ)\s*[\d\.]+", RegexOptions.IgnoreCase))
                return false; // НЕ пропускаем коды модулей

            // Границы слова (\b) обязательны: без них, например, "час" ловит
            // "Участвовать", "оценка" — "оценивать", а это реальные названия
            // предметов/результатов обучения, а не служебные заголовки колонок
            // (что и приводило к пропаданию таких столбцов из ведомости).
            if (Regex.IsMatch(lower, @"\bфио\b") || Regex.IsMatch(lower, @"\bтаж\b") ||
                lower.Contains("аты-жөні") || lower.Contains("аты жөні") ||
                Regex.IsMatch(lower, @"\bпайыз") || Regex.IsMatch(lower, @"\bәріп") ||
                Regex.IsMatch(lower, @"\bтрадиц") || Regex.IsMatch(lower, @"\bсандық") ||
                Regex.IsMatch(lower, @"\bбалл") || Regex.IsMatch(lower, @"\bбаға\b") ||
                Regex.IsMatch(lower, @"\bоценка\b") || Regex.IsMatch(lower, @"\bкредит(ы|ов|а)?\b") ||
                Regex.IsMatch(lower, @"\bсағат") || Regex.IsMatch(lower, @"\bчас(а|ов|ы)?\b"))
                return true;

            // Проверяем, что это не "ПМ 1", "ПМ.01" и т.д.
            if (Regex.IsMatch(text, @"^ПМ\s*[\d\.]+", RegexOptions.IgnoreCase))
                return false;

            if (Regex.IsMatch(text, @"^\d+\+\d+$"))
                return true;

            // Буквенные оценки не являются предметами
            if (Regex.IsMatch(text.Trim(), @"^(A|A-|A\+|B|B-|B\+|C|C-|C\+|D|F)$", RegexOptions.IgnoreCase))
                return true;

            // Служебные значения
            if (lower.Contains("зачет") ||
                lower.Contains("сынақ") ||
                lower.Contains("цифровой эквивалент") ||
                lower.Contains("цифравой эквивалент"))
                return true;

            return false;
        }

        private string CleanStudentFio(string text)
        {
            text = CleanCellText(text);
            text = Regex.Replace(text, @"^\d+\s*[\.)\-–]\s*", "");
            return text.Trim();
        }

        private string CleanCellText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            text = text.Replace("\r", " ").Replace("\n", " ");
            text = Regex.Replace(text, @"\s+", " ");
            return text.Trim();
        }

        private string Cell(ExcelWorksheet ws, int row, int col)
        {
            if (ws.Dimension == null) return "";
            if (row < 1 || col < 1 || row > ws.Dimension.Rows || col > ws.Dimension.Columns) return "";

            object value = ws.Cells[row, col].Value;
            if (value == null) return "";

            if (value is double d)
            {
                if (Math.Abs(d - Math.Round(d)) < 0.000001)
                    return ((int)Math.Round(d)).ToString();

                // Используем инвариантную культуру для чисел с точкой, а не запятой
                return d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            }

            return CleanCellText(value.ToString() ?? "");
        }

        private static bool HasCyrillic(string text)
        {
            return text.Any(c =>
                (c >= 'А' && c <= 'я') || c == 'ё' || c == 'Ё' ||
                c == 'Қ' || c == 'қ' || c == 'Ү' || c == 'ү' ||
                c == 'Ұ' || c == 'ұ' || c == 'Ә' || c == 'ә' ||
                c == 'І' || c == 'і' || c == 'Ғ' || c == 'ғ' ||
                c == 'Ң' || c == 'ң' || c == 'Һ' || c == 'һ' ||
                c == 'Ө' || c == 'ө');
        }
    }
}
