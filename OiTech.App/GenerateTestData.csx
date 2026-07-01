using OfficeOpenXml;
using System.IO;

// Создание тестового Excel-файла для OiTech
ExcelPackage.License.SetNonCommercialPersonal("OiTech");

using var package = new ExcelPackage();
var ws = package.Workbook.Worksheets.Add("Студенты");

// Заголовки
ws.Cells[1, 1].Value = "ФИО";
ws.Cells[1, 2].Value = "Специальность";
ws.Cells[1, 3].Value = "Группа";
ws.Cells[1, 4].Value = "Математика";
ws.Cells[1, 5].Value = "Физика";
ws.Cells[1, 6].Value = "Информатика";
ws.Cells[1, 7].Value = "История";
ws.Cells[1, 8].Value = "Английский язык";

// Данные
string[,] data = {
    { "Иванов Иван Иванович", "Информационные системы", "ИС-21", "5", "4", "5", "4", "5" },
    { "Петрова Мария Сергеевна", "Информационные системы", "ИС-21", "4", "5", "4", "5", "4" },
    { "Сидоров Алексей Павлович", "Программная инженерия", "ПИ-22", "3", "4", "5", "3", "4" },
    { "Козлова Анна Дмитриевна", "Программная инженерия", "ПИ-22", "5", "5", "5", "5", "5" },
    { "Морозов Дмитрий Олегович", "Кибербезопасность", "КБ-23", "4", "3", "4", "4", "3" },
    { "Волкова Елена Андреевна", "Кибербезопасность", "КБ-23", "5", "4", "5", "5", "4" },
    { "Новиков Артём Викторович", "Информационные системы", "ИС-21", "4", "4", "3", "4", "4" },
    { "Фёдорова Ольга Игоревна", "Программная инженерия", "ПИ-22", "5", "5", "4", "5", "5" },
};

for (int r = 0; r < data.GetLength(0); r++)
    for (int c = 0; c < data.GetLength(1); c++)
        ws.Cells[r + 2, c + 1].Value = data[r, c];

// Автоширина
ws.Cells.AutoFitColumns();

var path = Path.Combine(Directory.GetCurrentDirectory(), "test_students.xlsx");
package.SaveAs(new FileInfo(path));

Console.WriteLine($"Файл создан: {path}");
