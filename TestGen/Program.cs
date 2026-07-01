using OfficeOpenXml;

ExcelPackage.License.SetNonCommercialPersonal("OiTech");
Console.OutputEncoding = System.Text.Encoding.UTF8;

var dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
var files = Directory.GetFiles(dir, "*.xlsx").Where(f => !Path.GetFileName(f).StartsWith("~$")).ToArray();

foreach (var file in files)
{
    Console.WriteLine($"FILE: {Path.GetFileName(file)}");
    using var pkg = new ExcelPackage(new FileInfo(file));
    var ws = pkg.Workbook.Worksheets.FirstOrDefault(s => s.Dimension != null);
    if (ws == null) { Console.WriteLine("NO DATA"); continue; }

    int cols = ws.Dimension.Columns;
    Console.WriteLine($"Rows={ws.Dimension.Rows}, Cols={cols}");

    // Show row 6 (subject headers) - ALL columns
    Console.WriteLine("=== ROW 6 (subjects) ===");
    for (int c = 1; c <= cols; c++)
    {
        string v = ws.Cells[6, c].Text ?? "";
        if (!string.IsNullOrWhiteSpace(v))
            Console.WriteLine($"  Col {c}: [{v}]");
    }

    // Show row 10 (first student) - ALL columns
    Console.WriteLine("=== ROW 10 (student 1) ===");
    for (int c = 1; c <= cols; c++)
    {
        string v = ws.Cells[10, c].Text ?? "";
        if (!string.IsNullOrWhiteSpace(v))
            Console.WriteLine($"  Col {c}: [{v}]");
    }

    // Count students (rows with number in col 1)
    int count = 0;
    for (int r = 10; r <= ws.Dimension.Rows; r++)
    {
        string num = ws.Cells[r, 1].Text ?? "";
        string name = ws.Cells[r, 2].Text ?? "";
        if (int.TryParse(num.Trim(), out _) && !string.IsNullOrWhiteSpace(name))
        {
            count++;
            // Show first few students
            if (count <= 3)
                Console.WriteLine($"  Student {count}: #{num.Trim()} [{name.Trim()}] grades_count={(cols-2)/4}");
        }
    }
    Console.WriteLine($"Total students with numbers: {count}");
}
