using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Text.RegularExpressions;
using System.Xml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

if (args.Length < 2)
{
    Console.WriteLine("Usage:");
    Console.WriteLine("  dotnet run -- <student.xlsx> <answer.xlsx> [maxValueFormulaDiffPerSheet]");
    return;
}

var studentPath = args[0];
var answerPath = args[1];
var maxDiffPerSheet = args.Length >= 3 && int.TryParse(args[2], out var parsed) ? parsed : 200;

if (!File.Exists(studentPath))
{
    Console.WriteLine($"Student file not found: {studentPath}");
    return;
}

if (!File.Exists(answerPath))
{
    Console.WriteLine($"Answer file not found: {answerPath}");
    return;
}

using var studentPkg = new ExcelPackage(new FileInfo(studentPath));
using var answerPkg = new ExcelPackage(new FileInfo(answerPath));

Console.WriteLine("=======================================================");
Console.WriteLine($"Student : {studentPath}");
Console.WriteLine($"Answer  : {answerPath}");
Console.WriteLine("=======================================================");

DumpWorkbookSummary("Student", studentPkg);
DumpWorkbookSummary("Answer", answerPkg);

Console.WriteLine();
Console.WriteLine("=============== WORKBOOK DIFF ===============");
DumpMissingSheets(studentPkg, answerPkg);

Console.WriteLine();
Console.WriteLine("=============== SHEET DIFF ===============");

var allSheetNames = studentPkg.Workbook.Worksheets.Select(ws => ws.Name)
    .Concat(answerPkg.Workbook.Worksheets.Select(ws => ws.Name))
    .Distinct(StringComparer.OrdinalIgnoreCase)
    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
    .ToList();

foreach (var sheetName in allSheetNames)
{
    var studentSheet = studentPkg.Workbook.Worksheets[sheetName];
    var answerSheet = answerPkg.Workbook.Worksheets[sheetName];

    Console.WriteLine();
    Console.WriteLine($"--- Sheet: {sheetName} ---");

    if (studentSheet == null)
    {
        Console.WriteLine("  Missing in student workbook.");
        continue;
    }

    if (answerSheet == null)
    {
        Console.WriteLine("  Missing in answer workbook.");
        continue;
    }

    if (IsChartSheet(studentSheet) || IsChartSheet(answerSheet))
    {
        Console.WriteLine("  This is a chart sheet. Skip cell diff.");
        continue;
    }

    DumpStructureDiff(studentSheet, answerSheet);
    DumpChartDiff(studentSheet, answerSheet);
    DumpValueFormulaDiff(studentSheet, answerSheet, maxDiffPerSheet);
}

static void DumpWorkbookSummary(string label, ExcelPackage package)
{
    Console.WriteLine();
    Console.WriteLine($"[{label}] Worksheets: {package.Workbook.Worksheets.Count}");
    foreach (var ws in package.Workbook.Worksheets)
    {
        if (IsChartSheet(ws))
        {
            Console.WriteLine($"  - {ws.Name} | Type=ChartSheet");
            continue;
        }

        var dim = ws.Dimension?.Address ?? "EMPTY";
        var tableCount = ws.Tables.Count;
        var autoFilter = ws.AutoFilter?.Address?.Address ?? "None";
        Console.WriteLine(
            $"  - {ws.Name} | Dim={dim} | Drawings={ws.Drawings.Count} | MergedRanges={ws.MergedCells.Count} | Tables={tableCount} | AutoFilter={autoFilter}");
    }
}

static void DumpMissingSheets(ExcelPackage studentPkg, ExcelPackage answerPkg)
{
    var studentNames = studentPkg.Workbook.Worksheets.Select(ws => ws.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
    var answerNames = answerPkg.Workbook.Worksheets.Select(ws => ws.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);

    var missingInStudent = answerNames.Where(name => !studentNames.Contains(name)).OrderBy(name => name).ToList();
    var missingInAnswer = studentNames.Where(name => !answerNames.Contains(name)).OrderBy(name => name).ToList();

    if (missingInStudent.Count == 0 && missingInAnswer.Count == 0)
    {
        Console.WriteLine("No missing sheets.");
        return;
    }

    if (missingInStudent.Count > 0)
    {
        Console.WriteLine($"Missing in student: {string.Join(", ", missingInStudent)}");
    }

    if (missingInAnswer.Count > 0)
    {
        Console.WriteLine($"Extra in student (missing in answer): {string.Join(", ", missingInAnswer)}");
    }
}

static bool IsChartSheet(ExcelWorksheet? ws)
{
    if (ws == null)
    {
        return false;
    }

    return ws.GetType().Name.Contains("ChartSheet", StringComparison.OrdinalIgnoreCase);
}

static void DumpStructureDiff(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
{
    var studentMerge = studentSheet.MergedCells.Select(x => x ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).ToHashSet(StringComparer.OrdinalIgnoreCase);
    var answerMerge = answerSheet.MergedCells.Select(x => x ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).ToHashSet(StringComparer.OrdinalIgnoreCase);
    DumpSetDiff("Merged ranges", studentMerge, answerMerge, 20);

    var studentHyperlinks = GetHyperlinkCells(studentSheet);
    var answerHyperlinks = GetHyperlinkCells(answerSheet);
    DumpSetDiff("Hyperlink cells", studentHyperlinks, answerHyperlinks, 20);

    var studentAutoFilter = studentSheet.AutoFilter?.Address?.Address ?? string.Empty;
    var answerAutoFilter = answerSheet.AutoFilter?.Address?.Address ?? string.Empty;
    if (!string.Equals(studentAutoFilter, answerAutoFilter, StringComparison.OrdinalIgnoreCase))
    {
        Console.WriteLine($"  AutoFilter DIFF | Student='{studentAutoFilter}' | Answer='{answerAutoFilter}'");
    }

    var studentSortState = GetSortState(studentSheet);
    var answerSortState = GetSortState(answerSheet);
    if (!string.Equals(studentSortState, answerSortState, StringComparison.OrdinalIgnoreCase))
    {
        Console.WriteLine($"  SortState DIFF | Student='{studentSortState}' | Answer='{answerSortState}'");
    }

    var studentTables = studentSheet.Tables.Select(t => $"{t.Name}:{t.Address.Address}").ToHashSet(StringComparer.OrdinalIgnoreCase);
    var answerTables = answerSheet.Tables.Select(t => $"{t.Name}:{t.Address.Address}").ToHashSet(StringComparer.OrdinalIgnoreCase);
    DumpSetDiff("Tables", studentTables, answerTables, 20);
}

static HashSet<string> GetHyperlinkCells(ExcelWorksheet sheet)
{
    var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var addresses = GetValueFormulaAddresses(sheet);
    foreach (var address in addresses)
    {
        var link = sheet.Cells[address].Hyperlink;
        if (link != null)
        {
            result.Add(address);
        }
    }

    return result;
}

static string GetSortState(ExcelWorksheet sheet)
{
    var ns = new XmlNamespaceManager(new NameTable());
    ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    var node = sheet.WorksheetXml.SelectSingleNode("//x:sortState", ns);
    if (node == null)
    {
        return string.Empty;
    }

    return Regex.Replace(node.OuterXml, "\\s+", " ").Trim();
}

static void DumpSetDiff(string label, HashSet<string> studentSet, HashSet<string> answerSet, int maxItems)
{
    var missing = answerSet.Where(x => !studentSet.Contains(x)).OrderBy(x => x).ToList();
    var extra = studentSet.Where(x => !answerSet.Contains(x)).OrderBy(x => x).ToList();

    if (missing.Count == 0 && extra.Count == 0)
    {
        return;
    }

    Console.WriteLine($"  {label} DIFF");
    if (missing.Count > 0)
    {
        Console.WriteLine($"    Missing ({missing.Count}): {string.Join(", ", missing.Take(maxItems))}");
        if (missing.Count > maxItems)
        {
            Console.WriteLine("    ...");
        }
    }

    if (extra.Count > 0)
    {
        Console.WriteLine($"    Extra ({extra.Count}): {string.Join(", ", extra.Take(maxItems))}");
        if (extra.Count > maxItems)
        {
            Console.WriteLine("    ...");
        }
    }
}

static void DumpChartDiff(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
{
    var studentCharts = studentSheet.Drawings.OfType<ExcelChart>().ToList();
    var answerCharts = answerSheet.Drawings.OfType<ExcelChart>().ToList();
    Console.WriteLine($"  Charts student/answer: {studentCharts.Count}/{answerCharts.Count}");

    var max = Math.Max(studentCharts.Count, answerCharts.Count);
    for (var i = 0; i < max; i++)
    {
        var st = i < studentCharts.Count ? studentCharts[i] : null;
        var an = i < answerCharts.Count ? answerCharts[i] : null;

        if (st == null)
        {
            Console.WriteLine($"    [{i + 1}] Missing chart in student. Answer type={an?.ChartType}, name={an?.Name}");
            continue;
        }

        if (an == null)
        {
            Console.WriteLine($"    [{i + 1}] Extra chart in student. Student type={st.ChartType}, name={st.Name}");
            continue;
        }

        var sameType = st.ChartType == an.ChartType;
        var stPos = $"{CellAddress(st.From.Row + 1, st.From.Column + 1)}:{CellAddress(st.To.Row + 1, st.To.Column + 1)}";
        var anPos = $"{CellAddress(an.From.Row + 1, an.From.Column + 1)}:{CellAddress(an.To.Row + 1, an.To.Column + 1)}";
        Console.WriteLine($"    [{i + 1}] Type {(sameType ? "OK" : "DIFF")} | Student={st.ChartType}, Answer={an.ChartType}");
        Console.WriteLine($"         Pos  Student={stPos} | Answer={anPos}");

        var stSeries = st.Series.Count;
        var anSeries = an.Series.Count;
        Console.WriteLine($"         Series student/answer: {stSeries}/{anSeries}");
        for (var s = 0; s < Math.Max(stSeries, anSeries); s++)
        {
            var stSeriesObj = s < stSeries ? st.Series[s] : null;
            var anSeriesObj = s < anSeries ? an.Series[s] : null;
            if (stSeriesObj == null || anSeriesObj == null)
            {
                Console.WriteLine($"           - Series[{s}] missing on one side.");
                continue;
            }

            var stX = stSeriesObj.XSeries?.ToString() ?? string.Empty;
            var stY = stSeriesObj.Series?.ToString() ?? string.Empty;
            var anX = anSeriesObj.XSeries?.ToString() ?? string.Empty;
            var anY = anSeriesObj.Series?.ToString() ?? string.Empty;
            Console.WriteLine($"           - S{s + 1} X {(stX == anX ? "OK" : "DIFF")} | Student={stX} | Answer={anX}");
            Console.WriteLine($"             S{s + 1} Y {(stY == anY ? "OK" : "DIFF")} | Student={stY} | Answer={anY}");
        }
    }
}

static void DumpValueFormulaDiff(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet, int maxDiffPerSheet)
{
    var studentAddresses = GetValueFormulaAddresses(studentSheet);
    var answerAddresses = GetValueFormulaAddresses(answerSheet);

    var allAddresses = studentAddresses
        .Union(answerAddresses, StringComparer.OrdinalIgnoreCase)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(GetSortKey)
        .ToList();

    var diffCount = 0;
    foreach (var address in allAddresses)
    {
        var st = studentSheet.Cells[address];
        var an = answerSheet.Cells[address];

        var stFormula = NormalizeFormula(st.Formula);
        var anFormula = NormalizeFormula(an.Formula);
        var stVal = NormalizeValue(st.Text);
        var anVal = NormalizeValue(an.Text);

        var valueDiff = !string.Equals(stVal, anVal, StringComparison.Ordinal);
        var formulaDiff = !string.Equals(stFormula, anFormula, StringComparison.OrdinalIgnoreCase);
        if (!valueDiff && !formulaDiff)
        {
            continue;
        }

        diffCount++;
        if (diffCount <= maxDiffPerSheet)
        {
            Console.WriteLine($"  DIFF {address}");
            Console.WriteLine($"    Formula: student='{stFormula}' | answer='{anFormula}'");
            Console.WriteLine($"    Value  : student='{stVal}' | answer='{anVal}'");
        }
    }

    Console.WriteLine($"  Value/Formula differing cells: {diffCount}");
    if (diffCount > maxDiffPerSheet)
    {
        Console.WriteLine($"  (Only first {maxDiffPerSheet} shown)");
    }
}

static HashSet<string> GetValueFormulaAddresses(ExcelWorksheet sheet)
{
    var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var ns = new XmlNamespaceManager(new NameTable());
    ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

    var cellNodes = sheet.WorksheetXml.SelectNodes("//x:sheetData/x:row/x:c", ns);
    if (cellNodes == null)
    {
        return addresses;
    }

    foreach (XmlNode node in cellNodes)
    {
        var refAddress = node.Attributes?["r"]?.Value;
        if (string.IsNullOrWhiteSpace(refAddress))
        {
            continue;
        }

        var hasFormula = node.SelectSingleNode("x:f", ns) != null;
        var hasValue = node.SelectSingleNode("x:v", ns) != null || node.SelectSingleNode("x:is", ns) != null;
        if (!hasFormula && !hasValue)
        {
            continue;
        }

        addresses.Add(refAddress.ToUpperInvariant());
    }

    return addresses;
}

static string NormalizeFormula(string? formula)
{
    return (formula ?? string.Empty).Trim();
}

static string NormalizeValue(string? text)
{
    return (text ?? string.Empty).Trim();
}

static (int Row, int Col) GetSortKey(string address)
{
    var clean = address.Replace("$", string.Empty);
    var match = Regex.Match(clean, "^(?<col>[A-Z]+)(?<row>\\d+)$", RegexOptions.IgnoreCase);
    if (!match.Success)
    {
        return (int.MaxValue, int.MaxValue);
    }

    var colText = match.Groups["col"].Value.ToUpperInvariant();
    var rowText = match.Groups["row"].Value;

    var col = 0;
    foreach (var ch in colText)
    {
        col = (col * 26) + (ch - 'A' + 1);
    }

    return int.TryParse(rowText, out var row) ? (row, col) : (int.MaxValue, col);
}

static string CellAddress(int row, int col)
{
    var column = string.Empty;
    var current = col;
    while (current > 0)
    {
        current--;
        column = (char)('A' + (current % 26)) + column;
        current /= 26;
    }

    return $"{column}{row}";
}
