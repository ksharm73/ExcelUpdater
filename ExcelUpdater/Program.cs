using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;

// Exit codes for AutoSys: 0 = success, 1 = config/validation, 2 = file/IO (locked, missing), 3 = unexpected
const int ExitOk = 0, ExitConfig = 1, ExitIo = 2, ExitFailure = 3;

static void Log(string level, string msg) =>
    Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {msg}");

try
{
    var opts = new ConfigurationBuilder()
        .SetBasePath(AppContext.BaseDirectory)
        .AddJsonFile("appsettings.json", optional: false)
        .AddCommandLine(args)               // e.g. --FilePath "\\server\share\other.xlsx"
        .Build()
        .Get<JobOptions>() ?? new JobOptions();

    ValidateOptions(opts);
    Log("INFO", $"Starting. Target: {opts.FilePath}");

    using var fs = OpenExclusive(opts);      // exclusive lock held until the end of the run

    // 1. Read the workbook into memory
    var original = new MemoryStream();
    fs.CopyTo(original);
    original.Position = 0;

    // 2. Backup before touching anything
    Backup(opts, original);
    original.Position = 0;

    // 3. Update cells + recalculate (all in memory, file untouched so far)
    using var workbook = new XLWorkbook(original);
    ApplyUpdates(workbook, opts);
    Recalculate(workbook, opts);
    LogVerifyCells(workbook, opts);

    using var updated = new MemoryStream();
    workbook.SaveAs(updated);

    // 4. Only now write back, to the same locked handle
    fs.SetLength(0);
    fs.Position = 0;
    updated.Position = 0;
    updated.CopyTo(fs);
    fs.Flush(true);

    Log("INFO", "File updated successfully.");
    return ExitOk;
}
catch (ConfigException ex)
{
    Log("ERROR", ex.Message);
    return ExitConfig;
}
catch (IOException ex)
{
    Log("ERROR", $"File access problem: {ex.Message}");
    return ExitIo;
}
catch (UnauthorizedAccessException ex)
{
    Log("ERROR", $"Access denied: {ex.Message}");
    return ExitIo;
}
catch (Exception ex)
{
    Log("ERROR", $"Unexpected failure: {ex}");
    return ExitFailure;
}

// ---------------------------------------------------------------------------

static void ValidateOptions(JobOptions o)
{
    if (string.IsNullOrWhiteSpace(o.FilePath))
        throw new ConfigException("FilePath is required.");
    if (o.Updates.Count == 0)
        throw new ConfigException("At least one entry in Updates is required.");
    foreach (var u in o.Updates)
    {
        if (string.IsNullOrWhiteSpace(u.Sheet) || string.IsNullOrWhiteSpace(u.Cell))
            throw new ConfigException("Every update needs Sheet and Cell.");
    }
}

static FileStream OpenExclusive(JobOptions o)
{
    if (!File.Exists(o.FilePath))
        throw new FileNotFoundException($"File not found: {o.FilePath}");

    for (var attempt = 1; ; attempt++)
    {
        try
        {
            // FileShare.None: fails if a user has the workbook open, so we never clobber their edits
            return new FileStream(o.FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (IOException) when (attempt < o.LockRetryCount)
        {
            Log("WARN", $"File is locked (attempt {attempt}/{o.LockRetryCount}); retrying in {o.LockRetryDelaySeconds}s.");
            Thread.Sleep(TimeSpan.FromSeconds(o.LockRetryDelaySeconds));
        }
    }
}

static void Backup(JobOptions o, MemoryStream content)
{
    if (string.IsNullOrWhiteSpace(o.BackupFolder)) return;

    Directory.CreateDirectory(o.BackupFolder);
    var name = $"{Path.GetFileNameWithoutExtension(o.FilePath)}_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(o.FilePath)}";
    var dest = Path.Combine(o.BackupFolder, name);
    File.WriteAllBytes(dest, content.ToArray());
    Log("INFO", $"Backup written: {dest}");

    if (o.BackupRetentionDays <= 0) return;
    var cutoff = DateTime.Now.AddDays(-o.BackupRetentionDays);
    var prefix = Path.GetFileNameWithoutExtension(o.FilePath) + "_";
    foreach (var old in Directory.EnumerateFiles(o.BackupFolder, prefix + "*" + Path.GetExtension(o.FilePath)))
    {
        if (File.GetLastWriteTime(old) < cutoff)
        {
            try { File.Delete(old); Log("INFO", $"Deleted old backup: {old}"); }
            catch (Exception ex) { Log("WARN", $"Could not delete {old}: {ex.Message}"); }
        }
    }
}

static void ApplyUpdates(XLWorkbook wb, JobOptions o)
{
    foreach (var u in o.Updates)
    {
        if (!wb.TryGetWorksheet(u.Sheet, out var ws))
            throw new ConfigException($"Worksheet '{u.Sheet}' not found.");

        var cell = ws.Cell(u.Cell);
        if (cell.HasFormula)
            throw new ConfigException($"{u.Sheet}!{u.Cell} contains a formula; refusing to overwrite it.");

        var old = cell.Value.ToString();
        var raw = u.Value == "{today}" ? DateTime.Today.ToString("yyyy-MM-dd") : u.Value;

        switch (u.Type.ToLowerInvariant())
        {
            case "number":
                cell.Value = double.Parse(raw, CultureInfo.InvariantCulture);
                break;
            case "date":
                cell.Value = DateTime.Parse(raw, CultureInfo.InvariantCulture);
                break;
            case "bool":
                cell.Value = bool.Parse(raw);
                break;
            case "string":
                cell.Value = raw;
                break;
            default:
                throw new ConfigException($"Unknown Type '{u.Type}' for {u.Sheet}!{u.Cell} (use number, date, bool, string).");
        }

        Log("INFO", $"{u.Sheet}!{u.Cell}: '{old}' -> '{raw}'");
    }
}

static void Recalculate(XLWorkbook wb, JobOptions o)
{
    // Excel itself will also do a full recalc the next time the file is opened.
    wb.FullCalculationOnLoad = true;

    try
    {
        wb.RecalculateAllFormulas();   // stores fresh cached values so non-Excel readers see them too
    }
    catch (Exception ex)
    {
        // e.g. a function ClosedXML can't evaluate. Fall back to recalc-on-open and flag it.
        Log("WARN", $"Could not recalculate in-process ({ex.Message}). Excel will recalculate on next open; cached values may be stale.");
        if (o.FailOnFormulaError) throw;
        return;
    }

    var errors = wb.Worksheets
        .SelectMany(ws => ws.CellsUsed(c => c.HasFormula && c.CachedValue.IsError)
        .Select(c => $"{ws.Name}!{c.Address}: {c.CachedValue}"))
        .ToList();

    if (errors.Count == 0) return;

    foreach (var e in errors.Take(20)) Log("WARN", $"Formula error {e}");
    if (o.FailOnFormulaError)
        throw new InvalidOperationException($"{errors.Count} formula cell(s) evaluated to an error; file left unchanged.");
}

static void LogVerifyCells(XLWorkbook wb, JobOptions o)
{
    foreach (var v in o.VerifyCells)
    {
        if (wb.TryGetWorksheet(v.Sheet, out var ws))
            Log("INFO", $"Verify {v.Sheet}!{v.Cell} = {ws.Cell(v.Cell).Value}");
    }
}

// ---------------------------------------------------------------------------

class ConfigException(string message) : Exception(message);

class JobOptions
{
    public string FilePath { get; set; } = "";
    public string? BackupFolder { get; set; }
    public int BackupRetentionDays { get; set; } = 30;
    public int LockRetryCount { get; set; } = 6;
    public int LockRetryDelaySeconds { get; set; } = 10;
    public bool FailOnFormulaError { get; set; } = true;
    public List<CellUpdate> Updates { get; set; } = new();
    public List<CellRef> VerifyCells { get; set; } = new();
}

class CellUpdate
{
    public string Sheet { get; set; } = "";
    public string Cell { get; set; } = "";
    public string Type { get; set; } = "string";
    public string Value { get; set; } = "";
}

class CellRef
{
    public string Sheet { get; set; } = "";
    public string Cell { get; set; } = "";
}
