using System.Text.Json;
using MongoDB.Bson;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;

var options = ParseArgs(args);
if (options.ShowHelp)
{
    PrintUsage();
    return;
}

var connectionString = ResolveConnectionString(options.ConnectionString);
var databaseName = ResolveDatabaseName(options.DatabaseName);

if (string.IsNullOrWhiteSpace(connectionString))
{
    Console.WriteLine("ERROR: Không tìm thấy MongoDbSettings__ConnectionString.");
    Console.WriteLine("Hãy truyền --connection hoặc set env MongoDbSettings__ConnectionString.");
    return;
}

if (string.IsNullOrWhiteSpace(databaseName))
{
    Console.WriteLine("ERROR: Không tìm thấy MongoDbSettings__DatabaseName.");
    Console.WriteLine("Hãy truyền --database hoặc set env MongoDbSettings__DatabaseName.");
    return;
}

var client = new MongoClient(connectionString);
var database = client.GetDatabase(databaseName);
var assignments = database.GetCollection<Assignment>("assignments");
var publications = database.GetCollection<BsonDocument>("examPublications");

var filters = new List<FilterDefinition<Assignment>>();
if (!string.IsNullOrWhiteSpace(options.ClassId))
{
    filters.Add(BuildIdFilter(options.ClassId!, value => value.ClassId));
}

if (!string.IsNullOrWhiteSpace(options.AssignmentId))
{
    filters.Add(BuildIdFilter(options.AssignmentId!, value => value.Id));
}

if (!options.IncludeInactive)
{
    filters.Add(Builders<Assignment>.Filter.Eq(item => item.IsActive, true));
}

var baseFilter = filters.Count == 0
    ? Builders<Assignment>.Filter.Empty
    : filters.Count == 1
        ? filters[0]
        : Builders<Assignment>.Filter.And(filters);

var assignmentList = await assignments
    .Find(baseFilter)
    .SortBy(item => item.ClassId)
    .ThenBy(item => item.Name)
    .ToListAsync();

var assignmentIds = assignmentList
    .Where(item => !string.IsNullOrWhiteSpace(item.Id))
    .Select(item => item.Id)
    .ToHashSet(StringComparer.OrdinalIgnoreCase);

var publicationDocs = assignmentIds.Count == 0
    ? new List<BsonDocument>()
    : await publications.Find(Builders<BsonDocument>.Filter.Empty).ToListAsync();

var lockedIds = publicationDocs
    .SelectMany(ReadSourceAssignmentIds)
    .Where(id => assignmentIds.Contains(id))
    .ToHashSet(StringComparer.OrdinalIgnoreCase);

var rows = assignmentList
    .Select(item =>
    {
        item.IsLockedForPublication = !string.IsNullOrWhiteSpace(item.Id) && lockedIds.Contains(item.Id);
        var (isPublishable, reason) = AssignmentService.EvaluatePublicationEligibility(item);
        item.IsPublishable = isPublishable;
        item.PublishBlockReason = reason;

        var legacyIssues = GetLegacyIssues(item);
        return new AuditRow
        {
            Id = item.Id,
            ClassId = item.ClassId,
            Name = item.Name,
            Subject = item.Subject,
            ExamType = item.ExamType,
            ProjectCode = item.ProjectCode,
            GradingType = item.GradingType,
            GradingApiEndpoint = item.GradingApiEndpoint,
            IsActive = item.IsActive,
            IsLockedForPublication = item.IsLockedForPublication,
            IsPublishable = item.IsPublishable,
            PublishBlockReason = item.PublishBlockReason,
            LegacyIssues = legacyIssues
        };
    })
    .ToList();

var blockedRows = rows
    .Where(row => !row.IsPublishable || row.LegacyIssues.Count > 0)
    .ToList();

Console.WriteLine("=== ASSIGNMENT PUBLICATION AUDIT ===");
Console.WriteLine($"Connection       : {MaskConnection(connectionString)}");
Console.WriteLine($"Database         : {databaseName}");
Console.WriteLine($"ClassId          : {options.ClassId ?? "(all)"}");
Console.WriteLine($"AssignmentId     : {options.AssignmentId ?? "(all)"}");
Console.WriteLine($"Include inactive : {options.IncludeInactive}");
Console.WriteLine($"Matched rows     : {rows.Count}");
Console.WriteLine($"Blocked rows     : {blockedRows.Count}");
Console.WriteLine($"Preview          : {options.PreviewSize}");

if (rows.Count == 0)
{
    Console.WriteLine("Không có assignment nào khớp filter.");
    return;
}

Console.WriteLine();
Console.WriteLine("--- Summary ---");
Console.WriteLine($"Publishable                  : {rows.Count(row => row.IsPublishable)}");
Console.WriteLine($"Not publishable              : {rows.Count(row => !row.IsPublishable)}");
Console.WriteLine($"Locked for publication       : {rows.Count(row => row.IsLockedForPublication)}");
Console.WriteLine($"Missing examType             : {rows.Count(row => row.LegacyIssues.Contains("missing:examType"))}");
Console.WriteLine($"Missing subject              : {rows.Count(row => row.LegacyIssues.Contains("missing:subject"))}");
Console.WriteLine($"Missing projectCode          : {rows.Count(row => row.LegacyIssues.Contains("missing:projectCode"))}");
Console.WriteLine($"Missing gradingApiEndpoint   : {rows.Count(row => row.LegacyIssues.Contains("missing:gradingApiEndpoint"))}");
Console.WriteLine($"Invalid route metadata       : {rows.Count(row => row.LegacyIssues.Contains("invalid:graderRoute"))}");

Console.WriteLine();
Console.WriteLine("--- Preview blocked assignments ---");
foreach (var row in blockedRows.Take(options.PreviewSize))
{
    Console.WriteLine(
        $"id={row.Id} | class={row.ClassId} | active={row.IsActive} | locked={row.IsLockedForPublication} | publishable={row.IsPublishable}");
    Console.WriteLine(
        $"  name={row.Name} | examType={row.ExamType ?? "(null)"} | subject={row.Subject ?? "(null)"} | projectCode={row.ProjectCode ?? "(null)"}");
    Console.WriteLine(
        $"  gradingType={row.GradingType ?? "(null)"} | endpoint={row.GradingApiEndpoint ?? "(null)"}");

    if (row.LegacyIssues.Count > 0)
    {
        Console.WriteLine($"  legacyIssues={string.Join(", ", row.LegacyIssues)}");
    }

    if (!string.IsNullOrWhiteSpace(row.PublishBlockReason))
    {
        Console.WriteLine($"  publishBlockReason={row.PublishBlockReason}");
    }
}

if (blockedRows.Count > options.PreviewSize)
{
    Console.WriteLine($"... còn {blockedRows.Count - options.PreviewSize} assignment bị block chưa hiển thị.");
}

return;

static AuditOptions ParseArgs(string[] args)
{
    var options = new AuditOptions();
    for (var i = 0; i < args.Length; i++)
    {
        var current = args[i].Trim();
        switch (current)
        {
            case "--help":
            case "-h":
                options.ShowHelp = true;
                break;
            case "--connection":
                options.ConnectionString = ReadNextValue(args, ref i, "--connection");
                break;
            case "--database":
                options.DatabaseName = ReadNextValue(args, ref i, "--database");
                break;
            case "--class-id":
                options.ClassId = ReadNextValue(args, ref i, "--class-id");
                break;
            case "--assignment-id":
                options.AssignmentId = ReadNextValue(args, ref i, "--assignment-id");
                break;
            case "--include-inactive":
                options.IncludeInactive = true;
                break;
            case "--preview":
                options.PreviewSize = ParsePositiveInt(ReadNextValue(args, ref i, "--preview"), "--preview");
                break;
            default:
                throw new ArgumentException($"Unknown argument: {current}");
        }
    }

    return options;
}

static string ReadNextValue(string[] args, ref int index, string optionName)
{
    if (index + 1 >= args.Length)
    {
        throw new ArgumentException($"Missing value for {optionName}");
    }

    index++;
    return args[index];
}

static int ParsePositiveInt(string value, string optionName)
{
    if (!int.TryParse(value, out var parsed) || parsed <= 0)
    {
        throw new ArgumentException($"Invalid value for {optionName}: {value}");
    }

    return parsed;
}

static FilterDefinition<Assignment> BuildIdFilter(string rawId, System.Linq.Expressions.Expression<Func<Assignment, string>> field)
{
    var trimmed = rawId.Trim();
    return Builders<Assignment>.Filter.Eq(field, trimmed);
}

static IEnumerable<string> ReadSourceAssignmentIds(BsonDocument publication)
{
    if (!publication.TryGetValue("projectSequence", out var sequenceValue) ||
        sequenceValue.IsBsonNull ||
        sequenceValue.BsonType != BsonType.Array)
    {
        yield break;
    }

    foreach (var item in sequenceValue.AsBsonArray)
    {
        if (item.IsBsonNull || item.BsonType != BsonType.Document)
        {
            continue;
        }

        var document = item.AsBsonDocument;
        if (!document.TryGetValue("sourceAssignmentId", out var sourceAssignmentId) || sourceAssignmentId.IsBsonNull)
        {
            continue;
        }

        var value = sourceAssignmentId.BsonType == BsonType.ObjectId
            ? sourceAssignmentId.AsObjectId.ToString()
            : sourceAssignmentId.ToString();

        if (!string.IsNullOrWhiteSpace(value))
        {
            yield return value.Trim();
        }
    }
}

static List<string> GetLegacyIssues(Assignment assignment)
{
    var issues = new List<string>();

    if (string.IsNullOrWhiteSpace(assignment.ExamType))
    {
        issues.Add("missing:examType");
    }

    if (string.IsNullOrWhiteSpace(assignment.Subject))
    {
        issues.Add("missing:subject");
    }

    if (string.IsNullOrWhiteSpace(assignment.ProjectCode))
    {
        issues.Add("missing:projectCode");
    }

    if (assignment.GradingType == GradingTypes.Auto && string.IsNullOrWhiteSpace(assignment.GradingApiEndpoint))
    {
        issues.Add("missing:gradingApiEndpoint");
    }

    if (!string.IsNullOrWhiteSpace(assignment.ExamType) &&
        !string.IsNullOrWhiteSpace(assignment.Subject) &&
        !string.IsNullOrWhiteSpace(assignment.ProjectCode) &&
        assignment.GradingType == GradingTypes.Auto &&
        !string.IsNullOrWhiteSpace(assignment.GradingApiEndpoint) &&
        !GraderRouteRegistry.TryResolve(
            assignment.ExamType,
            assignment.Subject,
            assignment.ProjectCode,
            assignment.GradingApiEndpoint,
            out _))
    {
        issues.Add("invalid:graderRoute");
    }

    return issues;
}

static string ResolveConnectionString(string? explicitValue)
{
    if (!string.IsNullOrWhiteSpace(explicitValue))
    {
        return explicitValue.Trim();
    }

    var fromEnv = Environment.GetEnvironmentVariable("MongoDbSettings__ConnectionString");
    if (!string.IsNullOrWhiteSpace(fromEnv))
    {
        return fromEnv.Trim();
    }

    return ReadSettingFromDevelopmentFile("ConnectionString") ?? string.Empty;
}

static string ResolveDatabaseName(string? explicitValue)
{
    if (!string.IsNullOrWhiteSpace(explicitValue))
    {
        return explicitValue.Trim();
    }

    var fromEnv = Environment.GetEnvironmentVariable("MongoDbSettings__DatabaseName");
    if (!string.IsNullOrWhiteSpace(fromEnv))
    {
        return fromEnv.Trim();
    }

    return ReadSettingFromDevelopmentFile("DatabaseName") ?? string.Empty;
}

static string? ReadSettingFromDevelopmentFile(string settingName)
{
    var appSettingsPath = TryFindAppSettingsDevelopmentPath();
    if (string.IsNullOrWhiteSpace(appSettingsPath) || !File.Exists(appSettingsPath))
    {
        return null;
    }

    using var stream = File.OpenRead(appSettingsPath);
    using var document = JsonDocument.Parse(stream);

    if (!document.RootElement.TryGetProperty("MongoDbSettings", out var mongoSection))
    {
        return null;
    }

    if (!mongoSection.TryGetProperty(settingName, out var valueElement))
    {
        return null;
    }

    var value = valueElement.GetString();
    return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}

static string? TryFindAppSettingsDevelopmentPath()
{
    var current = new DirectoryInfo(Directory.GetCurrentDirectory());
    while (current != null)
    {
        var candidate = Path.Combine(current.FullName, "MOS.ExcelGrading.API", "appsettings.Development.json");
        if (File.Exists(candidate))
        {
            return candidate;
        }

        current = current.Parent;
    }

    return null;
}

static string MaskConnection(string connectionString)
{
    if (string.IsNullOrWhiteSpace(connectionString))
    {
        return "(empty)";
    }

    var atIndex = connectionString.IndexOf('@');
    if (atIndex <= 0)
    {
        return connectionString;
    }

    var protocolSeparator = connectionString.IndexOf("://", StringComparison.Ordinal);
    if (protocolSeparator < 0 || protocolSeparator + 3 >= atIndex)
    {
        return connectionString;
    }

    return connectionString[..(protocolSeparator + 3)] + "***:***" + connectionString[atIndex..];
}

static void PrintUsage()
{
    Console.WriteLine("Usage:");
    Console.WriteLine("  dotnet run --project tools/AssignmentPublicationAudit/AssignmentPublicationAudit.csproj -- [options]");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  --class-id <id>          Chỉ audit assignment của một lớp");
    Console.WriteLine("  --assignment-id <id>     Chỉ audit một assignment");
    Console.WriteLine("  --connection <value>     Mongo connection string");
    Console.WriteLine("  --database <value>       Mongo database name");
    Console.WriteLine("  --include-inactive       Bao gồm assignment đã inactive");
    Console.WriteLine("  --preview <n>            Số dòng preview assignment bị block (default 20)");
    Console.WriteLine("  -h | --help              Hiển thị hướng dẫn");
}

internal sealed class AuditOptions
{
    public bool ShowHelp { get; set; }
    public string? ConnectionString { get; set; }
    public string? DatabaseName { get; set; }
    public string? ClassId { get; set; }
    public string? AssignmentId { get; set; }
    public bool IncludeInactive { get; set; }
    public int PreviewSize { get; set; } = 20;
}

internal sealed class AuditRow
{
    public string Id { get; init; } = string.Empty;
    public string ClassId { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string? Subject { get; init; }
    public string? ExamType { get; init; }
    public string? ProjectCode { get; init; }
    public string? GradingType { get; init; }
    public string? GradingApiEndpoint { get; init; }
    public bool IsActive { get; init; }
    public bool IsLockedForPublication { get; init; }
    public bool IsPublishable { get; init; }
    public string? PublishBlockReason { get; init; }
    public List<string> LegacyIssues { get; init; } = new();
}
