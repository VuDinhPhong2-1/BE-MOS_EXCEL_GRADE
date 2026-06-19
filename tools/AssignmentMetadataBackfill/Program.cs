using System.Text;
using System.Text.Json;
using MongoDB.Bson;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.Models;

Console.OutputEncoding = Encoding.UTF8;

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
var assignments = database.GetCollection<BsonDocument>("assignments");

var filters = new List<FilterDefinition<BsonDocument>>();
if (!string.IsNullOrWhiteSpace(options.ClassId))
{
    filters.Add(BuildIdFilter("classId", options.ClassId!));
}

if (!string.IsNullOrWhiteSpace(options.AssignmentId))
{
    filters.Add(BuildIdFilter("_id", options.AssignmentId!));
}

if (!options.IncludeInactive)
{
    filters.Add(Builders<BsonDocument>.Filter.Eq("isActive", true));
}

var baseFilter = filters.Count == 0
    ? Builders<BsonDocument>.Filter.Empty
    : filters.Count == 1
        ? filters[0]
        : Builders<BsonDocument>.Filter.And(filters);

var docs = await assignments
    .Find(baseFilter)
    .Sort(Builders<BsonDocument>.Sort.Ascending("classId").Ascending("name"))
    .ToListAsync();

var plans = docs
    .Select(document => BuildPlan(document))
    .ToList();

var changedPlans = plans.Where(plan => plan.Changes.Count > 0).ToList();
var publishableAfterPlans = plans.Where(plan => plan.IsPublishableAfter).ToList();

Console.WriteLine("=== ASSIGNMENT METADATA BACKFILL ===");
Console.WriteLine($"Connection        : {MaskConnection(connectionString)}");
Console.WriteLine($"Database          : {databaseName}");
Console.WriteLine($"ClassId           : {options.ClassId ?? "(all)"}");
Console.WriteLine($"AssignmentId      : {options.AssignmentId ?? "(all)"}");
Console.WriteLine($"Include inactive  : {options.IncludeInactive}");
Console.WriteLine($"Matched docs      : {plans.Count}");
Console.WriteLine($"Would change      : {changedPlans.Count}");
Console.WriteLine($"Publishable after : {publishableAfterPlans.Count}");
Console.WriteLine($"Mode              : {(options.Apply ? "APPLY (backup + replace)" : "DRY-RUN")}");
Console.WriteLine($"Preview           : {options.PreviewSize}");

if (plans.Count == 0)
{
    Console.WriteLine("Không có assignment nào khớp filter.");
    return;
}

Console.WriteLine();
Console.WriteLine("--- Summary ---");
Console.WriteLine($"Missing examType before           : {plans.Count(plan => plan.MissingFieldsBefore.Contains("examType"))}");
Console.WriteLine($"Missing subject before            : {plans.Count(plan => plan.MissingFieldsBefore.Contains("subject"))}");
Console.WriteLine($"Missing projectCode before        : {plans.Count(plan => plan.MissingFieldsBefore.Contains("projectCode"))}");
Console.WriteLine($"Missing gradingType before        : {plans.Count(plan => plan.MissingFieldsBefore.Contains("gradingType"))}");
Console.WriteLine($"Missing gradingApiEndpoint before : {plans.Count(plan => plan.MissingFieldsBefore.Contains("gradingApiEndpoint"))}");
Console.WriteLine($"Invalid route before              : {plans.Count(plan => !plan.IsRouteValidBefore)}");
Console.WriteLine($"Invalid route after               : {plans.Count(plan => !plan.IsRouteValidAfter)}");

Console.WriteLine();
Console.WriteLine("--- Preview changes ---");
foreach (var plan in changedPlans.Take(options.PreviewSize))
{
    Console.WriteLine(
        $"id={plan.Id} | class={plan.ClassId} | active={plan.IsActive} | publishableAfter={plan.IsPublishableAfter}");
    Console.WriteLine($"  name={plan.Name}");
    Console.WriteLine(
        $"  before examType={plan.ExamTypeBefore ?? "(null)"} | subject={plan.SubjectBefore ?? "(null)"} | projectCode={plan.ProjectCodeBefore ?? "(null)"} | gradingType={plan.GradingTypeBefore ?? "(null)"} | endpoint={plan.GradingApiEndpointBefore ?? "(null)"}");
    Console.WriteLine(
        $"  after  examType={plan.ExamTypeAfter ?? "(null)"} | subject={plan.SubjectAfter ?? "(null)"} | projectCode={plan.ProjectCodeAfter ?? "(null)"} | gradingType={plan.GradingTypeAfter ?? "(null)"} | endpoint={plan.GradingApiEndpointAfter ?? "(null)"}");
    Console.WriteLine($"  changes={string.Join(" | ", plan.Changes)}");
    if (!string.IsNullOrWhiteSpace(plan.PublishBlockReasonAfter))
    {
        Console.WriteLine($"  publishBlockReasonAfter={plan.PublishBlockReasonAfter}");
    }
}

if (changedPlans.Count > options.PreviewSize)
{
    Console.WriteLine($"... còn {changedPlans.Count - options.PreviewSize} assignment có thay đổi chưa hiển thị.");
}

if (!options.Apply)
{
    Console.WriteLine();
    Console.WriteLine("Dry-run complete. Chạy thêm --apply để backup và backfill metadata.");
    return;
}

if (changedPlans.Count == 0)
{
    Console.WriteLine();
    Console.WriteLine("Không có assignment nào cần backfill.");
    return;
}

var archiveCollectionName = string.IsNullOrWhiteSpace(options.ArchiveCollection)
    ? $"assignments_archive_metadata_backfill_{DateTime.UtcNow:yyyyMMdd_HHmmss}"
    : options.ArchiveCollection.Trim();
var archiveCollection = database.GetCollection<BsonDocument>(archiveCollectionName);
var migrationTag = $"assignment-metadata-backfill-{DateTime.UtcNow:yyyyMMdd-HHmmss}";

Console.WriteLine();
Console.WriteLine($"Archive collection: {archiveCollectionName}");
Console.WriteLine($"Migration tag     : {migrationTag}");

var batchSize = options.BatchSize;
var totalBackedUp = 0;
var totalUpdated = 0;
var batchNo = 0;

foreach (var batch in changedPlans.Chunk(batchSize))
{
    batchNo++;
    var now = DateTime.UtcNow;
    var backups = new List<BsonDocument>(batch.Length);
    var replaceModels = new List<WriteModel<BsonDocument>>(batch.Length);

    foreach (var plan in batch)
    {
        var backupDoc = plan.OriginalDocument.DeepClone().AsBsonDocument;
        backupDoc.Remove("_id");
        backupDoc["sourceAssignmentId"] = plan.OriginalDocument["_id"];
        backupDoc["migratedAtUtc"] = now;
        backupDoc["migrationTag"] = migrationTag;
        backupDoc["changes"] = new BsonArray(plan.Changes);
        backups.Add(backupDoc);

        replaceModels.Add(new ReplaceOneModel<BsonDocument>(
            Builders<BsonDocument>.Filter.Eq("_id", plan.OriginalDocument["_id"]),
            plan.NormalizedDocument));
    }

    await archiveCollection.InsertManyAsync(backups);
    var result = await assignments.BulkWriteAsync(replaceModels);

    totalBackedUp += backups.Count;
    totalUpdated += (int)result.ModifiedCount;

    Console.WriteLine(
        $"Batch #{batchNo}: backedUp={backups.Count}, updated={result.ModifiedCount}, totalUpdated={totalUpdated}/{changedPlans.Count}");
}

Console.WriteLine();
Console.WriteLine("=== DONE ===");
Console.WriteLine($"Backed up docs : {totalBackedUp}");
Console.WriteLine($"Updated docs   : {totalUpdated}");
Console.WriteLine($"Archive        : {archiveCollectionName}");
Console.WriteLine($"Migration tag  : {migrationTag}");

static BackfillPlan BuildPlan(BsonDocument source)
{
    var normalized = source.DeepClone().AsBsonDocument;
    var changes = new List<string>();

    var originalName = ReadString(source, "name") ?? "(unnamed)";
    var originalClassId = ReadObjectIdLikeString(source, "classId") ?? string.Empty;
    var isActive = ReadBool(source, "isActive") ?? true;

    var gradingApiEndpointBefore = ReadString(source, "gradingApiEndpoint");
    var gradingTypeBefore = NormalizeNullable(ReadString(source, "gradingType"));
    var subjectBefore = NormalizeNullable(ReadString(source, "subject"));
    var examTypeBefore = NormalizeNullable(ReadString(source, "examType"));
    var projectCodeBefore = NormalizeNullable(ReadString(source, "projectCode"));

    var normalizedEndpoint = string.IsNullOrWhiteSpace(gradingApiEndpointBefore)
        ? null
        : GradingApiEndpoints.NormalizeEndpoint(gradingApiEndpointBefore);

    if (!string.Equals(gradingApiEndpointBefore, normalizedEndpoint, StringComparison.Ordinal))
    {
        SetOrRemove(normalized, "gradingApiEndpoint", normalizedEndpoint, changes, "gradingApiEndpoint");
    }

    var inferredGradingType = InferGradingType(gradingTypeBefore, normalizedEndpoint);
    if (!string.Equals(gradingTypeBefore, inferredGradingType, StringComparison.Ordinal))
    {
        SetOrRemove(normalized, "gradingType", inferredGradingType, changes, "gradingType");
    }

    var inferredSubject = InferSubject(subjectBefore, normalizedEndpoint);
    if (!string.Equals(subjectBefore, inferredSubject, StringComparison.Ordinal))
    {
        SetOrRemove(normalized, "subject", inferredSubject, changes, "subject");
    }

    var inferredExamType = InferExamType(examTypeBefore, inferredGradingType, inferredSubject, normalizedEndpoint, source);
    if (!string.Equals(examTypeBefore, inferredExamType, StringComparison.Ordinal))
    {
        SetOrRemove(normalized, "examType", inferredExamType, changes, "examType");
    }

    var inferredProjectCode = InferProjectCode(projectCodeBefore, inferredSubject, normalizedEndpoint);
    if (!string.Equals(projectCodeBefore, inferredProjectCode, StringComparison.Ordinal))
    {
        SetOrRemove(normalized, "projectCode", inferredProjectCode, changes, "projectCode");
    }

    var examTypeAfter = NormalizeNullable(ReadString(normalized, "examType"));
    var subjectAfter = NormalizeNullable(ReadString(normalized, "subject"));
    var projectCodeAfter = NormalizeNullable(ReadString(normalized, "projectCode"));
    var gradingTypeAfter = NormalizeNullable(ReadString(normalized, "gradingType"));
    var gradingApiEndpointAfter = NormalizeNullable(ReadString(normalized, "gradingApiEndpoint"));

    var routeBefore = TryResolveRoute(examTypeBefore, subjectBefore, projectCodeBefore, gradingApiEndpointBefore, out _);
    var routeAfter = TryResolveRoute(examTypeAfter, subjectAfter, projectCodeAfter, gradingApiEndpointAfter, out _);

    var assignmentAfter = new Assignment
    {
        Id = ReadObjectIdLikeString(normalized, "_id") ?? string.Empty,
        Name = originalName,
        ClassId = originalClassId,
        IsActive = isActive,
        GradingType = gradingTypeAfter ?? GradingTypes.Manual,
        GradingApiEndpoint = gradingApiEndpointAfter,
        Subject = subjectAfter ?? string.Empty,
        ExamType = examTypeAfter ?? string.Empty,
        ProjectCode = projectCodeAfter
    };

    var (isPublishableAfter, publishBlockReasonAfter) = MOS.ExcelGrading.Core.Services.AssignmentService.EvaluatePublicationEligibility(assignmentAfter);

    return new BackfillPlan
    {
        OriginalDocument = source,
        NormalizedDocument = normalized,
        Id = ReadObjectIdLikeString(source, "_id") ?? string.Empty,
        ClassId = originalClassId,
        Name = originalName,
        IsActive = isActive,
        ExamTypeBefore = examTypeBefore,
        SubjectBefore = subjectBefore,
        ProjectCodeBefore = projectCodeBefore,
        GradingTypeBefore = gradingTypeBefore,
        GradingApiEndpointBefore = gradingApiEndpointBefore,
        ExamTypeAfter = examTypeAfter,
        SubjectAfter = subjectAfter,
        ProjectCodeAfter = projectCodeAfter,
        GradingTypeAfter = gradingTypeAfter,
        GradingApiEndpointAfter = gradingApiEndpointAfter,
        IsRouteValidBefore = routeBefore,
        IsRouteValidAfter = routeAfter,
        IsPublishableAfter = isPublishableAfter,
        PublishBlockReasonAfter = publishBlockReasonAfter,
        MissingFieldsBefore = GetMissingFields(examTypeBefore, subjectBefore, projectCodeBefore, gradingTypeBefore, gradingApiEndpointBefore),
        Changes = changes
    };
}

static string? InferGradingType(string? currentGradingType, string? normalizedEndpoint)
{
    if (!string.IsNullOrWhiteSpace(currentGradingType))
    {
        var normalized = currentGradingType.Trim().ToLowerInvariant();
        if (normalized == GradingTypes.Auto || normalized == GradingTypes.Manual)
        {
            return normalized;
        }
    }

    return !string.IsNullOrWhiteSpace(normalizedEndpoint)
        ? GradingTypes.Auto
        : currentGradingType;
}

static string? InferSubject(string? currentSubject, string? normalizedEndpoint)
{
    var normalized = NormalizeNullable(currentSubject);
    if (AssignmentFileSubjects.IsValid(normalized))
    {
        return normalized;
    }

    if (GradingApiEndpoints.TryExtractSubject(normalizedEndpoint, out var endpointSubject))
    {
        return endpointSubject;
    }

    return normalized;
}

static string? InferExamType(
    string? currentExamType,
    string? inferredGradingType,
    string? inferredSubject,
    string? normalizedEndpoint,
    BsonDocument source)
{
    var normalized = NormalizeNullable(currentExamType);
    if (AssignmentExamTypes.IsValid(normalized))
    {
        return normalized;
    }

    var name = NormalizeText(ReadString(source, "name"));
    var description = NormalizeText(ReadString(source, "description"));
    var endpointText = NormalizeText(normalizedEndpoint);

    var isOnThi = name.Contains("on thi") ||
                  description.Contains("on thi") ||
                  name.Contains("exam review") ||
                  description.Contains("exam review") ||
                  endpointText.Contains("exam review");

    if (isOnThi)
    {
        return AssignmentExamTypes.OnThi;
    }

    if (inferredGradingType == GradingTypes.Auto &&
        !string.IsNullOrWhiteSpace(inferredSubject) &&
        !string.IsNullOrWhiteSpace(normalizedEndpoint))
    {
        return AssignmentExamTypes.OTTH;
    }

    return normalized;
}

static string? InferProjectCode(string? currentProjectCode, string? inferredSubject, string? normalizedEndpoint)
{
    var normalizedProjectCode = NormalizeNullable(currentProjectCode)?.ToUpperInvariant();
    if (!string.IsNullOrWhiteSpace(normalizedProjectCode))
    {
        return normalizedProjectCode;
    }

    if (string.IsNullOrWhiteSpace(inferredSubject) ||
        !GradingApiEndpoints.TryExtractProjectNumber(normalizedEndpoint, out var projectNumber))
    {
        return normalizedProjectCode;
    }

    return $"{inferredSubject.ToUpperInvariant()}_P{projectNumber:00}";
}

static bool TryResolveRoute(
    string? examType,
    string? subject,
    string? projectCode,
    string? gradingApiEndpoint,
    out GraderRouteDescriptor route)
{
    route = new GraderRouteDescriptor();
    if (string.IsNullOrWhiteSpace(examType) ||
        string.IsNullOrWhiteSpace(subject) ||
        string.IsNullOrWhiteSpace(projectCode) ||
        string.IsNullOrWhiteSpace(gradingApiEndpoint))
    {
        return false;
    }

    return GraderRouteRegistry.TryResolve(examType, subject, projectCode, gradingApiEndpoint, out route);
}

static List<string> GetMissingFields(
    string? examType,
    string? subject,
    string? projectCode,
    string? gradingType,
    string? gradingApiEndpoint)
{
    var missing = new List<string>();
    if (string.IsNullOrWhiteSpace(examType)) missing.Add("examType");
    if (string.IsNullOrWhiteSpace(subject)) missing.Add("subject");
    if (string.IsNullOrWhiteSpace(projectCode)) missing.Add("projectCode");
    if (string.IsNullOrWhiteSpace(gradingType)) missing.Add("gradingType");
    if (string.IsNullOrWhiteSpace(gradingApiEndpoint)) missing.Add("gradingApiEndpoint");
    return missing;
}

static void SetOrRemove(BsonDocument doc, string fieldName, string? value, List<string> changes, string label)
{
    var before = ReadString(doc, fieldName);
    var normalizedValue = NormalizeNullable(value);

    if (string.IsNullOrWhiteSpace(normalizedValue))
    {
        if (doc.Contains(fieldName))
        {
            doc.Remove(fieldName);
            changes.Add($"{label}: '{before ?? "(null)"}' -> <removed>");
        }

        return;
    }

    if (!doc.Contains(fieldName) || !string.Equals(before, normalizedValue, StringComparison.Ordinal))
    {
        doc[fieldName] = normalizedValue;
        changes.Add($"{label}: '{before ?? "(null)"}' -> '{normalizedValue}'");
    }
}

static BackfillOptions ParseArgs(string[] args)
{
    var options = new BackfillOptions();
    for (var i = 0; i < args.Length; i++)
    {
        var current = args[i].Trim();
        switch (current)
        {
            case "--help":
            case "-h":
                options.ShowHelp = true;
                break;
            case "--apply":
                options.Apply = true;
                break;
            case "--connection":
                options.ConnectionString = ReadNextValue(args, ref i, "--connection");
                break;
            case "--database":
                options.DatabaseName = ReadNextValue(args, ref i, "--database");
                break;
            case "--archive":
                options.ArchiveCollection = ReadNextValue(args, ref i, "--archive");
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
            case "--batch-size":
                options.BatchSize = ParsePositiveInt(ReadNextValue(args, ref i, "--batch-size"), "--batch-size");
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

static FilterDefinition<BsonDocument> BuildIdFilter(string fieldName, string rawId)
{
    var trimmed = rawId.Trim();
    var filter = Builders<BsonDocument>.Filter.Eq(fieldName, trimmed);
    if (ObjectId.TryParse(trimmed, out var objectId))
    {
        filter |= Builders<BsonDocument>.Filter.Eq(fieldName, objectId);
    }

    return filter;
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

static string? ReadString(BsonDocument document, string fieldName)
{
    if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
    {
        return null;
    }

    return value.BsonType switch
    {
        BsonType.ObjectId => value.AsObjectId.ToString(),
        BsonType.String => value.AsString,
        _ => value.ToString()
    };
}

static string? ReadObjectIdLikeString(BsonDocument document, string fieldName)
{
    return ReadString(document, fieldName)?.Trim();
}

static bool? ReadBool(BsonDocument document, string fieldName)
{
    if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
    {
        return null;
    }

    return value.BsonType switch
    {
        BsonType.Boolean => value.AsBoolean,
        BsonType.Int32 => value.AsInt32 != 0,
        BsonType.Int64 => value.AsInt64 != 0,
        BsonType.String when bool.TryParse(value.AsString, out var parsed) => parsed,
        _ => null
    };
}

static string? NormalizeNullable(string? value) =>
    string.IsNullOrWhiteSpace(value) ? null : value.Trim().ToLowerInvariant();

static string NormalizeText(string? value) =>
    (value ?? string.Empty)
        .Normalize(NormalizationForm.FormD)
        .Where(ch => ch < 128 || char.GetUnicodeCategory(ch) != System.Globalization.UnicodeCategory.NonSpacingMark)
        .Aggregate(new StringBuilder(), (builder, ch) => builder.Append(ch))
        .ToString()
        .ToLowerInvariant()
        .Trim();

static void PrintUsage()
{
    Console.WriteLine("Usage:");
    Console.WriteLine("  dotnet run --project tools/AssignmentMetadataBackfill/AssignmentMetadataBackfill.csproj -- [options]");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  --apply                  Run backup + replace (default is dry-run).");
    Console.WriteLine("  --connection <value>     Mongo connection string.");
    Console.WriteLine("  --database <value>       Mongo database name.");
    Console.WriteLine("  --archive <value>        Archive collection name.");
    Console.WriteLine("  --class-id <value>       Optional class filter.");
    Console.WriteLine("  --assignment-id <value>  Optional assignment filter.");
    Console.WriteLine("  --include-inactive       Include inactive assignments.");
    Console.WriteLine("  --batch-size <n>         Batch size for apply mode (default 200).");
    Console.WriteLine("  --preview <n>            Preview rows in dry-run/apply header (default 20).");
    Console.WriteLine("  -h | --help              Show this help.");
}

internal sealed class BackfillOptions
{
    public bool ShowHelp { get; set; }
    public bool Apply { get; set; }
    public string? ConnectionString { get; set; }
    public string? DatabaseName { get; set; }
    public string? ArchiveCollection { get; set; }
    public string? ClassId { get; set; }
    public string? AssignmentId { get; set; }
    public bool IncludeInactive { get; set; }
    public int BatchSize { get; set; } = 200;
    public int PreviewSize { get; set; } = 20;
}

internal sealed class BackfillPlan
{
    public required BsonDocument OriginalDocument { get; init; }
    public required BsonDocument NormalizedDocument { get; init; }
    public string Id { get; init; } = string.Empty;
    public string ClassId { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public bool IsActive { get; init; }
    public string? ExamTypeBefore { get; init; }
    public string? SubjectBefore { get; init; }
    public string? ProjectCodeBefore { get; init; }
    public string? GradingTypeBefore { get; init; }
    public string? GradingApiEndpointBefore { get; init; }
    public string? ExamTypeAfter { get; init; }
    public string? SubjectAfter { get; init; }
    public string? ProjectCodeAfter { get; init; }
    public string? GradingTypeAfter { get; init; }
    public string? GradingApiEndpointAfter { get; init; }
    public bool IsRouteValidBefore { get; init; }
    public bool IsRouteValidAfter { get; init; }
    public bool IsPublishableAfter { get; init; }
    public string? PublishBlockReasonAfter { get; init; }
    public List<string> MissingFieldsBefore { get; init; } = new();
    public List<string> Changes { get; init; } = new();
}
