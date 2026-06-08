using System.Globalization;
using System.Text.Json;
using MongoDB.Bson;
using MongoDB.Driver;

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
    Console.WriteLine("ERROR: Khong tim thay MongoDbSettings__ConnectionString.");
    Console.WriteLine("Hay truyen --connection hoac set env MongoDbSettings__ConnectionString.");
    return;
}

if (string.IsNullOrWhiteSpace(databaseName))
{
    Console.WriteLine("ERROR: Khong tim thay MongoDbSettings__DatabaseName.");
    Console.WriteLine("Hay truyen --database hoac set env MongoDbSettings__DatabaseName.");
    return;
}

var client = new MongoClient(connectionString);
var database = client.GetDatabase(databaseName);
var scores = database.GetCollection<BsonDocument>("scores");

var filters = new List<FilterDefinition<BsonDocument>>();
if (!string.IsNullOrWhiteSpace(options.AssignmentId))
{
    filters.Add(BuildIdFilter("assignmentId", options.AssignmentId));
}

if (!string.IsNullOrWhiteSpace(options.ClassId))
{
    filters.Add(BuildIdFilter("classId", options.ClassId));
}

var baseFilter = filters.Count == 0
    ? Builders<BsonDocument>.Filter.Empty
    : filters.Count == 1
        ? filters[0]
        : Builders<BsonDocument>.Filter.And(filters);

var totalMatched = await scores.CountDocumentsAsync(baseFilter);

Console.WriteLine("=== SCORE SCHEMA BACKFILL ===");
Console.WriteLine($"Connection   : {MaskConnection(connectionString)}");
Console.WriteLine($"Database     : {databaseName}");
Console.WriteLine($"AssignmentId : {options.AssignmentId ?? "(all)"}");
Console.WriteLine($"ClassId      : {options.ClassId ?? "(all)"}");
Console.WriteLine($"Matched docs : {totalMatched}");
Console.WriteLine($"Mode         : {(options.Apply ? "APPLY (backup + replace)" : "DRY-RUN")}");

if (totalMatched == 0)
{
    Console.WriteLine("Khong co score document nao khop filter.");
    return;
}

var previewDocs = await scores
    .Find(baseFilter)
    .Sort(Builders<BsonDocument>.Sort.Ascending("_id"))
    .Limit(options.PreviewSize)
    .ToListAsync();

Console.WriteLine("--- Preview ---");
foreach (var doc in previewDocs)
{
    var normalized = NormalizeScoreDocument(doc, out var changes);
    Console.WriteLine($"id={ReadString(doc, "_id") ?? "(unknown)"} | changed={changes.Count} | changes={string.Join(" | ", changes.Take(4))}");
    if (changes.Count == 0)
    {
        Console.WriteLine("  no-op");
        continue;
    }

    Console.WriteLine($"  scoreValue={ReadString(normalized, "scoreValue") ?? "(null)"} | feedback={ReadString(normalized, "feedback") ?? "(null)"} | autoErrors={ReadArrayCount(normalized, "autoGradingErrors")} | autoTasks={ReadArrayCount(normalized, "autoGradingTaskResults")}");
}

if (!options.Apply)
{
    Console.WriteLine();
    Console.WriteLine("Dry-run complete. Chay them --apply de backup va backfill du lieu.");
    return;
}

var archiveCollectionName = string.IsNullOrWhiteSpace(options.ArchiveCollection)
    ? $"scores_archive_schema_backfill_{DateTime.UtcNow:yyyyMMdd_HHmmss}"
    : options.ArchiveCollection.Trim();
var archive = database.GetCollection<BsonDocument>(archiveCollectionName);
var migrationTag = $"score-schema-backfill-{DateTime.UtcNow:yyyyMMdd-HHmmss}";

Console.WriteLine();
Console.WriteLine($"Archive      : {archiveCollectionName}");
Console.WriteLine($"Migration tag: {migrationTag}");

ObjectId? lastId = null;
var scanned = 0;
var updated = 0;
var backedUp = 0;
var batchNo = 0;

while (true)
{
    var batchFilter = lastId.HasValue
        ? Builders<BsonDocument>.Filter.And(baseFilter, Builders<BsonDocument>.Filter.Gt("_id", lastId.Value))
        : baseFilter;

    var batch = await scores
        .Find(batchFilter)
        .Sort(Builders<BsonDocument>.Sort.Ascending("_id"))
        .Limit(options.BatchSize)
        .ToListAsync();

    if (batch.Count == 0)
    {
        break;
    }

    batchNo++;
    var now = DateTime.UtcNow;
    var backups = new List<BsonDocument>();
    var replaceModels = new List<WriteModel<BsonDocument>>();

    foreach (var document in batch)
    {
        scanned++;
        lastId = document["_id"].AsObjectId;

        var normalized = NormalizeScoreDocument(document, out var changes);
        if (changes.Count == 0)
        {
            continue;
        }

        var backupDoc = document.DeepClone().AsBsonDocument;
        backupDoc.Remove("_id");
        backupDoc["sourceScoreId"] = document["_id"];
        backupDoc["migratedAtUtc"] = now;
        backupDoc["migrationTag"] = migrationTag;
        backupDoc["changes"] = new BsonArray(changes);
        backups.Add(backupDoc);

        replaceModels.Add(new ReplaceOneModel<BsonDocument>(
            Builders<BsonDocument>.Filter.Eq("_id", document["_id"]),
            normalized));
    }

    if (backups.Count > 0)
    {
        await archive.InsertManyAsync(backups);
        backedUp += backups.Count;
        await scores.BulkWriteAsync(replaceModels);
        updated += replaceModels.Count;
    }

    Console.WriteLine($"Batch #{batchNo}: scanned={scanned}, updated={updated}, backedUp={backedUp}");
}

Console.WriteLine();
Console.WriteLine("=== DONE ===");
Console.WriteLine($"Scanned docs : {scanned}");
Console.WriteLine($"Updated docs : {updated}");
Console.WriteLine($"Backed up    : {backedUp}");
Console.WriteLine($"Archive      : {archiveCollectionName}");
Console.WriteLine($"Migration tag: {migrationTag}");

return;

static ScoreBackfillOptions ParseArgs(string[] args)
{
    var options = new ScoreBackfillOptions();
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
            case "--assignment-id":
                options.AssignmentId = ReadNextValue(args, ref i, "--assignment-id");
                break;
            case "--class-id":
                options.ClassId = ReadNextValue(args, ref i, "--class-id");
                break;
            case "--archive":
                options.ArchiveCollection = ReadNextValue(args, ref i, "--archive");
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

static BsonDocument NormalizeScoreDocument(BsonDocument source, out List<string> changes)
{
    var normalized = source.DeepClone().AsBsonDocument;
    changes = new List<string>();

    NormalizeObjectIdLikeField(normalized, "studentId", changes, required: true);
    NormalizeObjectIdLikeField(normalized, "assignmentId", changes, required: true);
    NormalizeObjectIdLikeField(normalized, "classId", changes, required: true);
    NormalizeObjectIdLikeField(normalized, "gradedBy", changes, required: false);
    NormalizeObjectIdLikeField(normalized, "createdBy", changes, required: false);
    NormalizeObjectIdLikeField(normalized, "updatedBy", changes, required: false);

    NormalizeScoreValue(normalized, changes);
    NormalizeFeedback(normalized, changes);
    NormalizeStringArrayField(normalized, "autoGradingErrors", changes);
    NormalizeTaskResults(normalized, changes);
    NormalizeDateField(normalized, "gradedAt", changes);
    NormalizeDateField(normalized, "createdAt", changes, fallbackUtcNow: true);
    NormalizeDateField(normalized, "updatedAt", changes);

    return normalized;
}

static void NormalizeObjectIdLikeField(BsonDocument document, string fieldName, List<string> changes, bool required)
{
    if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
    {
        if (required)
        {
            changes.Add($"{fieldName}: missing required field");
        }

        return;
    }

    if (value.BsonType == BsonType.ObjectId)
    {
        return;
    }

    if (value.BsonType == BsonType.String && ObjectId.TryParse(value.AsString, out var parsedString))
    {
        document[fieldName] = parsedString;
        changes.Add($"{fieldName}: string -> objectId");
        return;
    }

    if (ObjectId.TryParse(value.ToString(), out var parsed))
    {
        document[fieldName] = parsed;
        changes.Add($"{fieldName}: {value.BsonType} -> objectId");
    }
}

static void NormalizeScoreValue(BsonDocument document, List<string> changes)
{
    if (!document.TryGetValue("scoreValue", out var value) || value.IsBsonNull)
    {
        document["scoreValue"] = BsonNull.Value;
        changes.Add("scoreValue: missing -> null");
        return;
    }

    var normalizedValue = TryReadDouble(value);
    if (normalizedValue.HasValue)
    {
        var rounded = Math.Round(Math.Max(0d, normalizedValue.Value), 2, MidpointRounding.AwayFromZero);
        if (value.BsonType != BsonType.Double || Math.Abs(value.AsDouble - rounded) > 0.0001d)
        {
            document["scoreValue"] = rounded;
            changes.Add($"scoreValue: {value.BsonType} -> double");
        }

        return;
    }

    document["scoreValue"] = BsonNull.Value;
    changes.Add($"scoreValue: invalid {value.BsonType} -> null");
}

static void NormalizeFeedback(BsonDocument document, List<string> changes)
{
    if (!document.TryGetValue("feedback", out var value) || value.IsBsonNull)
    {
        document["feedback"] = BsonNull.Value;
        changes.Add("feedback: missing -> null");
        return;
    }

    var feedback = (value.ToString() ?? string.Empty).Trim();
    if (feedback.Length == 0)
    {
        if (!value.IsBsonNull)
        {
            document["feedback"] = BsonNull.Value;
            changes.Add("feedback: empty -> null");
        }

        return;
    }

    if (feedback.Length > 500)
    {
        feedback = feedback[..500];
    }

    if (value.BsonType != BsonType.String || !string.Equals(value.AsString, feedback, StringComparison.Ordinal))
    {
        document["feedback"] = feedback;
        changes.Add($"feedback: {value.BsonType} -> trimmed string");
    }
}

static void NormalizeStringArrayField(BsonDocument document, string fieldName, List<string> changes)
{
    var normalizedItems = ReadStringList(document.TryGetValue(fieldName, out var value) ? value : BsonNull.Value);
    var normalizedArray = new BsonArray(normalizedItems);

    if (!document.TryGetValue(fieldName, out value) || value.IsBsonNull || value.BsonType != BsonType.Array || !BsonValue.Equals(value, normalizedArray))
    {
        document[fieldName] = normalizedArray;
        changes.Add($"{fieldName}: normalized to string array ({normalizedItems.Count})");
    }
}

static void NormalizeTaskResults(BsonDocument document, List<string> changes)
{
    var normalizedTasks = new BsonArray();
    if (document.TryGetValue("autoGradingTaskResults", out var value) && value is { IsBsonNull: false, BsonType: BsonType.Array })
    {
        var index = 0;
        foreach (var item in value.AsBsonArray)
        {
            index++;
            if (item.IsBsonNull || item.BsonType != BsonType.Document)
            {
                continue;
            }

            var task = item.AsBsonDocument;
            var taskId = ReadString(task, "taskId")?.Trim();
            var taskName = ReadString(task, "taskName")?.Trim();
            if (string.IsNullOrWhiteSpace(taskId))
            {
                taskId = $"TASK-{index:00}";
            }

            if (string.IsNullOrWhiteSpace(taskName))
            {
                taskName = taskId;
            }

            var normalizedTask = new BsonDocument
            {
                { "taskId", taskId },
                { "taskName", taskName },
                { "score", Math.Round(Math.Max(0d, TryReadDouble(task.GetValue("score", BsonNull.Value)) ?? 0d), 4, MidpointRounding.AwayFromZero) },
                { "maxScore", NormalizeMaxScore(task.GetValue("maxScore", BsonNull.Value)) },
                { "isPassed", ReadBoolean(task.GetValue("isPassed", BsonNull.Value)) },
                { "details", new BsonArray(ReadStringList(task.GetValue("details", BsonNull.Value))) },
                { "errors", new BsonArray(ReadStringList(task.GetValue("errors", BsonNull.Value))) },
                { "fixActions", new BsonArray(ReadStringList(task.GetValue("fixActions", BsonNull.Value))) },
                { "displayIssues", NormalizeDisplayIssues(task.GetValue("displayIssues", BsonNull.Value)) }
            };

            normalizedTasks.Add(normalizedTask);
        }
    }

    if (!document.TryGetValue("autoGradingTaskResults", out value) || value.IsBsonNull || value.BsonType != BsonType.Array || !BsonValue.Equals(value, normalizedTasks))
    {
        document["autoGradingTaskResults"] = normalizedTasks;
        changes.Add($"autoGradingTaskResults: normalized ({normalizedTasks.Count})");
    }
}

static double NormalizeMaxScore(BsonValue value)
{
    var raw = TryReadDouble(value) ?? 1d;
    raw = raw <= 0 ? 1d : raw;
    return Math.Round(raw, 4, MidpointRounding.AwayFromZero);
}

static BsonArray NormalizeDisplayIssues(BsonValue value)
{
    var results = new BsonArray();
    if (value.IsBsonNull || value.BsonType != BsonType.Array)
    {
        return results;
    }

    foreach (var item in value.AsBsonArray)
    {
        if (item.IsBsonNull || item.BsonType != BsonType.Document)
        {
            continue;
        }

        var issue = item.AsBsonDocument;
        var heading = ReadString(issue, "heading")?.Trim() ?? string.Empty;
        var message = ReadString(issue, "message")?.Trim() ?? string.Empty;
        var fixAction = ReadString(issue, "fixAction")?.Trim() ?? string.Empty;
        if (heading.Length == 0 || message.Length == 0)
        {
            continue;
        }

        results.Add(new BsonDocument
        {
            { "heading", heading },
            { "message", message },
            { "fixAction", fixAction }
        });
    }

    return results;
}

static void NormalizeDateField(BsonDocument document, string fieldName, List<string> changes, bool fallbackUtcNow = false)
{
    if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
    {
        if (fallbackUtcNow)
        {
            document[fieldName] = DateTime.UtcNow;
            changes.Add($"{fieldName}: missing -> utcNow");
        }

        return;
    }

    var normalized = TryReadDateTime(value);
    if (normalized.HasValue)
    {
        if (value.BsonType != BsonType.DateTime)
        {
            document[fieldName] = normalized.Value;
            changes.Add($"{fieldName}: {value.BsonType} -> dateTime");
        }

        return;
    }

    if (fallbackUtcNow)
    {
        document[fieldName] = DateTime.UtcNow;
        changes.Add($"{fieldName}: invalid -> utcNow");
    }
}

static double? TryReadDouble(BsonValue value)
{
    if (value.IsBsonNull)
    {
        return null;
    }

    return value.BsonType switch
    {
        BsonType.Double => value.AsDouble,
        BsonType.Int32 => value.AsInt32,
        BsonType.Int64 => value.AsInt64,
        BsonType.Decimal128 => (double)value.AsDecimal128,
        BsonType.String when double.TryParse(value.AsString, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedInvariant) => parsedInvariant,
        BsonType.String when double.TryParse(value.AsString, out var parsedCurrent) => parsedCurrent,
        _ => null
    };
}

static DateTime? TryReadDateTime(BsonValue value)
{
    if (value.IsBsonNull)
    {
        return null;
    }

    return value.BsonType switch
    {
        BsonType.DateTime => value.ToUniversalTime(),
        BsonType.String when DateTime.TryParse(value.AsString, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsedInvariant) => parsedInvariant,
        BsonType.String when DateTime.TryParse(value.AsString, out var parsedCurrent) => parsedCurrent,
        _ => null
    };
}

static bool ReadBoolean(BsonValue value)
{
    if (value.IsBsonNull)
    {
        return false;
    }

    return value.BsonType switch
    {
        BsonType.Boolean => value.AsBoolean,
        BsonType.Int32 => value.AsInt32 != 0,
        BsonType.Int64 => value.AsInt64 != 0,
        BsonType.String when bool.TryParse(value.AsString, out var parsed) => parsed,
        _ => false
    };
}

static List<string> ReadStringList(BsonValue value)
{
    if (value.IsBsonNull)
    {
        return new List<string>();
    }

    if (value.BsonType != BsonType.Array)
    {
        var singleValue = (value.ToString() ?? string.Empty).Trim();
        return singleValue.Length == 0 ? new List<string>() : new List<string> { singleValue };
    }

    return value.AsBsonArray
        .Select(item => item.IsBsonNull ? null : item.ToString()?.Trim())
        .Where(item => !string.IsNullOrWhiteSpace(item))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .Take(100)
        .Select(item => item!)
        .ToList();
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

static int ReadArrayCount(BsonDocument document, string fieldName)
{
    if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull || value.BsonType != BsonType.Array)
    {
        return 0;
    }

    return value.AsBsonArray.Count;
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
    Console.WriteLine("  dotnet run --project tools/ScoreSchemaBackfill -- [options]");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  --assignment-id <id>   Chi backfill score cua mot assignment");
    Console.WriteLine("  --class-id <id>        Chi backfill score cua mot class");
    Console.WriteLine("  --connection <value>   Mongo connection string");
    Console.WriteLine("  --database <value>     Mongo database name");
    Console.WriteLine("  --archive <name>       Collection backup");
    Console.WriteLine("  --batch-size <n>       So document moi batch (default 200)");
    Console.WriteLine("  --preview <n>          So document preview (default 10)");
    Console.WriteLine("  --apply                Backup va ghi du lieu da normalize");
}

internal sealed class ScoreBackfillOptions
{
    public bool ShowHelp { get; set; }
    public bool Apply { get; set; }
    public string? ConnectionString { get; set; }
    public string? DatabaseName { get; set; }
    public string? AssignmentId { get; set; }
    public string? ClassId { get; set; }
    public string? ArchiveCollection { get; set; }
    public int BatchSize { get; set; } = 200;
    public int PreviewSize { get; set; } = 10;
}
