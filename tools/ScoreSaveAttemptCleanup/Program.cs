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
var gradingAttempts = database.GetCollection<BsonDocument>("gradingAttempts");

var filters = new List<FilterDefinition<BsonDocument>>
{
    BuildScoreSaveFilter()
};

if (!string.IsNullOrWhiteSpace(options.ClassId))
{
    filters.Add(Builders<BsonDocument>.Filter.Eq("classId", options.ClassId));
}

if (!string.IsNullOrWhiteSpace(options.ProjectEndpoint))
{
    filters.Add(Builders<BsonDocument>.Filter.Eq("projectEndpoint", options.ProjectEndpoint));
}

var baseFilter = filters.Count == 1
    ? filters[0]
    : Builders<BsonDocument>.Filter.And(filters);

var totalMatch = await gradingAttempts.CountDocumentsAsync(baseFilter);

Console.WriteLine("=== SCORE-SAVE ATTEMPT CLEANUP ===");
Console.WriteLine($"Connection  : {MaskConnection(connectionString)}");
Console.WriteLine($"Database    : {databaseName}");
Console.WriteLine($"ClassId     : {options.ClassId ?? "(all)"}");
Console.WriteLine($"Endpoint    : {options.ProjectEndpoint ?? "(all)"}");
Console.WriteLine($"Matched docs: {totalMatch}");
Console.WriteLine($"Mode        : {(options.Apply ? "APPLY (backup + delete)" : "DRY-RUN")}");

var preview = await gradingAttempts
    .Find(baseFilter)
    .Sort(Builders<BsonDocument>.Sort.Descending("gradedAt"))
    .Limit(options.PreviewSize)
    .ToListAsync();

if (preview.Count > 0)
{
    Console.WriteLine("--- Preview ---");
    foreach (var doc in preview)
    {
        var id = doc.GetValue("_id", BsonNull.Value);
        var classId = doc.GetValue("classId", "").AsString;
        var assignmentId = doc.GetValue("assignmentId", "").AsString;
        var studentId = doc.GetValue("studentId", "").AsString;
        var endpoint = doc.GetValue("projectEndpoint", "").AsString;
        var gradedAt = doc.GetValue("gradedAt", BsonNull.Value);
        Console.WriteLine(
            $"id={id} | class={classId} | assignment={assignmentId} | student={studentId} | endpoint={endpoint} | gradedAt={gradedAt}");
    }
}

if (!options.Apply)
{
    Console.WriteLine();
    Console.WriteLine("Dry-run complete. Chạy thêm --apply để backup và xóa dữ liệu.");
    return;
}

if (totalMatch == 0)
{
    Console.WriteLine("Không có dữ liệu cần xử lý.");
    return;
}

var archiveCollectionName = string.IsNullOrWhiteSpace(options.ArchiveCollection)
    ? $"gradingAttempts_archive_score_save_{DateTime.UtcNow:yyyyMMdd_HHmmss}"
    : options.ArchiveCollection.Trim();
var archiveCollection = database.GetCollection<BsonDocument>(archiveCollectionName);
var migrationTag = $"score-save-cleanup-{DateTime.UtcNow:yyyyMMdd-HHmmss}";

Console.WriteLine();
Console.WriteLine($"Archive collection: {archiveCollectionName}");
Console.WriteLine($"Migration tag     : {migrationTag}");

var migratedCount = 0L;
var deletedCount = 0L;
var batchNo = 0;
ObjectId? lastId = null;
var scanSort = Builders<BsonDocument>.Sort.Ascending("_id");

while (true)
{
    var batchFilter = lastId.HasValue
        ? Builders<BsonDocument>.Filter.And(
            baseFilter,
            Builders<BsonDocument>.Filter.Gt("_id", lastId.Value))
        : baseFilter;

    var batch = await gradingAttempts
        .Find(batchFilter)
        .Sort(scanSort)
        .Limit(options.BatchSize)
        .ToListAsync();

    if (batch.Count == 0)
    {
        break;
    }

    batchNo++;

    var now = DateTime.UtcNow;
    var backupDocs = new List<BsonDocument>(batch.Count);
    var idsToDelete = new List<ObjectId>(batch.Count);

    foreach (var sourceDoc in batch)
    {
        var sourceId = sourceDoc["_id"].AsObjectId;
        idsToDelete.Add(sourceId);
        lastId = sourceId;

        var backupDoc = sourceDoc.DeepClone().AsBsonDocument;
        backupDoc.Remove("_id");
        backupDoc["sourceAttemptId"] = sourceId;
        backupDoc["migratedAtUtc"] = now;
        backupDoc["migrationTag"] = migrationTag;
        backupDocs.Add(backupDoc);
    }

    await archiveCollection.InsertManyAsync(backupDocs);

    var deleteResult = await gradingAttempts.DeleteManyAsync(
        Builders<BsonDocument>.Filter.In("_id", idsToDelete));

    migratedCount += backupDocs.Count;
    deletedCount += deleteResult.DeletedCount;

    Console.WriteLine(
        $"Batch #{batchNo}: backed up {backupDocs.Count}, deleted {deleteResult.DeletedCount}, total deleted={deletedCount}/{totalMatch}");
}

Console.WriteLine();
Console.WriteLine("=== DONE ===");
Console.WriteLine($"Backed up docs : {migratedCount}");
Console.WriteLine($"Deleted docs   : {deletedCount}");
Console.WriteLine($"Archive        : {archiveCollectionName}");
Console.WriteLine($"Migration tag  : {migrationTag}");

return;

static CleanupOptions ParseArgs(string[] args)
{
    var options = new CleanupOptions();
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
            case "--project-endpoint":
                options.ProjectEndpoint = ReadNextValue(args, ref i, "--project-endpoint");
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

static FilterDefinition<BsonDocument> BuildScoreSaveFilter()
{
    return new BsonDocument(
        "taskResults",
        new BsonDocument(
            "$elemMatch",
            new BsonDocument(
                "taskId",
                new BsonRegularExpression("^SCORE-SAVE$", "i"))));
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
        var candidate = Path.Combine(
            current.FullName,
            "MOS.ExcelGrading.API",
            "appsettings.Development.json");
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
    var atIndex = connectionString.IndexOf('@');
    if (atIndex <= 0)
    {
        return connectionString;
    }

    var protocolSep = connectionString.IndexOf("://", StringComparison.Ordinal);
    if (protocolSep < 0)
    {
        return connectionString;
    }

    return connectionString.Substring(0, protocolSep + 3) + "***:***" + connectionString[atIndex..];
}

static void PrintUsage()
{
    Console.WriteLine("Usage:");
    Console.WriteLine("  dotnet run --project tools/ScoreSaveAttemptCleanup/ScoreSaveAttemptCleanup.csproj -- [options]");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  --apply                      Run backup + delete (default is dry-run).");
    Console.WriteLine("  --connection <value>         Mongo connection string.");
    Console.WriteLine("  --database <value>           Mongo database name.");
    Console.WriteLine("  --archive <value>            Archive collection name.");
    Console.WriteLine("  --class-id <value>           Optional class filter.");
    Console.WriteLine("  --project-endpoint <value>   Optional endpoint filter.");
    Console.WriteLine("  --batch-size <n>             Batch size for apply mode (default 500).");
    Console.WriteLine("  --preview <n>                Preview rows in dry-run/apply header (default 10).");
    Console.WriteLine("  -h | --help                  Show this help.");
}

internal sealed class CleanupOptions
{
    public bool ShowHelp { get; set; }
    public bool Apply { get; set; }
    public string? ConnectionString { get; set; }
    public string? DatabaseName { get; set; }
    public string? ArchiveCollection { get; set; }
    public string? ClassId { get; set; }
    public string? ProjectEndpoint { get; set; }
    public int BatchSize { get; set; } = 500;
    public int PreviewSize { get; set; } = 10;
}
