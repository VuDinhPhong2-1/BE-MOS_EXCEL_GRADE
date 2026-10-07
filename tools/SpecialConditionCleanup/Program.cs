using MongoDB.Bson;
using MongoDB.Driver;

var connectionString = Environment.GetEnvironmentVariable("MongoDbSettings__ConnectionString") ?? "";
var databaseName = Environment.GetEnvironmentVariable("MongoDbSettings__DatabaseName") ?? "MOS";

for (int i = 0; i < args.Length; i++)
{
    if (args[i] == "--connectionString" && i + 1 < args.Length)
    {
        connectionString = args[i + 1];
    }
    else if (args[i] == "--database" && i + 1 < args.Length)
    {
        databaseName = args[i + 1];
    }
}

if (string.IsNullOrWhiteSpace(connectionString))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Lỗi: Thiếu chuỗi kết nối MongoDB. Hãy truyền qua biến môi trường MongoDbSettings__ConnectionString hoặc tham số --connectionString.");
    Console.ResetColor();
    return;
}

Console.WriteLine($"Connecting to MongoDB database '{databaseName}'...");
var client = new MongoClient(connectionString);
var db = client.GetDatabase(databaseName);

var projectColl = db.GetCollection<BsonDocument>("grading_rule_projects");
var setColl = db.GetCollection<BsonDocument>("grading_rule_sets");

string GetActiveConfigProp(string type)
{
    return type switch
    {
        "pictureBullet" => "config",
        "insertedImage" => "imageInsertConfig",
        _ => $"{type}Config"
    };
}

void CleanSpecialCondition(BsonDocument sc)
{
    if (!sc.Contains("type") || !sc["type"].IsString) return;
    var type = sc["type"].AsString;
    var activeConfig = GetActiveConfigProp(type);

    var toRemove = new List<string>();
    foreach (var element in sc.Elements)
    {
        var name = element.Name;
        if (name == "type" || name == "score" || name == "feedback") continue;
        if (element.Value.IsBsonNull)
        {
            toRemove.Add(name);
            continue;
        }
        if (name != activeConfig && (name == "config" || name.EndsWith("Config", StringComparison.OrdinalIgnoreCase)))
        {
            toRemove.Add(name);
        }
    }

    foreach (var r in toRemove)
    {
        Console.WriteLine($"  Removing field: {r}");
        sc.Remove(r);
    }
}

// 1. Clean grading_rule_projects
var projectDocs = await projectColl.Find(Builders<BsonDocument>.Filter.Empty).ToListAsync();
Console.WriteLine($"Found {projectDocs.Count} grading_rule_projects");

foreach (var doc in projectDocs)
{
    var id = doc["_id"];
    var projectCode = doc.Contains("projectCode") ? doc["projectCode"].AsString : "?";
    bool modified = false;

    if (doc.Contains("tasks") && doc["tasks"].IsBsonArray)
    {
        foreach (var taskVal in doc["tasks"].AsBsonArray)
        {
            if (taskVal.IsBsonDocument && taskVal.AsBsonDocument.Contains("specialCondition"))
            {
                var scVal = taskVal.AsBsonDocument["specialCondition"];
                if (scVal.IsBsonDocument)
                {
                    var sc = scVal.AsBsonDocument;
                    var type = sc.Contains("type") ? sc["type"].AsString : "";
                    Console.WriteLine($"Project {projectCode} task has specialCondition type={type}");
                    var beforeCount = sc.ElementCount;
                    CleanSpecialCondition(sc);
                    if (sc.ElementCount != beforeCount)
                    {
                        modified = true;
                    }
                }
            }
        }
    }

    if (modified)
    {
        Console.WriteLine($"Updating grading_rule_projects doc {id} ({projectCode})...");
        await projectColl.ReplaceOneAsync(Builders<BsonDocument>.Filter.Eq("_id", id), doc);
    }
}

// 2. Clean grading_rule_sets
var setDocs = await setColl.Find(Builders<BsonDocument>.Filter.Empty).ToListAsync();
Console.WriteLine($"Found {setDocs.Count} grading_rule_sets");

foreach (var doc in setDocs)
{
    var id = doc["_id"];
    var examId = doc.Contains("examId") ? doc["examId"].AsString : "?";
    bool modified = false;

    if (doc.Contains("projects") && doc["projects"].IsBsonArray)
    {
        foreach (var pVal in doc["projects"].AsBsonArray)
        {
            if (!pVal.IsBsonDocument) continue;
            var pDoc = pVal.AsBsonDocument;
            var pCode = pDoc.Contains("projectCode") ? pDoc["projectCode"].AsString : "?";

            if (pDoc.Contains("tasks") && pDoc["tasks"].IsBsonArray)
            {
                foreach (var taskVal in pDoc["tasks"].AsBsonArray)
                {
                    if (taskVal.IsBsonDocument && taskVal.AsBsonDocument.Contains("specialCondition"))
                    {
                        var scVal = taskVal.AsBsonDocument["specialCondition"];
                        if (scVal.IsBsonDocument)
                        {
                            var sc = scVal.AsBsonDocument;
                            var type = sc.Contains("type") ? sc["type"].AsString : "";
                            Console.WriteLine($"Set {examId} project {pCode} task has specialCondition type={type}");
                            var beforeCount = sc.ElementCount;
                            CleanSpecialCondition(sc);
                            if (sc.ElementCount != beforeCount)
                            {
                                modified = true;
                            }
                        }
                    }
                }
            }
        }
    }

    if (modified)
    {
        Console.WriteLine($"Updating grading_rule_sets doc {id} ({examId})...");
        await setColl.ReplaceOneAsync(Builders<BsonDocument>.Filter.Eq("_id", id), doc);
    }
}

Console.WriteLine("Cleanup completed successfully!");


