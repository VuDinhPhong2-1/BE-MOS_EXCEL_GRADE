using System;
using System.IO;
using System.Text;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Driver;

Console.OutputEncoding = Encoding.UTF8;

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

// 1. Read project07.json
var jsonPath = Path.Combine(AppContext.BaseDirectory, "project07.json");
if (!File.Exists(jsonPath))
{
    jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "project07.json");
}
if (!File.Exists(jsonPath))
{
    jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "BE-MOS_EXCEL_GRADE", "tools", "Project07RuleSeeder", "project07.json");
}

Console.WriteLine($"Reading rules definition from {jsonPath}...");
var jsonContent = await File.ReadAllTextAsync(jsonPath, Encoding.UTF8);
var projectPayload = BsonSerializer.Deserialize<BsonDocument>(jsonContent);

var examId = projectPayload.GetValue("examId", "exam01").AsString;
var projectId = projectPayload.GetValue("projectId", "project07").AsString;
var projectCode = projectPayload.GetValue("projectCode", "project07").AsString;
var projectName = projectPayload.GetValue("projectName", "Excel Project 07 – Employee Bonuses & Parts Sales").AsString;
var maxScore = projectPayload.GetValue("maxScore", 142.0).AsDouble;
var tasks = projectPayload["tasks"].AsBsonArray;

// 2. Upsert into 'grading_rules' collection
var gradingRulesColl = db.GetCollection<BsonDocument>("grading_rules");
var gradingRuleFilter = Builders<BsonDocument>.Filter.And(
    Builders<BsonDocument>.Filter.Eq("examId", examId),
    Builders<BsonDocument>.Filter.Eq("projectId", projectId)
);

var existingGradingRule = await gradingRulesColl.Find(gradingRuleFilter).FirstOrDefaultAsync();
var gradingRuleDoc = new BsonDocument
{
    { "subject", "EXCEL" },
    { "examId", examId },
    { "projectId", projectId },
    { "projectCode", projectCode },
    { "projectName", projectName },
    { "maxScore", maxScore },
    { "tasks", tasks },
    { "updatedAt", DateTime.UtcNow }
};

if (existingGradingRule != null)
{
    Console.WriteLine($"Updating existing grading_rules doc {existingGradingRule["_id"]} for {examId}/{projectId}...");
    gradingRuleDoc["_id"] = existingGradingRule["_id"];
    await gradingRulesColl.ReplaceOneAsync(gradingRuleFilter, gradingRuleDoc, new ReplaceOptions { IsUpsert = true });
}
else
{
    Console.WriteLine($"Inserting new grading_rules doc for {examId}/{projectId}...");
    await gradingRulesColl.InsertOneAsync(gradingRuleDoc);
}
Console.WriteLine($"Successfully seeded 'grading_rules' collection for {examId}/{projectId}.");

// 3. Upsert into 'grading_rule_projects' and 'grading_rule_sets' if collections exist
var ruleSetsColl = db.GetCollection<BsonDocument>("grading_rule_sets");
var ruleProjectsColl = db.GetCollection<BsonDocument>("grading_rule_projects");

var activeSetFilter = Builders<BsonDocument>.Filter.And(
    Builders<BsonDocument>.Filter.Eq("subject", "excel"),
    Builders<BsonDocument>.Filter.Eq("isActive", true)
);

var activeRuleSet = await ruleSetsColl.Find(activeSetFilter).FirstOrDefaultAsync();
if (activeRuleSet == null)
{
    Console.WriteLine("Notice: No active ruleset with subject 'excel' and isActive = true. Searching for any 'excel' ruleset...");
    activeRuleSet = await ruleSetsColl.Find(Builders<BsonDocument>.Filter.Eq("subject", "excel")).FirstOrDefaultAsync();
}

if (activeRuleSet != null)
{
    var ruleSetId = activeRuleSet["_id"].AsObjectId;
    var subject = activeRuleSet.GetValue("subject", "excel").AsString;
    var version = activeRuleSet.GetValue("version", "v1.0").AsString;
    Console.WriteLine($"Found Excel ruleset: Id={ruleSetId}, subject={subject}, version={version}");

    var projectFilter = Builders<BsonDocument>.Filter.And(
        Builders<BsonDocument>.Filter.Eq("ruleSetId", ruleSetId),
        Builders<BsonDocument>.Filter.Eq("projectCode", projectCode)
    );

    var existingProject = await ruleProjectsColl.Find(projectFilter).FirstOrDefaultAsync();

    var projectDoc = new BsonDocument
    {
        { "ruleSetId", ruleSetId },
        { "subject", subject },
        { "version", version },
        { "isActive", true },
        { "projectCode", projectCode },
        { "sortOrder", 7 },
        { "projectName", projectName },
        { "maxScore", maxScore },
        { "tasks", tasks }
    };

    if (existingProject != null)
    {
        Console.WriteLine($"Updating existing grading_rule_projects doc {existingProject["_id"]} for {projectCode}...");
        projectDoc["_id"] = existingProject["_id"];
        await ruleProjectsColl.ReplaceOneAsync(Builders<BsonDocument>.Filter.Eq("_id", existingProject["_id"]), projectDoc);
    }
    else
    {
        Console.WriteLine($"Inserting new grading_rule_projects doc for {projectCode}...");
        await ruleProjectsColl.InsertOneAsync(projectDoc);
    }

    if (activeRuleSet.Contains("projects") && activeRuleSet["projects"].IsBsonArray)
    {
        var projectsArray = activeRuleSet["projects"].AsBsonArray;
        int existingIndex = -1;
        for (int i = 0; i < projectsArray.Count; i++)
        {
            if (projectsArray[i].IsBsonDocument &&
                projectsArray[i].AsBsonDocument.GetValue("projectCode", "").AsString == projectCode)
            {
                existingIndex = i;
                break;
            }
        }

        var embeddedDoc = new BsonDocument
        {
            { "projectCode", projectCode },
            { "projectName", projectName },
            { "maxScore", maxScore },
            { "tasks", tasks }
        };

        if (existingIndex >= 0)
        {
            Console.WriteLine($"Updating embedded {projectCode} at index {existingIndex} in grading_rule_sets...");
            projectsArray[existingIndex] = embeddedDoc;
        }
        else
        {
            Console.WriteLine($"Appending embedded {projectCode} to grading_rule_sets...");
            projectsArray.Add(embeddedDoc);
        }

        await ruleSetsColl.ReplaceOneAsync(Builders<BsonDocument>.Filter.Eq("_id", ruleSetId), activeRuleSet);
        Console.WriteLine("Successfully updated embedded project in 'grading_rule_sets'.");
    }
}
else
{
    Console.WriteLine("Notice: 'grading_rule_sets' Excel ruleset not found, skipped secondary update.");
}

Console.WriteLine("Done! Successfully seeded Project 07 rules into MongoDB.");
