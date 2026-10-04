using System;
using System.IO;
using System.Text;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Driver;

Console.OutputEncoding = Encoding.UTF8;

var connectionString = Environment.GetEnvironmentVariable("MongoDbSettings__ConnectionString") 
    ?? "mongodb+srv://vudinhphong261001_db_user:Alo0906412535@cluster0.k98fhem.mongodb.net/MOS?retryWrites=true&w=majority&ssl=true";
var databaseName = "MOS";

Console.WriteLine("Connecting to MongoDB...");
var client = new MongoClient(connectionString);
var db = client.GetDatabase(databaseName);

var ruleSetsColl = db.GetCollection<BsonDocument>("grading_rule_sets");
var ruleProjectsColl = db.GetCollection<BsonDocument>("grading_rule_projects");

// 1. Read project05.json
var jsonPath = Path.Combine(AppContext.BaseDirectory, "project05.json");
if (!File.Exists(jsonPath))
{
    jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "project05.json");
}
Console.WriteLine($"Reading {jsonPath}...");
var jsonContent = await File.ReadAllTextAsync(jsonPath, Encoding.UTF8);
var projectPayload = BsonSerializer.Deserialize<BsonDocument>(jsonContent);

// 2. Find active ruleset for excel
var activeSetFilter = Builders<BsonDocument>.Filter.And(
    Builders<BsonDocument>.Filter.Eq("subject", "excel"),
    Builders<BsonDocument>.Filter.Eq("isActive", true)
);

var activeRuleSet = await ruleSetsColl.Find(activeSetFilter).FirstOrDefaultAsync();
if (activeRuleSet == null)
{
    Console.WriteLine("Warning: No active ruleset with subject 'excel' and isActive = true. Searching for any 'excel' ruleset...");
    activeRuleSet = await ruleSetsColl.Find(Builders<BsonDocument>.Filter.Eq("subject", "excel")).FirstOrDefaultAsync();
}

if (activeRuleSet == null)
{
    Console.WriteLine("Error: No ruleset found for subject 'excel'.");
    return;
}

var ruleSetId = activeRuleSet["_id"].AsObjectId;
var subject = activeRuleSet.GetValue("subject", "excel").AsString;
var version = activeRuleSet.GetValue("version", "v1.0").AsString;
Console.WriteLine($"Found Excel ruleset: Id={ruleSetId}, subject={subject}, version={version}");

// 3. Upsert into grading_rule_projects
var projectCode = projectPayload["projectCode"].AsString;
var projectName = projectPayload["projectName"].AsString;
var maxScore = projectPayload["maxScore"].AsDouble;
var tasks = projectPayload["tasks"].AsBsonArray;

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
    { "sortOrder", 5 },
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

// 4. Update embedded projects in grading_rule_sets if present
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
}

Console.WriteLine("Done! Successfully seeded Project 05 rules.");
