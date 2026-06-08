using System;
using System.Collections.Generic;
using System.Reflection;
using MongoDB.Bson;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class ScoreControllerAndServiceTests
    {
        [Fact]
        public void TryMapScoreDocument_MapsLegacyPayloadWithoutDroppingScoreFields()
        {
            var document = new BsonDocument
            {
                { "_id", ObjectId.GenerateNewId() },
                { "studentId", ObjectId.GenerateNewId() },
                { "assignmentId", ObjectId.GenerateNewId().ToString() },
                { "classId", ObjectId.GenerateNewId().ToString() },
                { "scoreValue", "8.5" },
                { "autoGradingErrors", "Khong tim thay file" },
                {
                    "autoGradingTaskResults",
                    new BsonArray
                    {
                        new BsonDocument
                        {
                            { "taskId", "task-compile" },
                            { "taskName", "Compile" },
                            { "score", "1" },
                            { "maxScore", 1 },
                            { "isPassed", "true" },
                            { "details", "Compiled successfully" },
                            { "errors", new BsonArray() },
                            { "fixActions", new BsonArray() },
                            {
                                "displayIssues",
                                new BsonArray
                                {
                                    new BsonDocument
                                    {
                                        { "heading", "Sheet1" },
                                        { "message", "Correct format" },
                                        { "fixAction", "No action needed" }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            var method = typeof(ScoreService).GetMethod(
                "TryMapScoreDocument",
                BindingFlags.NonPublic | BindingFlags.Static);

            Assert.NotNull(method);

            var args = new object?[] { document, null };
            var success = (bool)method!.Invoke(null, args)!;

            Assert.True(success);
            var mappedScore = Assert.IsType<Score>(args[1]);
            Assert.Equal(8.5d, mappedScore.ScoreValue);
            Assert.Null(mappedScore.Feedback);
            Assert.Equal(new[] { "Khong tim thay file" }, mappedScore.AutoGradingErrors);

            var taskResult = Assert.Single(mappedScore.AutoGradingTaskResults);
            Assert.True(taskResult.IsPassed);
            Assert.Equal(1d, taskResult.Score);
            Assert.Equal("Correct format", Assert.Single(taskResult.DisplayIssues).Message);
        }

        [Fact]
        public void InspectScoreDocumentShape_FlagsInvalidNestedTaskResultFields()
        {
            var document = new BsonDocument
            {
                { "_id", ObjectId.GenerateNewId() },
                { "studentId", ObjectId.GenerateNewId() },
                { "assignmentId", ObjectId.GenerateNewId() },
                { "classId", ObjectId.GenerateNewId() },
                {
                    "autoGradingTaskResults",
                    new BsonArray
                    {
                        new BsonDocument
                        {
                            { "taskId", "task-compile" },
                            { "score", new BsonDocument("bad", 1) },
                            { "isPassed", "maybe" },
                            { "displayIssues", new BsonArray { "bad-issue" } }
                        }
                    }
                }
            };

            var method = typeof(ScoreService).GetMethod(
                "InspectScoreDocumentShape",
                BindingFlags.NonPublic | BindingFlags.Static);

            Assert.NotNull(method);

            var problems = Assert.IsType<List<string>>(method!.Invoke(null, new object[] { document })!);

            Assert.Contains(problems, item => item.Contains("autoGradingTaskResults[0].score", StringComparison.Ordinal));
            Assert.Contains(problems, item => item.Contains("autoGradingTaskResults[0].isPassed", StringComparison.Ordinal));
            Assert.Contains(problems, item => item.Contains("autoGradingTaskResults[0].displayIssues[0]", StringComparison.Ordinal));
        }
    }
}
