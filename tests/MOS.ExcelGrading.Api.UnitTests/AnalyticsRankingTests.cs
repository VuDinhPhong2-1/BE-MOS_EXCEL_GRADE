using System;
using System.Linq;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

public class AnalyticsRankingTests
{
    private static GradingAttempt Attempt(string? student, string? assignment, int time, bool passed, string id = "1") => new()
    {
        Id = id,
        StudentId = student,
        AssignmentId = assignment,
        ProjectEndpoint = "word/project01",
        ProjectId = "W01",
        GradedAt = new DateTime(2026, 1, 1).AddMinutes(time),
        TaskResults = new() { new() { TaskId = "T1", TaskName = "Task 1", IsPassed = passed } }
    };

    [Fact]
    public void SharedEndpointKeepsAssignmentsSeparateAndUsesLatestResult()
    {
        var rows = AnalyticsService.RankWeakTasks(new[] {
            Attempt("s1", "a", 1, false), Attempt("s1", "a", 2, true),
            Attempt("s1", "b", 3, false), Attempt("s2", "a", 1, false)
        }, 10, true);
        Assert.Equal(2, rows.Count);
        Assert.Equal("b", rows[0].AssignmentId);
        Assert.Equal(100d, rows[0].FailedRate);
        Assert.Equal(1, rows[0].AttemptCount);
        Assert.Equal(2, rows[1].AttemptCount);
        Assert.Equal(1, rows[1].FailedCount);
        Assert.Equal(50d, rows[1].FailedRate);
    }

    [Fact]
    public void TimestampTieUsesDescendingIdAndIgnoresUnattributedAttempts()
    {
        var rows = AnalyticsService.RankWeakTasks(new[] {
            Attempt("s1", "a", 1, false, "1"), Attempt("s1", "a", 1, true, "2"),
            Attempt(null, "a", 2, false), Attempt("s2", null, 2, false)
        }, 10, true);
        var row = Assert.Single(rows);
        Assert.Equal(1, row.AttemptCount);
        Assert.Equal(0, row.FailedCount);
        Assert.Empty(AnalyticsService.RankWeakTasks(new[] { Attempt(null, "a", 1, false) }, 10, true));
    }

    [Fact]
    public void TopIsAppliedAfterCombinedRankingAndScoreSaveIsExcluded()
    {
        var scoreSave = Attempt("s1", "c", 1, false);
        scoreSave.TaskResults[0].TaskId = "SCORE-SAVE";
        var row = Assert.Single(AnalyticsService.RankWeakTasks(new[] {
            Attempt("s1", "a", 1, true), Attempt("s1", "b", 1, false), scoreSave
        }, 1, true));
        Assert.Equal("b", row.AssignmentId);
    }

    [Fact]
    public void UnfilteredModePreservesLatestPerEndpoint()
    {
        var row = Assert.Single(AnalyticsService.RankWeakTasks(new[] {
            Attempt("s1", "a", 1, false), Attempt("s1", "b", 2, true)
        }, 10, false));
        Assert.Null(row.AssignmentId);
        Assert.Equal(0, row.FailedCount);
    }
}