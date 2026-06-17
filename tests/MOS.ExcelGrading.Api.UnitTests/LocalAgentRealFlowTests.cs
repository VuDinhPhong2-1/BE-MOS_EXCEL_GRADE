using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Moq;
using MOS.ExcelGrading.LocalAgent.DTOs;
using MOS.ExcelGrading.LocalAgent.Services;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class LocalAgentRealFlowTests
    {
        [Fact]
        public async Task GradeAsync_WithRealApi_UsesStudentFileMultipartField()
        {
            var tempFile = CreateTempFile(".xlsx");

            try
            {
                var handler = new RecordingHttpMessageHandler(_ =>
                    JsonResponse(HttpStatusCode.OK, new
                    {
                        isSuccess = true,
                        score = 125,
                        maxScore = 125,
                        feedback = "graded"
                    }));

                var service = new LocalGradingService(
                    new HttpClient(handler),
                    BuildConfiguration(new Dictionary<string, string?>
                    {
                        ["Grading:UseMockGrading"] = "false",
                        ["Grading:BaseUrl"] = "https://localhost:7001"
                    }),
                    Mock.Of<ILogger<LocalGradingService>>());

                var result = await service.GradeAsync(new LocalAgentState
                {
                    SessionId = "session-1",
                    ProjectCode = "EXCEL_P01",
                    Subject = "excel",
                    GradingApiEndpoint = "/api/gradings/excel/project-1",
                    WorkingFilePath = tempFile
                });

                Assert.True(result.IsSuccess);
                Assert.NotNull(handler.LastRequest);
                var body = handler.LastRequestBody ?? string.Empty;
                Assert.Contains("name=studentFile", body);
                Assert.DoesNotContain("name=file", body);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public async Task GradeAsync_WhenGradingApiReturnsErrorStatus_ReturnsApiErrorMessage()
        {
            var tempFile = CreateTempFile(".xlsx");

            try
            {
                var handler = new RecordingHttpMessageHandler(_ =>
                    new HttpResponseMessage(HttpStatusCode.InternalServerError)
                    {
                        Content = new StringContent("boom", Encoding.UTF8, "text/plain")
                    });

                var service = new LocalGradingService(
                    new HttpClient(handler),
                    BuildConfiguration(new Dictionary<string, string?>
                    {
                        ["Grading:UseMockGrading"] = "false",
                        ["Grading:BaseUrl"] = "https://localhost:7001"
                    }),
                    Mock.Of<ILogger<LocalGradingService>>());

                var result = await service.GradeAsync(new LocalAgentState
                {
                    SessionId = "session-1",
                    ProjectCode = "EXCEL_P01",
                    Subject = "excel",
                    GradingApiEndpoint = "/api/gradings/excel/project-1",
                    WorkingFilePath = tempFile
                });

                Assert.False(result.IsSuccess);
                Assert.Equal("Grading API lỗi 500: boom", result.ErrorMessage);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public async Task SubmitCurrentProject_WhenGradingSucceeds_SetsGradedState()
        {
            var tempFile = CreateTempFile(".xlsx");

            try
            {
                var handler = new RecordingHttpMessageHandler(_ =>
                    JsonResponse(HttpStatusCode.OK, new { message = "saved" }));

                var service = CreateLocalAgentService(
                    httpClient: new HttpClient(handler),
                    state: new LocalAgentState
                    {
                        SessionId = "session-1",
                        PublicationToken = "token-1",
                        ProjectCode = "EXCEL_P01",
                        Subject = "excel",
                        GradingApiEndpoint = "/api/gradings/excel/project-1",
                        WorkingFilePath = tempFile,
                        Status = "in_progress",
                        IsCurrentProjectGraded = false
                    },
                    gradingResult: new LocalGradingResult
                    {
                        IsSuccess = true,
                        Score = 100,
                        MaxScore = 125,
                        Feedback = "ok"
                    });

                var state = await service.SubmitCurrentProjectAsync();

                Assert.Equal("graded", state.Status);
                Assert.True(state.IsCurrentProjectGraded);
                Assert.False(state.IsBusy);
                Assert.Null(state.LastError);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public async Task SubmitCurrentProject_WhenWorkingFileMissing_ThrowsExpectedMessage()
        {
            var missingFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            var handler = new RecordingHttpMessageHandler(_ =>
                throw new InvalidOperationException("HTTP should not be called when working file is missing."));

            var service = CreateLocalAgentService(
                httpClient: new HttpClient(handler),
                state: new LocalAgentState
                {
                    SessionId = "session-1",
                    PublicationToken = "token-1",
                    ProjectCode = "EXCEL_P01",
                    Subject = "excel",
                    GradingApiEndpoint = "/api/gradings/excel/project-1",
                    WorkingFilePath = missingFile,
                    Status = "in_progress"
                },
                gradingResult: new LocalGradingResult { IsSuccess = true, Score = 100, MaxScore = 125 });

            var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => service.SubmitCurrentProjectAsync());
            Assert.Contains("Không tìm thấy working file", ex.Message);

            var currentState = service.GetCurrentState();
            Assert.Equal("error", currentState.Status);
            Assert.False(currentState.IsCurrentProjectGraded);
            Assert.False(currentState.IsBusy);
        }

        [Fact]
        public async Task SubmitCurrentProject_WhenFileRecentlyModified_RequiresForceSubmit()
        {
            var tempFile = CreateTempFile(".xlsx");

            try
            {
                var handler = new RecordingHttpMessageHandler(_ =>
                    JsonResponse(HttpStatusCode.OK, new { message = "saved" }));

                var service = CreateLocalAgentService(
                    httpClient: new HttpClient(handler),
                    state: new LocalAgentState
                    {
                        SessionId = "session-1",
                        PublicationToken = "token-1",
                        ProjectCode = "EXCEL_P01",
                        Subject = "excel",
                        GradingApiEndpoint = "/api/gradings/excel/project-1",
                        WorkingFilePath = tempFile,
                        Status = "in_progress"
                    },
                    gradingResult: new LocalGradingResult
                    {
                        IsSuccess = true,
                        Score = 100,
                        MaxScore = 125,
                        Feedback = "ok"
                    },
                    configurationOverrides: new Dictionary<string, string?>
                    {
                        ["PreSubmit:WarnIfFileRecentlyModifiedSeconds"] = "10"
                    });

                var ex = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                    service.SubmitCurrentProjectAsync(new SubmitCurrentProjectAgentRequest
                    {
                        ConfirmSaved = true
                    }));

                Assert.Contains("Không thể submit:", ex.Message);
                Assert.Contains("Có cảnh báo trước khi submit. Gửi ForceSubmit=true nếu vẫn muốn submit.", ex.Message);
                Assert.Empty(handler.Requests);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public async Task SubmitCurrentProject_WhenFileRecentlyModified_ForceSubmitStillSubmits()
        {
            var tempFile = CreateTempFile(".xlsx");

            try
            {
                var handler = new RecordingHttpMessageHandler(_ =>
                    JsonResponse(HttpStatusCode.OK, new { message = "saved" }));

                var service = CreateLocalAgentService(
                    httpClient: new HttpClient(handler),
                    state: new LocalAgentState
                    {
                        SessionId = "session-1",
                        PublicationToken = "token-1",
                        ProjectCode = "EXCEL_P01",
                        Subject = "excel",
                        GradingApiEndpoint = "/api/gradings/excel/project-1",
                        WorkingFilePath = tempFile,
                        Status = "in_progress"
                    },
                    gradingResult: new LocalGradingResult
                    {
                        IsSuccess = true,
                        Score = 100,
                        MaxScore = 125,
                        Feedback = "ok"
                    },
                    configurationOverrides: new Dictionary<string, string?>
                    {
                        ["PreSubmit:WarnIfFileRecentlyModifiedSeconds"] = "10"
                    });

                var state = await service.SubmitCurrentProjectAsync(new SubmitCurrentProjectAgentRequest
                {
                    ConfirmSaved = true,
                    ForceSubmit = true
                });

                Assert.Equal("graded", state.Status);
                Assert.True(state.IsCurrentProjectGraded);
                Assert.Single(handler.Requests);
                Assert.Contains("/score-upload", handler.Requests[0].RequestUri!.AbsolutePath);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public async Task NextProject_WhenSubmitFails_DoesNotCallAdvance_AndKeepsCurrentProject()
        {
            var tempFile = CreateTempFile(".xlsx");

            try
            {
                var handler = new RecordingHttpMessageHandler(_ =>
                    JsonResponse(HttpStatusCode.OK, new { message = "saved" }));

                var service = CreateLocalAgentService(
                    httpClient: new HttpClient(handler),
                    state: new LocalAgentState
                    {
                        SessionId = "session-1",
                        PublicationToken = "token-1",
                        ProjectCode = "EXCEL_P01",
                        Subject = "excel",
                        GradingApiEndpoint = "/api/gradings/excel/project-1",
                        WorkingFilePath = tempFile,
                        Status = "in_progress",
                        CurrentProjectNumber = 1,
                        TotalProjectCount = 2
                    },
                    gradingResult: new LocalGradingResult
                    {
                        IsSuccess = false,
                        ErrorMessage = "Lỗi gọi Grading API: server down"
                    });

                var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => service.NextProjectAsync());
                Assert.Equal("Lỗi gọi Grading API: server down", ex.Message);

                Assert.Empty(handler.Requests);
                Assert.DoesNotContain(handler.Requests, request => request.RequestUri!.AbsolutePath.EndsWith("/advance", StringComparison.Ordinal));

                var currentState = service.GetCurrentState();
                Assert.Equal("EXCEL_P01", currentState.ProjectCode);
                Assert.Equal(1, currentState.CurrentProjectNumber);
                Assert.False(currentState.IsCurrentProjectGraded);
                Assert.False(currentState.IsBusy);
                Assert.Equal("error", currentState.Status);
                Assert.Equal("Lỗi gọi Grading API: server down", currentState.LastError);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        private static LocalAgentService CreateLocalAgentService(
            HttpClient httpClient,
            LocalAgentState state,
            LocalGradingResult gradingResult,
            Dictionary<string, string?>? configurationOverrides = null)
        {
            var values = new Dictionary<string, string?>
            {
                ["Backend:BaseUrl"] = "https://localhost:7001",
                ["PreSubmit:WarnIfFileRecentlyModifiedSeconds"] = "0"
            };

            if (configurationOverrides != null)
            {
                foreach (var entry in configurationOverrides)
                {
                    values[entry.Key] = entry.Value;
                }
            }

            var configuration = BuildConfiguration(values);

            return new LocalAgentService(
                httpClient,
                configuration,
                Mock.Of<ILogger<LocalAgentService>>(),
                new StubLocalFileService(),
                new InMemoryStateStore(state),
                new StubLocalGradingService(gradingResult),
                new PreSubmitCheckService(
                    configuration,
                    Mock.Of<ILogger<PreSubmitCheckService>>()));
        }

        private static IConfiguration BuildConfiguration(Dictionary<string, string?> values)
        {
            return new ConfigurationBuilder()
                .AddInMemoryCollection(values)
                .Build();
        }

        private static HttpResponseMessage JsonResponse(HttpStatusCode statusCode, object body)
        {
            return new HttpResponseMessage(statusCode)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(body),
                    Encoding.UTF8,
                    "application/json")
            };
        }

        private static string CreateTempFile(string extension)
        {
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
            File.WriteAllText(path, "test");
            return path;
        }

        private sealed class RecordingHttpMessageHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> _responseFactory;

            public RecordingHttpMessageHandler(Func<HttpRequestMessage, HttpResponseMessage> responseFactory)
            {
                _responseFactory = responseFactory;
            }

            public List<HttpRequestMessage> Requests { get; } = new();

            public HttpRequestMessage? LastRequest => Requests.Count == 0 ? null : Requests[^1];

            public string? LastRequestBody { get; private set; }

            protected override async Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                Requests.Add(request);
                LastRequestBody = request.Content == null
                    ? null
                    : await request.Content.ReadAsStringAsync(cancellationToken);
                return _responseFactory(request);
            }
        }

        private sealed class InMemoryStateStore : ILocalAgentStateStore
        {
            private LocalAgentState _state;

            public InMemoryStateStore(LocalAgentState state)
            {
                _state = state;
            }

            public LocalAgentState Load() => _state;

            public void Save(LocalAgentState state) => _state = state;

            public void Clear() => _state = new LocalAgentState();
        }

        private sealed class StubLocalFileService : ILocalFileService
        {
            public string PrepareWorkingFile(ExamSessionProjectBootstrapDto bootstrap) => throw new NotImplementedException();

            public void OpenWorkingFile(string workingFilePath)
            {
            }
        }

        private sealed class StubLocalGradingService : ILocalGradingService
        {
            private readonly LocalGradingResult _result;

            public StubLocalGradingService(LocalGradingResult result)
            {
                _result = result;
            }

            public Task<LocalGradingResult> GradeAsync(LocalAgentState state) => Task.FromResult(_result);
        }
    }
}
