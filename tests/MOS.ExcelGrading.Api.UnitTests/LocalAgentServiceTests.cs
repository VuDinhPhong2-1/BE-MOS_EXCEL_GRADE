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
using MOS.ExcelGrading.MockAgent.DTOs;
using MOS.ExcelGrading.MockAgent.Services;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class LocalAgentServiceTests
    {
        [Fact]
        public async Task NextProject_WhenAdvanceCompletesFinalProject_SetsCompletedCurrentState()
        {
            var examDir = Path.Combine(Path.GetTempPath(), "mos-agent-test-exam-" + Guid.NewGuid().ToString("N"));
            var templateDir = Path.Combine(Path.GetTempPath(), "mos-agent-test-template-" + Guid.NewGuid().ToString("N"));

            Directory.CreateDirectory(examDir);
            Directory.CreateDirectory(templateDir);

            try
            {
                var startState = new ExamSessionStateDto
                {
                    SessionId = "session-1",
                    CurrentProjectIndex = 0,
                    CurrentProjectNumber = 1,
                    TotalProjectCount = 1,
                    Status = "in_progress",
                    CurrentProjectStatus = "in_progress",
                    CurrentProject = new ExamSessionProjectBootstrapDto
                    {
                        Order = 1,
                        ProjectCode = "WORD_P01",
                        Subject = "word",
                        TemplateFileName = "Word_Project_01.docx",
                        GradingApiEndpoint = "word/project01"
                    }
                };

                var bootstrap = new ExamSessionProjectBootstrapDto
                {
                    Order = 1,
                    ProjectCode = "WORD_P01",
                    Subject = "word",
                    TemplateFileName = "Word_Project_01.docx",
                    GradingApiEndpoint = "word/project01"
                };

                var completedAdvance = new AdvanceExamSessionResponse
                {
                    IsCompleted = true,
                    State = new ExamSessionStateDto
                    {
                        SessionId = "session-1",
                        CurrentProjectIndex = 0,
                        CurrentProjectNumber = 1,
                        TotalProjectCount = 1,
                        CompletedProjectCount = 1,
                        Status = "completed",
                        CurrentProjectStatus = "graded"
                    }
                };

                var handler = new QueueHttpMessageHandler(new[]
                {
                    Response(HttpStatusCode.OK, startState),
                    Response(HttpStatusCode.OK, bootstrap),
                    Response(HttpStatusCode.OK, new { message = "saved" }),
                    Response(HttpStatusCode.OK, completedAdvance)
                });

                var httpClient = new HttpClient(handler);
                var configuration = new ConfigurationBuilder()
                    .AddInMemoryCollection(new Dictionary<string, string?>
                    {
                        ["Backend:BaseUrl"] = "http://localhost:9999",
                        ["MosPaths:TemplateDir"] = templateDir,
                        ["MosPaths:ExamDir"] = examDir
                    })
                    .Build();

                var logger = new Mock<ILogger<LocalAgentService>>();
                var service = new LocalAgentService(httpClient, configuration, logger.Object);

                await service.StartExamAsync(new StartExamAgentRequest
                {
                    PublicationToken = "token-1",
                    StudentName = "Nguyen Van A"
                });

                var state = await service.NextProjectAsync();
                var currentState = service.GetCurrentState();

                Assert.Equal("completed", state.Status);
                Assert.True(state.IsCompleted);
                Assert.True(state.IsCurrentProjectGraded);
                Assert.False(state.IsBusy);
                Assert.Null(state.LastError);

                Assert.Equal("completed", currentState.Status);
                Assert.True(currentState.IsCompleted);
                Assert.False(currentState.IsBusy);
            }
            finally
            {
                if (Directory.Exists(examDir))
                {
                    Directory.Delete(examDir, recursive: true);
                }

                if (Directory.Exists(templateDir))
                {
                    Directory.Delete(templateDir, recursive: true);
                }
            }
        }

        private static HttpResponseMessage Response(HttpStatusCode statusCode, object body)
        {
            return new HttpResponseMessage(statusCode)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(body),
                    Encoding.UTF8,
                    "application/json")
            };
        }

        private sealed class QueueHttpMessageHandler : HttpMessageHandler
        {
            private readonly Queue<HttpResponseMessage> _responses;

            public QueueHttpMessageHandler(IEnumerable<HttpResponseMessage> responses)
            {
                _responses = new Queue<HttpResponseMessage>(responses);
            }

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                if (_responses.Count == 0)
                {
                    throw new InvalidOperationException($"No mocked response left for {request.Method} {request.RequestUri}.");
                }

                return Task.FromResult(_responses.Dequeue());
            }
        }
    }
}
