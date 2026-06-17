using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Moq;
using MOS.ExcelGrading.API.Controllers;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class AgentExamSessionsControllerTests
    {
        [Fact]
        public async Task Advance_WhenSessionIsAdvancing_ReturnsBadRequestWithBusyMessage()
        {
            var examSessionService = new Mock<IExamSessionService>();
            var logger = new Mock<ILogger<AgentExamSessionsController>>();

            examSessionService
                .Setup(x => x.AdvanceAsync("session-1"))
                .ThrowsAsync(new ArgumentException("Session đang chuyển project, vui lòng chờ."));

            var controller = new AgentExamSessionsController(
                examSessionService.Object,
                logger.Object);

            var result = await controller.Advance("session-1");

            var badRequest = Assert.IsType<BadRequestObjectResult>(result);
            Assert.Equal(
                "Session đang chuyển project, vui lòng chờ.",
                GetMessage(badRequest.Value));
        }

        [Fact]
        public async Task UploadScore_WhenSessionIsAdvancing_ReturnsBadRequestWithAdvancingMessage()
        {
            var examSessionService = new Mock<IExamSessionService>();
            var logger = new Mock<ILogger<AgentExamSessionsController>>();

            examSessionService
                .Setup(x => x.UploadScoreAsync("session-1", "WORD_P01", It.IsAny<ScoreUploadRequest>()))
                .ThrowsAsync(new ArgumentException("Session đang chuyển project, không thể upload điểm lúc này."));

            var controller = new AgentExamSessionsController(
                examSessionService.Object,
                logger.Object);

            var result = await controller.UploadScore(
                "session-1",
                "WORD_P01",
                new ScoreUploadRequest());

            var badRequest = Assert.IsType<BadRequestObjectResult>(result);
            Assert.Equal(
                "Session đang chuyển project, không thể upload điểm lúc này.",
                GetMessage(badRequest.Value));
        }

        [Fact]
        public async Task RestartCurrentProject_WhenSessionIsAdvancing_ReturnsBadRequestWithAdvancingMessage()
        {
            var examSessionService = new Mock<IExamSessionService>();
            var logger = new Mock<ILogger<AgentExamSessionsController>>();

            examSessionService
                .Setup(x => x.RestartCurrentProjectAsync("session-1"))
                .ThrowsAsync(new ArgumentException("Session Ä‘ang chuyá»ƒn project, khÃ´ng thá»ƒ restart lÃºc nÃ y."));

            var controller = new AgentExamSessionsController(
                examSessionService.Object,
                logger.Object);

            var result = await controller.RestartCurrentProject("session-1");

            var badRequest = Assert.IsType<BadRequestObjectResult>(result);
            Assert.Equal(
                "Session Ä‘ang chuyá»ƒn project, khÃ´ng thá»ƒ restart lÃºc nÃ y.",
                GetMessage(badRequest.Value));
        }

        private static string? GetMessage(object? value)
        {
            return value?
                .GetType()
                .GetProperty("message")?
                .GetValue(value) as string;
        }
    }
}
