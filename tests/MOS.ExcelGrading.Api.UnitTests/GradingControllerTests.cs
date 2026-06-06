using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Moq;
using MOS.ExcelGrading.API.Controllers;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Security.Claims;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class GradingControllerTests
    {
        [Fact]
        public async Task GradeWordProject_Project07_AllowsTxtFile_ReturnsOk()
        {
            // Arrange
            var gradingService = new Mock<IGradingService>();
            var analytics = new Mock<IAnalyticsService>();
            var logger = new Mock<ILogger<GradingController>>();

            gradingService
                .Setup(s => s.GradeWordProjectAsync(7, It.IsAny<Stream>(), It.IsAny<string?>()))
                .ReturnsAsync(new GradingResult { ProjectId = "word07", TotalScore = 1, MaxScore = 1 });

            analytics
                .Setup(a => a.SaveGradingAttemptAsync(It.IsAny<GradingResult>(), It.IsAny<string>(), It.IsAny<string?>(), It.IsAny<string?>(), It.IsAny<string?>(), It.IsAny<bool>() ))
                .Returns(Task.CompletedTask);

            var controller = new GradingController(gradingService.Object, analytics.Object, logger.Object);

            var claims = new[] {
                new Claim("permission", Permissions.CreateGrades),
                new Claim(ClaimTypes.NameIdentifier, "tester")
            };
            controller.ControllerContext = new ControllerContext
            {
                HttpContext = new DefaultHttpContext { User = new ClaimsPrincipal(new ClaimsIdentity(claims, "test")) }
            };

            var content = Encoding.UTF8.GetBytes("Hello plain text");
            using var ms = new MemoryStream(content);
            ms.Position = 0;
            IFormFile file = new FormFile(ms, 0, ms.Length, "studentFile", "submission.txt")
            {
                Headers = new HeaderDictionary(),
                ContentType = "text/plain"
            };

            // Act
            var actionResult = await controller.GradeWordProject("project07", file, null, null, null);

            // Assert
            Assert.IsType<OkObjectResult>(actionResult);
        }

        [Fact]
        public async Task GradeWordProject_Project08_RejectsNonDocx_ReturnsBadRequest()
        {
            // Arrange
            var gradingService = new Mock<IGradingService>();
            var analytics = new Mock<IAnalyticsService>();
            var logger = new Mock<ILogger<GradingController>>();

            var controller = new GradingController(gradingService.Object, analytics.Object, logger.Object);

            var claims = new[] {
                new Claim("permission", Permissions.CreateGrades),
                new Claim(ClaimTypes.NameIdentifier, "tester")
            };
            controller.ControllerContext = new ControllerContext
            {
                HttpContext = new DefaultHttpContext { User = new ClaimsPrincipal(new ClaimsIdentity(claims, "test")) }
            };

            var content = Encoding.UTF8.GetBytes("Not a docx");
            using var ms = new MemoryStream(content);
            ms.Position = 0;
            IFormFile file = new FormFile(ms, 0, ms.Length, "studentFile", "submission.txt")
            {
                Headers = new HeaderDictionary(),
                ContentType = "text/plain"
            };

            // Act
            var actionResult = await controller.GradeWordProject("project08", file, null, null, null);

            // Assert
            Assert.IsType<BadRequestObjectResult>(actionResult);
        }
    }
}
