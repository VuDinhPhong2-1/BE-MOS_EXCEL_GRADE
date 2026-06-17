using MongoDB.Bson;
using MOS.ExcelGrading.Core.Models;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class UserModelSerializationTests
    {
        [Fact]
        public void ToBsonDocument_WhenGoogleIdIsNull_OmitsGoogleIdField()
        {
            var user = new User
            {
                Email = "teacher@example.com",
                Username = "teacher01",
                PasswordHash = "hashed-password",
                AuthProvider = "Local",
                GoogleId = null
            };

            var document = user.ToBsonDocument();

            Assert.False(document.Contains(nameof(User.GoogleId)));
        }
    }
}
