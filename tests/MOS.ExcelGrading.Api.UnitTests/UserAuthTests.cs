using MongoDB.Bson;
using MOS.ExcelGrading.Core.Models;
using Xunit;

namespace MOS.ExcelGrading.Api.UnitTests
{
    public class UserAuthTests
    {
        [Fact]
        public void CheckHasPassword_WhenLocalUserWithPassword_ReturnsTrue()
        {
            var user = new User
            {
                Email = "teacher@example.com",
                Username = "teacher01",
                PasswordHash = "some-valid-hash",
                AuthProvider = "Local",
                HasPassword = true
            };

            Assert.True(user.CheckHasPassword());
            Assert.False(user.CheckHasGoogleLinked());
        }

        [Fact]
        public void CheckHasPassword_WhenGoogleOnlyUser_ReturnsFalse()
        {
            var user = new User
            {
                Email = "teacher@example.com",
                Username = "teacher_google",
                PasswordHash = "random-guid-hash",
                AuthProvider = "Google",
                GoogleId = "google-sub-12345",
                HasPassword = false
            };

            Assert.False(user.CheckHasPassword());
            Assert.True(user.CheckHasGoogleLinked());
        }

        [Fact]
        public void CheckHasPassword_WhenLinkedBothUser_ReturnsTrue()
        {
            var user = new User
            {
                Email = "teacher@example.com",
                Username = "teacher01",
                PasswordHash = "user-set-hash",
                AuthProvider = "Both",
                GoogleId = "google-sub-12345",
                HasPassword = true
            };

            Assert.True(user.CheckHasPassword());
            Assert.True(user.CheckHasGoogleLinked());
        }

        [Fact]
        public void UserSerialization_IncludesHasPasswordByDefault()
        {
            var user = new User
            {
                Email = "test@example.com",
                Username = "testuser",
                PasswordHash = "hash",
                HasPassword = true
            };

            var doc = user.ToBsonDocument();
            Assert.True(doc.Contains(nameof(User.HasPassword)));
            Assert.True(doc[nameof(User.HasPassword)].AsBoolean);
        }
    }
}
