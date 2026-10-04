using System;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class WordTrackChangesTests
    {
        [Fact]
        public void ModernPasswordHashUsesWordIteratorBeforeDigest()
        {
            const string password = "Legal";
            var salt = Encoding.UTF8.GetBytes("test-salt");
            const int spinCount = 1;
            var protection = CreateProtection(password, salt, spinCount);

            Assert.True(Validate(protection, password));
            Assert.False(Validate(protection, "legal"));
            Assert.False(Validate(protection, "LEGAL"));
            Assert.False(Validate(protection, "WrongPassword"));
            Assert.False(Validate(CreateProtection("legal", salt, spinCount), password));
            Assert.False(Validate(CreateProtection("LEGAL", salt, spinCount), password));
        }
        [Fact]
        public void StandardWordDocumentProtectionPasswordMatchesMicrosoftWordEcma376()
        {
            const string realWordXml = @"<w:documentProtection xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
                w:edit=""trackedChanges""
                w:enforcement=""1""
                w:cryptProviderType=""rsaAES""
                w:cryptAlgorithmClass=""hash""
                w:cryptAlgorithmType=""typeAny""
                w:cryptAlgorithmSid=""14""
                w:cryptSpinCount=""100000""
                w:hash=""KKneFdnfPlYKL59jSjJr9C/NWqfxwcxn3h3dgaWuH+7od/MsRTnlvT2ssiD0lHfGuVF0cokdhpuY8brKBcXKcA==""
                w:salt=""/eEQ8CKxpJTK2J83dIUYdg=="" />";

            var protection = XElement.Parse(realWordXml);
            Assert.True(Validate(protection, "Legal"));
            Assert.False(Validate(protection, "legal"));
            Assert.False(Validate(protection, "LEGAL"));
            Assert.False(Validate(protection, "WrongPassword"));
        }


        [Theory]
        [InlineData("word", true)]
        [InlineData("Word", false)]
        [InlineData(" word ", false)]
        public void SpecialConditionSubjectComparisonIsCaseAndWhitespaceSensitive(string subject, bool expected)
        {
            var method = typeof(XmlGradingRuleService).GetMethod(
                "IsSpecialConditionSupportedForSubject",
                BindingFlags.Static | BindingFlags.NonPublic)!;

            var result = (bool)method.Invoke(
                null,
                new object[] { "wordTrackChanges", subject })!;

            Assert.Equal(expected, result);
        }

        private static byte[] HashWordPassword(string password, byte[] salt, int spinCount)
        {
            var passwordBytes = Encoding.Unicode.GetBytes(password);
            var input = new byte[salt.Length + passwordBytes.Length];
            Buffer.BlockCopy(salt, 0, input, 0, salt.Length);
            Buffer.BlockCopy(passwordBytes, 0, input, salt.Length, passwordBytes.Length);
            var hash = SHA512.HashData(input);
            for (var i = 0; i < spinCount; i++)
            {
                var iteration = BitConverter.GetBytes(i);
                var iterationInput = new byte[iteration.Length + hash.Length];
                Buffer.BlockCopy(iteration, 0, iterationInput, 0, iteration.Length);
                Buffer.BlockCopy(hash, 0, iterationInput, iteration.Length, hash.Length);
                hash = SHA512.HashData(iterationInput);
            }
            return hash;
        }

        private static XElement CreateProtection(string password, byte[] salt, int spinCount)
        {
            var hash = HashWordPassword(password, salt, spinCount);
            return XElement.Parse($"<w:documentProtection xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' w:hashValue='{Convert.ToBase64String(hash)}' w:saltValue='{Convert.ToBase64String(salt)}' w:spinCount='{spinCount}'/>");
        }

        private static bool Validate(XElement protection, string password)
        {
            var method = typeof(XmlGradingRuleService).GetMethod("IsWordProtectionPasswordValid", BindingFlags.Static | BindingFlags.NonPublic)!;
            return (bool)method.Invoke(null, new object[] { protection, password })!;
        }
    }
}