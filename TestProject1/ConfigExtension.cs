using NUnit.Framework;
using Config.Extention;
using Config.Extension;

namespace TestMyLib
{
    public class ConfigExtension
    {
        [Test]
        public void Test01_ReadAppconfig()
        {
            #region Arrange
            #endregion

            #region Action
            AppConfig _config = new AppConfig("appsettings.dev.json");
            string tt = AppConfig.Config["ConnectionStrings:DefaultConnection"];
            #endregion

            #region Assert
            Assert.That(tt == "This is dev");
            #endregion
        }
        [Test]
        public void Test02_AESEncryp()
        {
            #region Arrange
            using AESCrypto crypto = new AESCrypto();
            #endregion

            #region Action
            string encrypted = crypto.Encrypt("This is a test");
            #endregion

            #region Assert
            Assert.That(encrypted == "");
            #endregion
        }

        [Test]
        public void Test02_AESDecrypt()
        {
            #region Arrange
            #endregion

            #region Action
            #endregion

            #region Assert
            #endregion
        }
    }
}
