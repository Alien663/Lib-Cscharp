using NUnit.Framework;
using Config.Extention;
using Config.Extension;
using System;

namespace TestMyLib
{
    public class ConfigExtension
    {
        [Test]
        [Order(1)]
        public void ReadAppconfig()
        {
            #region Arrange
            AppConfig _config = new AppConfig("appsettings.dev.json");
            #endregion

            #region Action
            #endregion

            #region Assert
            Assert.That(_config.Configuration["ConnectionStrings:DefaultConnection"] == "This is dev");
            #endregion
        }
        
        [Test]
        [Order(2)]
        public void AESEncryp()
        {
            #region Arrange
            using AESCrypto crypto = new AESCrypto();
            #endregion

            #region Action
            string encrypted = crypto.Encrypt("This is a test");
            Console.WriteLine(encrypted);
            #endregion

            #region Assert
            Assert.That(encrypted == "JeYrjV9buz+kZ902e0pD");
            #endregion
        }

        [Test]
        [Order(3)]
        public void AESDecrypt()
        {
            #region Arrange
            using AESCrypto crypto = new AESCrypto();
            #endregion

            #region Action
            string decrypted = crypto.Decrypt("JeYrjV9buz+kZ902e0pD");
            #endregion

            #region Assert
            Assert.That(decrypted == "This is a test");
            #endregion
        }
    }
}
