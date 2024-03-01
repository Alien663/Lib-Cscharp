using NUnit.Framework;
using Mail.Extension;
using System.Collections.Generic;
using System.Net.Mail;

namespace TestMyLib
{
    [TestFixture]
    public class TestMailExtension
    {
        string smtp;
        string sender;
        string receiver;
        string cc;

        [OneTimeSetUp]
        public void Initialize()
        {
            smtp = "GLTDSSMTP.quanta.corp";
            sender = "tds@quantatw.com";
            receiver = "Bryan_Lin@quantatw.com";
            cc = "leon-chen@quantatw.com";
        }

        [Test]
        public void Test1()
        {
            #region Arrange
            string context = @"
                Hi Leon,<br>
                Help me to do the unit test of mail module<br>
                <img src='cid:logo'></img><br>
                Test Send Mail<br>
            ";
            #endregion

            #region Act
            using (MailGenerater _mail = new MailGenerater(smtp, sender))
            {
                _mail.setReceiver(receiver);
                _mail.setCC(cc);
                _mail.setPicture(context, "logo", @".\icon.png", "img/png");
                _mail.setAttachment(@".\icon.png");
                _mail.setAttachment(new string[] { @".\icon.png" });
                _mail.sendMail("[Test] Unit Test Mail", context);
            }
            #endregion

            #region Assert
            /*
             * I don't know how to check mail modules success or not...
             */
            #endregion
        }

        [Test]
        public void Test2()
        {
            #region Arrange
            string context = @"
                Hi Leon,<br>
                Help me to do the unit test of mail module<br>
                <img src='cid:logo1'></img><br>
                <img src='cid:logo2'></img><br>
                Test Send Mail<br>
            ";
            #endregion

            #region Act
            using (MailGenerater _mail = new MailGenerater(smtp, sender))
            {
                _mail.setReceiver(new string[] { receiver });
                _mail.setCC(new string[] { cc });
                _mail.setPicture(context, new MailPictureModel
                {
                    ID = "logo1",
                    FilePath = @".\icon.png",
                    Mime = "img/png"
                });
                _mail.setPicture(context, new List<MailPictureModel>
            {
                new MailPictureModel
                {
                    ID="logo2",
                    FilePath=@".\icon.png",
                    Mime="img/png"
                }
            });
                _mail.setAttachment(new List<Attachment> { new Attachment(@".\icon.png") });
                _mail.sendMail("[Test] Unit Test Mail", context);
            }
            #endregion

            #region Assert
            /*
             * I don't know how to check mail modules success or not...
             */
            #endregion
        }
    }
}
