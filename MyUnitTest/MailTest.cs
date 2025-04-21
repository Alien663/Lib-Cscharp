using NUnit.Framework;
using System.Net.Mail;
using System.Collections.Generic;
using Alien.Common.Mail;
using Alien.Common.Mail.Models;

namespace TestMyLib;

[TestFixture]
public class MailTest
{
    [Test]
    public void MailSender_SouldCallSendMethod()
    {
        #region Arrange
        var fakeClient = new FakeMailCleint();
        var config = new MailConfigDto { SMTPServer = "smtp.demo.com" };
        var mail = new MailDto
        {
            Sender = "test@demo.com",
            To = new List<string> { "test@demo.com" },
            CC = new List<string> { "test@demo.com" },
            Subject = "Test",
            Body = "Hi Leon,<br>\r\nHelp me to do the unit test of mail module<br>\r\n<img src='cid:logo1'></img><br>\r\n<img src='cid:logo2'></img><br>\r\nTest Send Mail<br>",
            IsBodyHtml = true,
            attachments = new List<Attachment>
            {
                new Attachment(@".\icon.png")
            }
        };
        mail.setPicture(ID: "logo1", FilePath: @".\icon.png", Mime: "img/png");
        mail.setPicture(new List<MailPictureModel>
        {
            new MailPictureModel
            {
                ID = "logo2",
                FilePath = @".\icon.png",
                Mime = "img/png"
            }
        });
        #endregion

        #region Act
        fakeClient.Send(mail);
        #endregion

        #region Assert
        Assert.That(fakeClient.Sent, Is.True);
        #endregion
    }

    public class FakeMailCleint : IMailClient
    {
        public bool Sent = false;
        public void Send(MailDto dto) => Sent = true;
    }
}
