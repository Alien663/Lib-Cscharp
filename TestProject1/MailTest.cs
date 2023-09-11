using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SimpleMail;
using NUnit.Framework;

namespace TestMyLib
{
    public class MailTest
    {
        [SetUp]
        private void Setup()
        {
        }

        [Test]
        public void Test1()
        {
            string context = @"
<img src='cid:logo'></img><br>
Test Send Mail<br>";
            MailComponent _mail = new MailComponent("SMTP Server", "Sender");
            _mail.setReceiver("Mail Receiver");
            _mail.setCC("CC to some one");
            _mail.setPicture(context, "logo", "picture path", "memetype, ex: img/png");
            _mail.setAttachment("file path");
            _mail.sendMail("Test Mail", context);
        }

        [Test]
        public void Test2()
        {
            string context = @"
<img src='cid:logo'></img><br>
Test Send Mail<br>";

            MailPictureModel _picture = new MailPictureModel
            {
                ID = "logo",
                FilePath = "picture path",
                Mime = "img/png",
            };
            MailComponent _mail = new MailComponent("SMTP Server", "Sender");
            _mail.setReceiver("Mail Receiver");
            _mail.setPicture(context, _picture);
            _mail.sendMail("Test Mail", context);
        }
    }
}
