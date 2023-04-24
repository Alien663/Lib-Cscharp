using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyLib;
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
Test Send Mail<br>
";
            MailComponent _mail = new MailComponent("SMTP Server", "Sender");
            _mail.SetReceiver("Mail Receiver");
            _mail.SetCC("CC to some one");
            _mail.SetPicture(context, "logo", "picture path", "memetype, ex: img/png");
            _mail.SetAttachment("file path");
            _mail.SendMail("Test Mail", context);
        }
        
        public void Test2()
        {
            string context = @"
<img src='cid:logo'></img><br>
Test Send Mail<br>
";

            MailPictureModel _picture = new MailPictureModel
            {
                ID = "logo",
                FilePath = "picture path",
                Mime = "img/png",
            };

            MailComponent _mail = new MailComponent("SMTP Server", "Sender");
            _mail.SetReceiver("Mail Receiver");
            _mail.SetPicture(context, _picture);
            _mail.SendMail("Test Mail", context);
        }
    }
}
