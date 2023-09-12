using NUnit.Framework;
using SimpleMail;
using System.Collections.Generic;
using System.Net.Mail;

namespace TestMyLib
{
    public class MailTest
    {
        [Test]
        public void Test1()
        {
            string context = @"
Hi Leon,<br>
Help me to do the unit test of mail module<br>
<img src='cid:logo'></img><br>
Test Send Mail<br>";
            MailComponent _mail = new MailComponent("10.243.29.110", "tds@quantatw.com");
            _mail.setReceiver ("Bryan_Lin@quantatw.com");
            _mail.setCC("leon-chen@quantatw.com");
            _mail.setPicture(context, "logo", @".\icon.png", "img/png");
            _mail.setAttachment(@".\icon.png");
            _mail.setAttachment(new string[] { @".\icon.png" });
            _mail.sendMail("[Test] Unit Test Mail", context);
        }

        [Test]
        public void Test2()
        {
            string context = @"
Hi Leon,<br>
Help me to do the unit test of mail module<br>
<img src='cid:logo1'></img><br>
<img src='cid:logo2'></img><br>
Test Send Mail<br>";
            MailComponent _mail = new MailComponent("10.243.29.110", "tds@quantatw.com");
            _mail.setReceiver(new string[] { "Bryan_Lin@quantatw.com" });
            _mail.setCC(new string[] { "leon-chen@quantatw.com" });
            _mail.setPicture(context, new MailPictureModel
            {
                ID="logo1",
                FilePath=@".\icon.png",
                Mime="img/png"
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
            _mail.setAttachment(new List<Attachment>{ new Attachment(@".\icon.png") });
            _mail.sendMail("[Test] Unit Test Mail", context);
        }
    }
}
