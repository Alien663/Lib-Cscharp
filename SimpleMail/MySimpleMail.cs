using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net.Mime;

namespace SimpleMail
{
    public class MailComponent
    {
        private string SMTPServer;
        private List<Attachment> Attachments = new List<Attachment>();
        private MailMessage mail = new MailMessage();

        public MailComponent(string SMTPServer, string Sender)
        {
            this.SMTPServer = SMTPServer;
            this.mail.From = new MailAddress(Sender);
        }

        public void SendMail(string title, string context, bool isHTML = true)
        {
#if DEBUG
            title = "[Test]" + title;
            context = "It's a test mail" + context;
#endif
            this.mail.Subject = title;
            this.mail.Body = context;
            this.mail.IsBodyHtml = isHTML;
            SmtpClient smtp = new SmtpClient(SMTPServer);
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Send(this.mail);
        }
        public void SetReceiver(string receiver)
        {
            this.mail.To.Add(receiver);
        }
        public void SetReceiver(string[] receivers)
        {
            foreach (string receiver in receivers)
            {
                this.SetReceiver(receiver);
            }
        }
        public void SetCC(string cc)
        {
            this.mail.CC.Add(cc);
        }
        public void SetCC(string[] ccs)
        {
            foreach (string cc in ccs)
            {
                this.SetCC(cc);
            }
        }
        public void SetPicture(string context, string ID, string FilePath, string Mime)
        {
            AlternateView htmlview = AlternateView.CreateAlternateViewFromString(context, null, MediaTypeNames.Text.Html);
            LinkedResource imageLink = new LinkedResource(FilePath, Mime);
            imageLink.ContentId = ID;
            imageLink.TransferEncoding = TransferEncoding.Base64;
            htmlview.LinkedResources.Add(imageLink);
            this.mail.AlternateViews.Add(htmlview);
        }
        public void SetPicture(string context, MailPictureModel picture)
        {
            AlternateView htmlview = AlternateView.CreateAlternateViewFromString(context, null, MediaTypeNames.Text.Html);
            LinkedResource imageLink = new LinkedResource(picture.FilePath, picture.Mime);
            imageLink.ContentId = picture.ID;
            imageLink.TransferEncoding = TransferEncoding.Base64;
            htmlview.LinkedResources.Add(imageLink);
            this.mail.AlternateViews.Add(htmlview);
        }
        public void SetPicture(string context, List<MailPictureModel> pictures)
        {
            foreach (MailPictureModel picture in pictures)
            {
                SetPicture(context, picture);
            }
        }
        public void SetAttachment(string filepath)
        {
            this.mail.Attachments.Add(new Attachment(filepath));
        }
        public void SetAttachment(string[] filepaths)
        {
            foreach (string filepath in filepaths)
            {
                this.SetAttachment(filepath);
            }
        }
        public void SetAttachment(Attachment attachment)
        {
            this.mail.Attachments.Add(attachment);
        }
        public void SetAttachment(List<Attachment> attachments)
        {
            foreach (Attachment attachment in attachments)
            {
                this.SetAttachment(attachment);
            }
        }
    }

    public class MailPictureModel
    {
        public string ID { get; set; }
        public string FilePath { get; set; }
        public string Mime { get; set; }
    }
}
