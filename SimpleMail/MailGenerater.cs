using System.Net.Mail;
using System.Net.Mime;

namespace Mail.Extension
{
    public class MailGenerater : IDisposable
    {
        private bool _disposed;
        private string SMTPServer;
        private List<Attachment> Attachments = new List<Attachment>();
        private MailMessage mail = new MailMessage();

        public MailGenerater(string SMTPServer, string Sender)
        {
            this.SMTPServer = SMTPServer;
            this.mail.From = new MailAddress(Sender);
        }

        public void sendMail(string title, string context, bool isHTML = true)
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
        public void setReceiver(string receiver)
        {
            this.mail.To.Add(receiver);
        }
        public void setReceiver(string[] receivers)
        {
            foreach (string receiver in receivers)
            {
                this.setReceiver(receiver);
            }
        }
        public void setCC(string cc)
        {
            this.mail.CC.Add(cc);
        }
        public void setCC(string[] ccs)
        {
            foreach (string cc in ccs)
            {
                this.setCC(cc);
            }
        }
        public void setPicture(string context, string ID, string FilePath, string Mime)
        {
            AlternateView htmlview = AlternateView.CreateAlternateViewFromString(context, null, MediaTypeNames.Text.Html);
            LinkedResource imageLink = new LinkedResource(FilePath, Mime);
            imageLink.ContentId = ID;
            imageLink.TransferEncoding = TransferEncoding.Base64;
            htmlview.LinkedResources.Add(imageLink);
            this.mail.AlternateViews.Add(htmlview);
        }
        public void setPicture(string context, MailPictureModel picture)
        {
            AlternateView htmlview = AlternateView.CreateAlternateViewFromString(context, null, MediaTypeNames.Text.Html);
            LinkedResource imageLink = new LinkedResource(picture.FilePath, picture.Mime);
            imageLink.ContentId = picture.ID;
            imageLink.TransferEncoding = TransferEncoding.Base64;
            htmlview.LinkedResources.Add(imageLink);
            this.mail.AlternateViews.Add(htmlview);
        }
        public void setPicture(string context, List<MailPictureModel> pictures)
        {
            foreach (MailPictureModel picture in pictures)
            {
                setPicture(context, picture);
            }
        }
        public void setAttachment(string filepath)
        {
            this.mail.Attachments.Add(new Attachment(filepath));
        }
        public void setAttachment(string[] filepaths)
        {
            foreach (string filepath in filepaths)
            {
                this.setAttachment(filepath);
            }
        }
        public void setAttachment(Attachment attachment)
        {
            this.mail.Attachments.Add(attachment);
        }
        public void setAttachment(List<Attachment> attachments)
        {
            foreach (Attachment attachment in attachments)
            {
                this.setAttachment(attachment);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                if (disposing)
                {
                    mail.Dispose();
                }
                _disposed = true;
            }
        }
    }
}
