using MimeKit;

namespace Alien.Common.Mail.Models;

public class MailDto
{
    public required string Sender { get; set; }
    public required List<string> To { get; set; }
    public List<string> CC { get; set; } = new List<string>();
    public List<string> BCC { get; set; } = new List<string>();
    public List<string> Attachments { get; set; } = new List<string>();
    public required string Subject { get; set; }
    public string Body { get; set; } = "";

    public MimeMessage message
    {
        get
        {
            message.Subject = this.Subject;
            message.Sender = new MailboxAddress(Sender, Sender);
            foreach (var item in To) message.To.Add(new MailboxAddress(item, item));
            foreach (var item in CC) message.Cc.Add(new MailboxAddress(item, item));
            foreach (var item in BCC) message.Bcc.Add(new MailboxAddress(item, item));
            var bodyBuilder = new BodyBuilder();
            bodyBuilder.HtmlBody = this.Body;
            foreach (var item in Attachments)
            {
                if (!File.Exists(item))
                {
                    throw new ArgumentException("File not found", nameof(item));
                }
                bodyBuilder.Attachments.Add(item);
            }
            message.Body = bodyBuilder.ToMessageBody();
            return message;
        }
        private set { message = value; }
    }

    public void setPicture(string ID, string FilePath, string Mime)
    {
        if (!File.Exists(FilePath))
        {
            throw new ArgumentException("File not found", nameof(FilePath));
        }
        var bodyBuilder = new BodyBuilder();
        bodyBuilder.HtmlBody = message.HtmlBody;

        if(!bodyBuilder.HtmlBody.Contains($"cid:{ID}"))
        {
            throw new ArgumentException("ID not found in body", nameof(ID));
        }

        bodyBuilder.LinkedResources.Add(new MimePart(Mime)
        {
            ContentId = ID,
            ContentTransferEncoding = ContentEncoding.Base64,
            FileName = Path.GetFileName(FilePath),
            ContentDisposition = new ContentDisposition(ContentDisposition.Inline)
            {
                IsAttachment = false
            }
        });
        message.Body = bodyBuilder.ToMessageBody();
    }
    public void setPicture(MailPictureModel picture)
    {
        if (!File.Exists(picture.FilePath))
        {
            throw new ArgumentException("File not found", nameof(picture.FilePath));
        }
        var bodyBuilder = new BodyBuilder();
        bodyBuilder.HtmlBody = message.HtmlBody;
        if (!bodyBuilder.HtmlBody.Contains($"cid:{picture.ID}"))
        {
            throw new ArgumentException("ID not found in body", nameof(picture.ID));
        }
        bodyBuilder.LinkedResources.Add(new MimePart(picture.Mime)
        {
            ContentId = picture.ID,
            ContentTransferEncoding = ContentEncoding.Base64,
            FileName = Path.GetFileName(picture.FilePath),
            ContentDisposition = new ContentDisposition(ContentDisposition.Inline)
            {
                IsAttachment = false
            }
        });
        message.Body = bodyBuilder.ToMessageBody();
    }
    public void setPicture(List<MailPictureModel> pictures)
    {
        foreach (MailPictureModel picture in pictures)
        {
            setPicture(picture);
        }
    }
}
