using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Net.Mail;
using System.Net.Mime;

namespace Alien.Common.Mail.Models;

public class MailDto
{
    public string Sender { get; set; }
    public List<string> To { get; set; }
    public List<string> CC { get; set; }
    public List<Attachment> attachments { get; set; } = new List<Attachment>();
    public string Subject { get; set; }
    public string Body { get; set; }
    public bool IsBodyHtml { get; set; } = true;

    public MailMessage message
    {
        get
        {
            message.From = new MailAddress(Sender);
            foreach (var item in To) message.To.Add(item);
            foreach (var item in CC) message.CC.Add(item);
            foreach (var item in attachments) message.Attachments.Add(item);
            message.Subject = this.Subject;
            message.Body = this.Body;
            IsBodyHtml = this.IsBodyHtml;
            return message;
        }
        private set { message = value; }
    }

    public void setPicture(string ID, string FilePath, string Mime)
    {
        AlternateView htmlview = AlternateView.CreateAlternateViewFromString(this.Body, null, MediaTypeNames.Text.Html);
        LinkedResource imageLink = new LinkedResource(FilePath, Mime);
        imageLink.ContentId = ID;
        imageLink.TransferEncoding = TransferEncoding.Base64;
        htmlview.LinkedResources.Add(imageLink);
        this.message.AlternateViews.Add(htmlview);
    }
    public void setPicture(MailPictureModel picture)
    {
        AlternateView htmlview = AlternateView.CreateAlternateViewFromString(this.Body, null, MediaTypeNames.Text.Html);
        LinkedResource imageLink = new LinkedResource(picture.FilePath, picture.Mime);
        imageLink.ContentId = picture.ID;
        imageLink.TransferEncoding = TransferEncoding.Base64;
        htmlview.LinkedResources.Add(imageLink);
        this.message.AlternateViews.Add(htmlview);
    }
    public void setPicture(List<MailPictureModel> pictures)
    {
        foreach (MailPictureModel picture in pictures)
        {
            setPicture(picture);
        }
    }
}
