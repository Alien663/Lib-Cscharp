namespace Alien.Common.Mail.Models
{
    public class MailPictureModel
    {
        public required string ID { get; set; }
        public required string FilePath { get; set; }
        public string Mime { get; set; } = "img/png";
    }
}
