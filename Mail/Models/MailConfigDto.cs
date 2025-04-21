namespace Alien.Common.Mail.Models;

public class MailConfigDto
{
    public string SMTPServer { get; set; }
    public int Port { get; set; } = 25;
}
