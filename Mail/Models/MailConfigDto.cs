namespace Alien.Common.Mail.Models;

public class MailConfigDto
{
    public required string SMTPServer { get; set; }
    public int Port { get; set; } = 25;
}
