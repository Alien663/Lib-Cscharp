using Alien.Common.Mail.Models;

namespace Alien.Common.Mail;

public interface ISmtpClientWrapper
{
    void Send(MailDto mail);
    void Send(IEnumerable<MailDto> mails);
}
