using Alien.Common.Mail.Models;

namespace Alien.Common.Mail;

public interface IMailClient
{
    public void Send(MailDto mail);
}
