using System.Net.Mail;
using Alien.Common.Mail.Models;

namespace Alien.Common.Mail;

internal class MailClient : IMailClient, IDisposable
{
    private bool _disposed;
    private readonly SmtpClient smtpClient;

    public MailClient(MailConfigDto config)
    {
        smtpClient = new SmtpClient(config.SMTPServer, config.Port);
    }

    public void Send(MailDto mail)
    {
        CheckDisposed();
        smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
        smtpClient.Send(mail.message);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                smtpClient.Dispose();
            }
            _disposed = true;
        }
    }

    protected void CheckDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(MailClient));
        }
    }
}
