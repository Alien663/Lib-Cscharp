using MailKit.Net.Smtp;
using Alien.Common.Mail.Models;

namespace Alien.Common.Mail;

internal class MailClient : IMailClient, IDisposable
{
    private bool _disposed;
    private readonly SmtpClient smtpClient;

    public MailClient(MailConfigDto config)
    {
        smtpClient = new SmtpClient();
        smtpClient.Connect(config.SMTPServer, config.Port, false);
    }

    public void Send(MailDto mail)
    {
        CheckDisposed();
        smtpClient.Send(mail.message);
        smtpClient.Disconnect(true);
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
