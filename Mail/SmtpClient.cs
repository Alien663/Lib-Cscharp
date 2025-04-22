using MailKit.Net.Smtp;
using Alien.Common.Mail.Models;

namespace Alien.Common.Mail;

internal class SmtpClientWrapper : ISmtpClientWrapper, IDisposable
{
    private bool _disposed;
    private readonly ISmtpClient smtpClient;

    public SmtpClientWrapper(MailConfigDto config)
    {
        smtpClient = new SmtpClient();
        try
        {
            smtpClient.Connect(config.SMTPServer, config.Port, false);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Fail to connect SMTP server", ex);
        }
    }

    public SmtpClientWrapper(ISmtpClient smtpClient)
    {
        this.smtpClient = smtpClient;
    }

    public void Send(MailDto mail)
    {
        CheckDisposed();
        smtpClient.Send(mail.message);
    }

    public void Send(IEnumerable<MailDto> mails)
    {
        CheckDisposed();
        foreach (var mail in mails)
        {
            smtpClient.Send(mail.message);
        }
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
                if (smtpClient.IsConnected)
                    smtpClient.Disconnect(true);
                smtpClient.Dispose();
            }
            _disposed = true;
        }
    }

    protected void CheckDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(SmtpClientWrapper), "The SMTP client has already been disposed.");
        }
    }
}
