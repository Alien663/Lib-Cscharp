# Alien.Common.Mail

Utility library to simplify sending emails using SMTP in .NET.

## 📦 Installation

```bash
Install-Package Alien.Common.Mail
```

## 🚀 Features
- Send plain text or HTML emails
- Attach files
- Use SMTP with customizable config

## 🧪 Example Usage

```csharp
using Alien.Common.Mail;

var config = new MailConfigDto
{
    SmtpHost = "smtp.example.com",
    SmtpPort = 587,
    UseSsl = true,
    UserName = "your@email.com",
    Password = "yourpassword"
};

var mail = new MailDto
{
    To = ["target@email.com"],
    Subject = "Hello from Alien",
    Body = "<h1>Welcome</h1>",
    IsBodyHtml = true
};

MailSender.Send(mail, config);
```

## 📘 Dependencies
- System.Net.Mail

## 🛠 Compatibility
- .NET 6 and above

## 👨‍💻 Author
Alien663
