# Mail Extension

I write this module because of needs.
There are several actions can use.

* declare MailGenerater(SMTP, Sender)
* send mail
* set receiver
* set cc
* set picture (as icon)
* set attechment

Just add the action you need.

```csharp
string context = @"
<img src='cid:logo'></img><br>
Test Send Mail<br>";
using (MailGenerater _mail = new MailComponent("SMTP Server", "Sender"))
{
    _mail.setReceiver("Mail Receiver");
    _mail.setCC("CC to some one");
    _mail.setPicture(context, "logo", "picture path", "memetype, ex: img/png");
    _mail.setAttachment("file path");
    _mail.sendMail("Test Mail", context);
}
```

Or you can send text mail rather than html content.
But you can't append picture if you use text content.

```csharp
string context = @"I'm king of the word.
My name is Jack."
using (MailGenerater _mail = new MailComponent("SMTP Server", "Sender"))
{
    _mail.setReceiver("Mail Receiver");
    _mail.sendMail("Test Mail", context, isHTML:false);
}
```
