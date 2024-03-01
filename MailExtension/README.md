# Mail Module

I write this module because of needs.
There are several actions can use.

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
MailComponent _mail = new MailComponent("SMTP Server", "Sender");
_mail.setReceiver("Mail Receiver");
_mail.setCC("CC to some one");
_mail.setPicture(context, "logo", "picture path", "memetype, ex: img/png");
_mail.setAttachment("file path");
_mail.sendMail("Test Mail", context);
```

Or you can send text mail rather than html content.
But you can't append picture if you use text content.

```csharp
string context = @"I'm king of the word.
My name is Jack."
MailComponent _mail = new MailComponent("SMTP Server", "Sender");
_mail.setReceiver("Mail Receiver");
_mail.sendMail("Test Mail", context, isHTML:false);
```
