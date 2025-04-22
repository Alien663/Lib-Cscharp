# Lib-CSharp 📦

🚀 A personal utility library for .NET developers, built with modular design across Email, Excel, Configuration, and Data processing.

> 給你重複造輪子的人一點救贖：這裡有封裝過的常用功能可以偷懶。

---

## 🧩 模組簡介

- **[Mail](./Mail/README.md)** - 使用 MailKit 快速寄送文字/HTML/附檔信件。
- **[Excel](./Excel/README.md)** - NPOI 操作封裝，輕鬆讀寫 Excel。
- **[Config](./Config/README.md)** - 封裝 JSON / INI 讀取，快速整合設定檔。
- **[Utility](./UnitTest/README.md)** - 其他懶人專用工具方法集。

---

## 💡 安裝方式

請自行 clone 專案，或將需要的模組引入你的專案中。

未來會提供 NuGet 套件（如果我哪天良心發現有包起來的話）。

---

## 🔧 使用範例

更多參數使用請參照Unit Test(其實就是我懶得寫那麼詳細的README)

```csharp
// Mail example
var mail = new MailDto
{
    Sender = "test@demo.com",
    To = new List<string> { "test@demo.com" },
    Subject = "Test",
    Body = "Hello World~",
};
var config = new MailConfigDto { SMTPServer = "smtp.demo.com" };
ISmtpClientWrapper client = SmtpClientWrapper(config);
client.Send(mail);
```

```csharp
// Config example
AppConfig _config = new AppConfig(); // default is appsettings.json
if(_config.Configuration["isDebug"])
{
    _config = new AppConfig("appsettings.development.json");
}
string connectionString = _config.Configuration["ConnectionStrings:DefaultConnection"]
```

```csharp
// Excel example
using Alien.Common.Excel;

var data = new List<MyModel>
{
    new MyModel { Name = "John", Age = 30 },
    new MyModel { Name = "Jane", Age = 25 }
};

using ExcelConverter excel = new ExcelConverter();
using FileStream fs = File.Create(filename);
byte[] data = excel.export(rawData);
fs.Write(data, 0, data.Length);
```

## 🧪單元測試

測試位於 MyUnitTest/ 資料夾。

建議搭配 xUnit/NUnit 跑測試並整合 GitHub Actions 自動測。

## 🤝貢獻方式

目前尚未開放貢獻（這是我私人輪子樂園）。

不過你可以開 issue 吐槽我寫得多爛，我會視心情修。

## 📝授權

MIT License - 請隨意使用、改造、甚至忘記我曾經存在。
