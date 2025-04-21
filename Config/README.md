# ConfigExtension

`ConfigExtension` 提供了應用程式設定與加密功能。

## 功能

1. **AppConfig**: 用於讀取應用程式的設定檔（如 `appsettings.json`）。
2. **AESCrypto**: 提供 AES 加密與解密功能。

## 使用方式

### AppConfig

```csharp
var config = new Config.Extention.AppConfig("appsettings.json");
var settingValue = config.Configuration["SettingKey"];
```

### AESCrypto

```csharp
using (var aes = new Config.Extension.AESCrypto())
{
    string encrypted = aes.Encrypt("Hello World");
    string decrypted = aes.Decrypt(encrypted);
}
```