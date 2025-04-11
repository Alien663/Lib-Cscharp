# Config Extension

This library is made to read json config file.
Thus, I add AES crypto function here.

## Appconfig

```csharp
string appConfig = new AppConfig(); // default is appsettings.json
if(appConfig.Confuration["runMode"] == "T")
{
	appConfig = new AppConfig("appsettings.development.json");
}
string connectionString = appConfig.Confuration["ConnectionStrings:DefaultConnection"];
```


## AES Crypto

```csharp
using AESCrypto crypto = new AESCrypto();
string encrypted = crypto.Encrypt("This is a test");
string decrypted = crypto.Decrypt("JeYrjV9buz+kZ902e0pD");
```

```csharp
using AESCrypto crypto = new AESCrypto(key: "123456", iv: "abcdefg");
string encrypted = crypto.Encrypt("This is a test");
string decrypted = crypto.Decrypt("JeYrjV9buz+kZ902e0pD");
```