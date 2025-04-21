# Alien.Common.Data

Lightweight wrapper for Dapper-based database access.

## 📦 Installation

```bash
Install-Package Alien.Common.Data
```

## 🚀 Features
- Execute queries and stored procedures
- Map results to POCO objects
- Simplified transaction support

## 🧪 Example Usage

```csharp
using Alien.Common.Data;

var db = new DbService("YourConnectionString");

var user = db.QuerySingleOrDefault<User>("SELECT * FROM Users WHERE Id = @Id", new { Id = 1 });
```

## 📘 Dependencies
- Dapper

## 🛠 Compatibility
- .NET 6 and above

## 👨‍💻 Author
Alien663