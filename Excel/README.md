# Alien.Common.Excel

A utility library for importing/exporting Excel files using ClosedXML and OpenXML.

## 📦 Installation

```bash
Install-Package Alien.Common.Excel
```

## 🚀 Features
- Export data to Excel from List<T>
- Import data from Excel into DataTable
- Apply styling and formatting

## 🧪 Example Usage

```csharp
using Alien.Common.Excel;

var data = new List<MyModel>
{
    new MyModel { Name = "John", Age = 30 },
    new MyModel { Name = "Jane", Age = 25 }
};

var stream = ExcelExporter.ExportToExcel(data);
File.WriteAllBytes("output.xlsx", stream.ToArray());
```

## 📘 Dependencies
- ClosedXML
- DocumentFormat.OpenXml

## 🛠 Compatibility
- .NET 6 and above

## 🧙 Author
Alien663
