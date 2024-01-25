# ExcelComponent

I make this tool based on NuGet package NPOI.
There are kinds of functions :

1. export excel file from DataTable
2. export excel file from DataTable and strat with specific location
3. export excel file from DataSet(with multiple sheet)
4. export excel file from DataSet(with multiple sheet) and strat with specific location
5. export excel file from DataModel and strat with specific location
6. set datatype mapping to style before export file
7. read DataTable data from excel file
8. read DataSet data from excel file(with multiple sheet)
9. read DataModel data from excel file (Work on it now)

If you wants to use this component to read/write excel file, you can see the sample code in UnitTest project.There are all functions' sample code I prepared to test my library.

The follow content will introduce you my functions.

## Introduction

### Export Excel with custom Display Name

* Now you can define your data model as bellow, when it exports file from data model, it will use "DisplayName" as column name in excel file automatically.

```csharp
public class TestModel
{
  public string ID { get; set; }
  
  [DisplayName("Test Name")]
  public string Name { get; set; }
}
```

### Setting the start with of your exel file

You can defined your table start with a specific location.
For example, start at cell(2, 2)

```csharp
ExcelComponent myexcel = new ExcelComponent();
myexcel.SetRange(new SheetRange{
  MinRowIndex = 2,
  MinColIndex = 2,
});
byte[] data = myexcel.export(this.dtStudent);
using (FileStream fs = File.Create(this.folder + "test.xlsx"))
{
    fs.Write(data, 0, data.Length);
}
```

### Custom Sheet Name and Multiple Sheets

```csharp
DataTable dtData2 = dtData1.Copy()
dtData2.TableName = "Custom Sheet Name" // custom sheet name by table name
DataSet stData = new DataSet(); // use DataSet to process multiple sheets
stData.Add(dtData1);
stData.Add(dtData2);
byte[] data = myexcel.export(st);
using (FileStream fs = File.Create(this.folder + "test2.xlsx"))
{
    fs.Write(data, 0, data.Length);
}
```

### Set DataType's Style

You can change the style to output.
I use this function to change the style of double originally.

```csharp
ExcelComponent myexcel = new ExcelComponent();
myexcel.setDataTypeStyle(new Dictionary<string, string> { { "Double", "#,##0.0000" } });
var data = myexcel.export(this.Students);
using (FileStream fs = File.Create(this.folder + "test4.xlsx"))
{
    fs.Write(data, 0, data.Length);
}
``` 