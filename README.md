## ExcelComponent

I make this tool based on NuGet package NPOI.
There are only basic functions to use :
1. export excel file from DataTable
2. export excel file from DataSet(with multiple sheet)
3. export excel file from DataModel
4. read DataTable data from excel file
5. read DataSet data from excel file(with multiple sheet)
6. read DataModel data from excel file (Work on it now)

If you wants to use this component to read/write excel file, you can see the sample code in TestProject1 project.There are all unit test about this component.

* Now you can define your data model as bellow, when it exports file from data model, it will use "DisplayName" as column name in excel file automatically.

```csharp
public class TestModel
{
  public string ID { get; set; }
  
  [DisplayName("Test Name")]
  public string Name { get; set; }
}
```


* add isHeader

If you set isHeader is false, it won't create the header row when export. 
Only give this variable when you don't want to set header automatically beacause the default is true.

```csharp
ExcelComponent myexcel = new ExcelComponent();
// set headerIndex = -1, then the data will start at headerIndex = 0
byte[] data = myexcel.export(this.dtStudent, headerIndex:-1, isHeader:false);
using (FileStream fs = File.Create(this.folder + "test.xlsx"))
{
    fs.Write(data, 0, data.Length);
}
```
