## ExcelComponent

I make this tool based on NuGet package NPOI.
There are only basic functions to use :
1. export excel file from DataTable
2. export excel file from DataSet(with multiple sheet)
3. export excel file from DataModel
4. read DataTable data from excel file
5. read DataSet data from excel file(with multiple sheet)
6. read DataModel data from excel file

If you wants to use this component to read/write excel file, you can see the sample code in TestProject1 project.There are all unit test about this component.

Now you can define your data model as bellow, when it exports file from data model, it will use "DisplayName" as column name in excel file automatically.
```csharp
public class TestModel
{
  public string ID { get; set; }
  
  [DisplayName("Test Name")]
  public string Name { get; set; }
}
```
