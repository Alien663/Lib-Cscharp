using NUnit.Framework;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using Newtonsoft.Json;
using NUnit.Framework.Legacy;
using Alien.Common.Excel;

namespace TestMyLib;

[TestFixture, Order(1)]
public class ExcelConverter_DataTable
{
    private DataTable dtRawData = new DataTable();

    [OneTimeSetUp]
    public void Initailize()
    {
        dtRawData.TableName = "Test 1";
        dtRawData.Columns.Add("StudentId", typeof(int));
        dtRawData.Columns.Add("Name", typeof(string));
        dtRawData.Columns.Add("Age", typeof(double));
        dtRawData.Rows.Add(10000, "Jack", 15.00);
        dtRawData.Rows.Add(10100, "Smith", 17.02);
        dtRawData.Rows.Add(10200, "Keroro", 20.321);
    }

    [Test, Order(1)]
    public void DataTable2Excel()
    {
        #region Arrange
        string filename = @".\Test01_DataTable2Excel.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            byte[] data = excel.export(dtRawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            DataTable result = excel.readFileDT(fs2);
            for (int i = 0; i < result.Rows.Count; i++)
            {
                for (int j = 0; j < result.Columns.Count; j++)
                {
                    Assert.That(result.Rows[i][j].ToString(), Is.EqualTo(dtRawData.Rows[i][j].ToString()));
                }
            }
        }
        #endregion
    }

    [Test, Order(2)]
    public void DataTable2Excel_Anchor()
    {
        #region Arrange
        string filename = @".\Test02_DataTable2Excel_Anchor.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setAnchor(2, 3);
            byte[] data = excel.export(dtRawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            DataTable result = excel.readFileDT(fs2);
            for (int i = 0; i < result.Rows.Count; i++)
            {
                for (int j = 0; j < result.Columns.Count; j++)
                {
                    Assert.That(result.Rows[i][j].ToString(), Is.EqualTo(dtRawData.Rows[i][j].ToString()));
                }
            }
        }
        #endregion
    }

    [Test, Order(3)]
    public void DataTable2Excel_DataType()
    {
        #region Arrange
        string filename = @".\Test07_DataTable2Excel.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setDataTypeStyle(new Dictionary<string, string> { { "Double", "#,##0.0000" } });
            byte[] data = excel.export(dtRawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            DataTable result = excel.readFileDT(fs2);
            for (int i = 0; i < result.Rows.Count; i++)
            {
                for (int j = 0; j < result.Columns.Count; j++)
                {
                    Assert.That(result.Rows[i][j].ToString(), Is.EqualTo(dtRawData.Rows[i][j].ToString()));
                }
            }
        }
        #endregion
    }

    [Test, Order(6)]
    public void Excel2DataTable_DataRange()
    {
        #region Arrange
        string filename = @".\Test01_DataTable2Excel.xlsx";
        FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        DataTable result = new DataTable();
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setDataRange(3, 1);
            result = excel.readFileDT(fs);
        }
        #endregion

        #region Assert
        Assert.That(result.Rows.Count == 1);
        Assert.That(result.Columns.Count == 3);
        Assert.That(result.Rows[0]["Name"].ToString() == "Jack");
        Assert.That(result.Rows[0]["StudentId"].ToString() == "10000");
        Assert.That(result.Rows[0]["Age"].ToString() == "15");
        #endregion
    }
    
    [OneTimeTearDown]
    public void CleanFile()
    {
        File.Delete(@".\Test01_DataTable2Excel.xlsx");
        File.Delete(@".\Test02_DataTable2Excel_Anchor.xlsx");
        File.Delete(@".\Test07_DataTable2Excel.xlsx");
    }
}

[TestFixture, Order(2)]
public class ExcelConverter_DataSet
{
    private DataTable dtRawData1 = new DataTable();
    private DataTable dtRawData2 = new DataTable();
    private DataSet dsRawData = new DataSet();

    [OneTimeSetUp]
    public void Initailize()
    {
        dtRawData1.TableName = "Test 1";
        dtRawData1.Columns.Add("StudentId", typeof(int));
        dtRawData1.Columns.Add("Name", typeof(string));
        dtRawData1.Columns.Add("Age", typeof(double));
        dtRawData1.Rows.Add(10000, "Jack", 15.00);
        dtRawData1.Rows.Add(10100, "Smith", 17.02);
        dtRawData1.Rows.Add(10200, "Keroro", 20.321);

        dtRawData2.TableName = "Test 2";
        dtRawData2.Columns.Add("StudentId", typeof(int));
        dtRawData2.Columns.Add("Name", typeof(string));
        dtRawData2.Columns.Add("Age", typeof(double));
        dtRawData2.Rows.Add(10300, "Rose", 14.00);
        dtRawData2.Rows.Add(10400, "Ted", 16.01);
        dtRawData2.Rows.Add(10500, "Tamama", 21.123);

        dsRawData.Tables.Add(dtRawData1);
        dsRawData.Tables.Add(dtRawData2);
    }

    [Test, Order(1)]
    public void DataSet2Excel()
    {
        #region Arrange
        string filename = @".\Test01_DataSet2Excel.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            byte[] data = excel.export(dsRawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            DataSet result = excel.readFileDS(fs2);
            for(int k = 0; k < result.Tables.Count; k++)
            {
                for (int i = 0; i < result.Tables[k].Rows.Count; i++)
                {
                    for (int j = 0; j < result.Tables[k].Columns.Count; j++)
                    {
                        Assert.That(result.Tables[k].Rows[i][j].ToString(), Is.EqualTo(dsRawData.Tables[k].Rows[i][j].ToString()));
                    }
                }
            }
        }
        #endregion
    }

    [Test, Order(2)]
    public void DataSet2Excel_Anchor()
    {
        #region Arrange
        string filename = @".\Test02_DataSet2Excel_Anchor.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setAnchor(2, 3);
            byte[] data = excel.export(dsRawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            DataSet result = excel.readFileDS(fs2);
            for (int k = 0; k < result.Tables.Count; k++)
            {
                for (int i = 0; i < result.Tables[k].Rows.Count; i++)
                {
                    for (int j = 0; j < result.Tables[k].Columns.Count; j++)
                    {
                        Assert.That(result.Tables[k].Rows[i][j].ToString(), Is.EqualTo(dsRawData.Tables[k].Rows[i][j].ToString()));
                    }
                }
            }
        }
        #endregion
    }

    [Test, Order(3)]
    public void DataSet2Excel_SheetRange()
    {
        #region Arrange
        string filename = @".\Test04_DataSet2Excel_SheetRange.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setSheetRange(0, 1);
            byte[] data = excel.export(dsRawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            DataSet result = excel.readFileDS(fs2);
            for (int k = 0; k < result.Tables.Count; k++)
            {
                for (int i = 0; i < result.Tables[k].Rows.Count; i++)
                {
                    for (int j = 0; j < result.Tables[k].Columns.Count; j++)
                    {
                        Assert.That(result.Tables[k].Rows[i][j].ToString(), Is.EqualTo(dsRawData.Tables[k].Rows[i][j].ToString()));
                    }
                }
            }
        }
        #endregion
    }

    [Test, Order(4)]
    public void DataSet2Excel_DataType()
    {
        #region Arrange
        string filename = @".\Test09_Excel2DataSet_DataType.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setDataTypeStyle(new Dictionary<string, string> { { "Double", "#,##0.0000" } });
            byte[] data = excel.export(dsRawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            DataSet result = excel.readFileDS(fs2);
            for (int k = 0; k < result.Tables.Count; k++)
            {
                for (int i = 0; i < result.Tables[k].Rows.Count; i++)
                {
                    for (int j = 0; j < result.Tables[k].Columns.Count; j++)
                    {
                        Assert.That(result.Tables[k].Rows[i][j].ToString(), Is.EqualTo(dsRawData.Tables[k].Rows[i][j].ToString()));
                    }
                }
            }
        }
        #endregion
    }

    [Test, Order(7)]
    public void Excel2DataSet_DataRange()
    {
        #region Arrange
        string filename = @".\Test01_DataSet2Excel.xlsx";
        FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        DataSet result = new DataSet();
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setDataRange(3, 1);
            result = excel.readFileDS(fs);
        }
        #endregion

        #region Assert
        Assert.That(result.Tables.Count, Is.EqualTo(2));
        using (ExcelConverter excel = new ExcelConverter())
        {
            for (int k = 0; k < result.Tables.Count; k++)
            {
                for (int i = 0; i < result.Tables[k].Rows.Count; i++)
                {
                    for (int j = 0; j < result.Tables[k].Columns.Count; j++)
                    {
                        Assert.That(result.Tables[k].Rows[i][j].ToString(), Is.EqualTo(dsRawData.Tables[k].Rows[i][j].ToString()));
                    }
                }
            }
        }
        #endregion
    }

    [Test, Order(8)]
    public void Excel2DataSet_SheetRange()
    {
        #region Arrange
        string filename = @".\Test04_DataSet2Excel_SheetRange.xlsx";
        FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        DataSet result = new DataSet();
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            result = excel.readFileDS(fs);
        }
        #endregion

        #region Assert
        for (int k = 0; k < result.Tables.Count; k++)
        {
            for (int i = 0; i < result.Tables[k].Rows.Count; i++)
            {
                for (int j = 0; j < result.Tables[k].Columns.Count; j++)
                {
                    Assert.That(result.Tables[k].Rows[i][j].ToString() == dsRawData.Tables[k].Rows[i][j].ToString());
                }
            }
        }
        Console.WriteLine(JsonConvert.SerializeObject(result));
        #endregion
    }

    [OneTimeTearDown]
    public void CleanFile()
    {
        File.Delete(@".\Test01_DataSet2Excel.xlsx");
        File.Delete(@".\Test02_DataSet2Excel_Anchor.xlsx");
        File.Delete(@".\Test04_DataSet2Excel_SheetRange.xlsx");
        File.Delete(@".\Test09_Excel2DataSet_DataType.xlsx");
    }
}

[TestFixture, Order(3)]
public class ExcelConverter_ClassModel
{
    private List<StudentModel> rawData;

    [OneTimeSetUp]
    public void Initailize()
    {
        rawData = new List<StudentModel>
        {
            new StudentModel {Name = "Jack", Age = 15, StudentId = 100},
            new StudentModel {Name = "Smith", Age = 17, StudentId = 101 },
            new StudentModel {Name = "Keroro", Age = 20, StudentId = 102 },
        };
    }

    [Test, Order(1)]
    public void DataModel2Excel()
    {
        #region Arrange
        string filename = @".\Test01_DataModel2Excel.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            byte[] data = excel.export(rawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            List<StudentModel> result = excel.readFileDM<StudentModel>(fs2);
            for(int i=0; i < result.Count; i++)
            {
                Assert.That(result[i].StudentId, Is.EqualTo(rawData[i].StudentId));
                Assert.That(result[i].Name, Is.EqualTo(rawData[i].Name));
                Assert.That(result[i].Age, Is.EqualTo(rawData[i].Age));
            }
        }
        #endregion
    }

    [Test, Order(2)]
    public void DataModel2Excel_Anchor()
    {
        #region Arrange
        string filename = @".\Test02_DataModel2Excel_Anchor.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setAnchor(2, 3);
            byte[] data = excel.export(rawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            List<StudentModel> result = excel.readFileDM<StudentModel>(fs2);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.That(result[i].StudentId, Is.EqualTo(rawData[i].StudentId));
                Assert.That(result[i].Name, Is.EqualTo(rawData[i].Name));
                Assert.That(result[i].Age, Is.EqualTo(rawData[i].Age));
            }
        }
        #endregion
    }

    [Test, Order(3)]
    public void DataModel2Excel_DataType()
    {
        #region Arrange
        string filename = @".\Test07_Excel2DataModel_DataType.xlsx";
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setDataTypeStyle(new Dictionary<string, string> { { "Double", "#,##0.0000" } });
            byte[] data = excel.export(rawData);
            using (FileStream fs = File.Create(filename))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        #endregion

        #region Assert
        FileAssert.Exists(filename);
        FileStream fs2 = new FileStream(filename, FileMode.Open, FileAccess.Read);
        using (ExcelConverter excel = new ExcelConverter())
        {
            List<StudentModel> result = excel.readFileDM<StudentModel>(fs2);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.That(result[i].StudentId, Is.EqualTo(rawData[i].StudentId));
                Assert.That(result[i].Name, Is.EqualTo(rawData[i].Name));
                Assert.That(result[i].Age, Is.EqualTo(rawData[i].Age));
            }
        }
        #endregion
    }

    [Test, Order(6)]
    public void Excel2DataModel_DataRange()
    {
        #region Arrange
        string filename = @".\Test01_DataModel2Excel.xlsx";
        FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        List<StudentModel> result = new List<StudentModel>();
        #endregion

        #region Act
        using (ExcelConverter excel = new ExcelConverter())
        {
            excel.setDataRange(3, 1);
            result = excel.readFileDM<StudentModel>(fs);
        }
        #endregion

        #region Assert
        Assert.That(result.Count == 1);
        Assert.That(rawData[0].Name == result[0].Name);
        Assert.That(rawData[0].StudentId == result[0].StudentId);
        Assert.That(rawData[0].Age == result[0].Age);
        Console.WriteLine(JsonConvert.SerializeObject(result));
        #endregion
    }


    [OneTimeTearDown]
    public void CleanFile()
    {
        File.Delete(@".\Test01_DataModel2Excel.xlsx");
        File.Delete(@".\Test02_DataModel2Excel_Anchor.xlsx");
        File.Delete(@".\Test07_Excel2DataModel_DataType.xlsx");
    }
}
