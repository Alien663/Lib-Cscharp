using NUnit.Framework;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using Newtonsoft.Json;
using Data.Extension;
using Excel.Extension;

namespace TestMyLib
{
    [TestFixture]
    public class ExcelConverter_DataTable
    {
        private List<StudentModel> dmstudents;
        private DataTable dtRawData = new DataTable();

        [OneTimeSetUp]
        public void Initailize()
        {
            dmstudents = new List<StudentModel>
            {
                new StudentModel { Name = "Jack", Age = 15.00, StudentId = 10000 },
                new StudentModel { Name = "Smith", Age = 17.02, StudentId = 10100 },
                new StudentModel { Name = "Keroro", Age = 20.321, StudentId = 10200 }
            };
            dtRawData = ClassModelConvert.ToDataTable(dmstudents);
        }

        [Test]
        public void Test01_DataTable2Excel()
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
            #endregion
        }

        [Test]
        public void Test02_DataTable2Excel_Anchor()
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
            #endregion
        }

        [Test]
        public void Test04_Excel2DataTable()
        {
            #region Arrange
            string filename = @".\Test01_DataTable2Excel.xlsx";
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            DataTable result = new DataTable();
            #endregion

            #region Act
            using (ExcelConverter excel = new ExcelConverter())
            {
                result = excel.readFileDT(fs);
            }
            #endregion

            #region Assert
            for (int i = 0; i < result.Rows.Count; i++)
            {
                for (int j = 0; j < result.Columns.Count; j++)
                {
                    Assert.AreEqual(result.Rows[i][j].ToString(), dtRawData.Rows[i][j].ToString());
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test05_Excel2DataTable_Anchor()
        {
            #region Arrange
            string filename = @".\Test02_DataTable2Excel_Anchor.xlsx";
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            DataTable result = new DataTable();
            #endregion

            #region Act
            using (ExcelConverter excel = new ExcelConverter())
            {
                excel.setAnchor(2, 3);
                result = excel.readFileDT(fs);
            }
            #endregion

            #region Assert
            for (int i = 0; i < result.Rows.Count; i++)
            {
                for (int j = 0; j < result.Columns.Count; j++)
                {
                    Assert.AreEqual(result.Rows[i][j].ToString(), dtRawData.Rows[i][j].ToString());
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test06_Excel2DataTable_DataRange()
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
            Assert.IsTrue(result.Rows.Count == 1);
            Assert.IsTrue(result.Columns.Count == 3);
            Assert.AreEqual(result.Rows[0]["Name"], "Jack");
            Assert.AreEqual(result.Rows[0]["StudentId"], "10000");
            Assert.AreEqual(result.Rows[0]["Age"], "15");
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test07_Excel2DataTable_DataType()
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
            #endregion
        }

        [Test]
        public void Test08_Excel2DataTable_DataType()
        {
            #region Arrange
            string filename = @".\Test07_DataTable2Excel.xlsx";
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            DataTable result = new DataTable();
            #endregion

            #region Act
            using (ExcelConverter excel = new ExcelConverter())
            {
                result = excel.readFileDT(fs);
            }
            #endregion

            #region Assert
            for (int i = 0; i < result.Rows.Count; i++)
            {
                for (int j = 0; j < result.Columns.Count; j++)
                {
                    Assert.AreEqual(result.Rows[i][j].ToString(), dtRawData.Rows[i][j].ToString());
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }
    }

    [TestFixture]
    public class ExcelConverter_DataSet
    {
        private List<StudentModel> dmRawData;
        private DataSet dsRawData = new DataSet();

        [OneTimeSetUp]
        public void Initailize()
        {
            dmRawData = new List<StudentModel>
            {
                new StudentModel {Name = "Jack", Age = 15, StudentId = 100},
                new StudentModel {Name = "Smith", Age = 17, StudentId = 101 },
                new StudentModel {Name = "Karoro", Age = 20, StudentId = 102 },
            };

            DataTable dtRawData = ClassModelConvert.ToDataTable(dmRawData);
            DataTable dtRawData2 = dtRawData.Copy();
            dtRawData.TableName = "Test 1";
            dtRawData2.TableName = "Test 2";
            dsRawData.Tables.Add(dtRawData);
            dsRawData.Tables.Add(dtRawData2);
        }

        [Test]
        public void Test01_DataSet2Excel()
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
            #endregion
        }

        [Test]
        public void Test02_DataSet2Excel_Anchor()
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
            #endregion
        }

        [Test]
        public void Test04_DataSet2Excel_SheetRange()
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
            #endregion
        }

        [Test]
        public void Test05_Excel2DataSet()
        {
            #region Arrange
            string filename = @".\Test01_DataSet2Excel.xlsx";
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
                        Assert.AreEqual(result.Tables[k].Rows[i][j].ToString(), dsRawData.Tables[k].Rows[i][j].ToString());
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test06_Excel2DataSet_Anchor()
        {
            #region Arrange
            string filename = @".\Test02_DataSet2Excel_Anchor.xlsx";
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            DataSet result = new DataSet();
            #endregion

            #region Act
            using (ExcelConverter excel = new ExcelConverter())
            {
                excel.setAnchor(2, 3);
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
                        Assert.AreEqual(result.Tables[k].Rows[i][j].ToString(), dsRawData.Tables[k].Rows[i][j].ToString());
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test07_Excel2DataSet_DataRange()
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
            for (int i = 0; i < result.Tables.Count; i++)
            {
                Assert.IsTrue(result.Tables[i].Rows.Count == 1);
                Assert.IsTrue(result.Tables[i].Columns.Count == 3);
                Assert.AreEqual(result.Tables[i].Rows[0]["Name"], "Jack");
                Assert.AreEqual(result.Tables[i].Rows[0]["StudentId"], "100");
                Assert.AreEqual(result.Tables[i].Rows[0]["Age"], "15");
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test08_Excel2DataSet_SheetRange()
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
                        Assert.AreEqual(result.Tables[k].Rows[i][j].ToString(), dsRawData.Tables[k].Rows[i][j].ToString());
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test09_Excel2DataSet_DataType()
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
            #endregion
        }

        [Test]
        public void Test10_Excel2DataSet_DataType()
        {
            #region Arrange
            string filename = @".\Test09_Excel2DataSet_DataType.xlsx";
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
                        Assert.AreEqual(result.Tables[k].Rows[i][j].ToString(), dsRawData.Tables[k].Rows[i][j].ToString());
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }
    }

    [TestFixture]
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
                new StudentModel {Name = "Karoro", Age = 20, StudentId = 102 },
            };
        }

        [Test]
        public void Test01_DataModel2Excel()
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
            #endregion
        }

        [Test]
        public void Test02_DataModel2Excel_Anchor()
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
            #endregion
        }

        [Test]
        public void Test04_Excel2DataModel()
        {
            #region Arrange
            string filename = @".\Test01_DataModel2Excel.xlsx";
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            List<StudentModel> result = new List<StudentModel>();
            #endregion

            #region Act
            using (ExcelConverter excel = new ExcelConverter())
            {
                result = excel.readFileDM<StudentModel>(fs);
            }
            #endregion

            #region Assert
            for(int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(rawData[i].Name, result[i].Name);   
                Assert.AreEqual(rawData[i].StudentId, result[i].StudentId);
                Assert.AreEqual(rawData[i].Age, result[i].Age);
                Assert.AreEqual(rawData[i].Birth, result[i].Birth);
                Assert.AreEqual(rawData[i].TestTime.ToString(), result[i].TestTime.ToString());
                Assert.AreEqual(rawData[i].UpdateTime.ToString(), result[i].UpdateTime.ToString());
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test05_Excel2DataModel_Anchor()
        {
            #region Arrange
            string filename = @".\Test02_DataModel2Excel_Anchor.xlsx";
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            List<StudentModel> result = new List<StudentModel>();
            #endregion

            #region Act
            using (ExcelConverter excel = new ExcelConverter())
            {
                excel.setAnchor(2, 3);
                result = excel.readFileDM<StudentModel>(fs);
            }
            #endregion

            #region Assert
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(rawData[i].Name, result[i].Name);
                Assert.AreEqual(rawData[i].StudentId, result[i].StudentId);
                Assert.AreEqual(rawData[i].Age, result[i].Age);
                Assert.AreEqual(rawData[i].Birth, result[i].Birth);
                Assert.AreEqual(rawData[i].TestTime.ToString(), result[i].TestTime.ToString());
                Assert.AreEqual(rawData[i].UpdateTime.ToString(), result[i].UpdateTime.ToString());
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test06_Excel2DataModel_DataRange()
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
            Assert.IsTrue(result.Count == 1);
            Assert.AreEqual(rawData[0].Name, result[0].Name);
            Assert.AreEqual(rawData[0].StudentId, result[0].StudentId);
            Assert.AreEqual(rawData[0].Age, result[0].Age);
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test]
        public void Test07_Excel2DataModel_DataType()
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
            #endregion
        }

        [Test]
        public void Test08_Excel2DataModel_DataType()
        {
            #region Arrange
            string filename = @".\Test07_Excel2DataModel_DataType.xlsx";
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            List<StudentModel> result = new List<StudentModel>();
            #endregion

            #region Act
            using (ExcelConverter excel = new ExcelConverter())
            {
                result = excel.readFileDM<StudentModel>(fs);
            }
            #endregion

            #region Assert
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(rawData[i].Name, result[i].Name);
                Assert.AreEqual(rawData[i].StudentId, result[i].StudentId);
                Assert.AreEqual(rawData[i].Age, result[i].Age);
                Assert.AreEqual(rawData[i].Birth, result[i].Birth);
                Assert.AreEqual(rawData[i].TestTime.ToString(), result[i].TestTime.ToString());
                Assert.AreEqual(rawData[i].UpdateTime.ToString(), result[i].UpdateTime.ToString());
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }
    }
}
