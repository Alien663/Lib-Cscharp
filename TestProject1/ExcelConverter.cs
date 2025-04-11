using NUnit.Framework;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using Newtonsoft.Json;
using Data.Extension;
using Excel.Extension;
using NUnit.Framework.Legacy;

namespace TestMyLib
{
    [TestFixture, Order(1)]
    public class ExcelConverter_DataTable
    {
        private DataTable dtRawData = new DataTable();

        [OneTimeSetUp]
        public void Initailize()
        {
            dtRawData.TableName = "Test 1";
            #region Column Setting
            DataColumn column;
            column = new DataColumn();
            column.ColumnName = "Name";
            column.DataType = Type.GetType("System.String");
            dtRawData.Columns.Add(column);

            column = new DataColumn();
            column.ColumnName = "Age";
            column.DataType = Type.GetType("System.Double");
            dtRawData.Columns.Add(column);

            column = new DataColumn();
            column.ColumnName = "StudentId";
            column.DataType = Type.GetType("System.Int32");
            dtRawData.Columns.Add(column);
            #endregion

            #region Row Add
            DataRow row;
            row = dtRawData.NewRow();
            row["Name"] = "Jack";
            row["Age"] = 15.00;
            row["StudentId"] = 10000;
            dtRawData.Rows.Add(row);

            row = dtRawData.NewRow();
            row["Name"] = "Smith";
            row["Age"] = 17.02;
            row["StudentId"] = 10100;
            dtRawData.Rows.Add(row);

            row = dtRawData.NewRow();
            row["Name"] = "Keroro";
            row["Age"] = 20.321;
            row["StudentId"] = 10200;
            dtRawData.Rows.Add(row);
            #endregion
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
            #endregion
        }

        [Test, Order(3)]
        public void Excel2DataTable()
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
                    Assert.That(result.Rows[i][j].ToString() == dtRawData.Rows[i][j].ToString());
                }
            }
            #endregion
        }

        [Test, Order(4)]
        public void Excel2DataTable_Anchor()
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
                    Assert.That(result.Rows[i][j].ToString() == dtRawData.Rows[i][j].ToString());
                }
            }
            #endregion
        }

        [Test, Order(5)]
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

        [Test, Order(6)]
        public void Excel2DataTable_DataType()
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

        [Test, Order(7)]
        public void Excel2DataTable_DataType2()
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
                    Assert.That(result.Rows[i][j].ToString() == dtRawData.Rows[i][j].ToString());
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }
    }

    [TestFixture, Order(2)]
    public class ExcelConverter_DataSet
    {
        private List<StudentModel> dmRawData;
        private DataTable dtRawData1 = new DataTable();
        private DataTable dtRawData2 = new DataTable();
        private DataSet dsRawData = new DataSet();

        [OneTimeSetUp]
        public void Initailize()
        {
            dtRawData1.TableName = "Test 1";
            #region Column Setting
            DataColumn column;
            column = new DataColumn();
            column.ColumnName = "Name";
            column.DataType = Type.GetType("System.String");
            dtRawData1.Columns.Add(column);

            column = new DataColumn();
            column.ColumnName = "Age";
            column.DataType = Type.GetType("System.Double");
            dtRawData1.Columns.Add(column);

            column = new DataColumn();
            column.ColumnName = "StudentId";
            column.DataType = Type.GetType("System.Int32");
            dtRawData1.Columns.Add(column);
            #endregion

            #region Row Add
            DataRow row;
            row = dtRawData1.NewRow();
            row["Name"] = "Jack";
            row["Age"] = 15.00;
            row["StudentId"] = 10000;
            dtRawData1.Rows.Add(row);

            row = dtRawData1.NewRow();
            row["Name"] = "Smith";
            row["Age"] = 17.02;
            row["StudentId"] = 10100;
            dtRawData1.Rows.Add(row);

            row = dtRawData1.NewRow();
            row["Name"] = "Keroro";
            row["Age"] = 20.321;
            row["StudentId"] = 10200;
            dtRawData1.Rows.Add(row);
            #endregion
            
            dtRawData2 = dtRawData1.Copy();
            dtRawData2.TableName = "Test 2";
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
            #endregion
        }

        [Test, Order(4)]
        public void Excel2DataSet()
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
                        Assert.That(result.Tables[k].Rows[i][j].ToString() == dsRawData.Tables[k].Rows[i][j].ToString());
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test, Order(5)]
        public void Excel2DataSet_Anchor()
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
                        Assert.That(result.Tables[k].Rows[i][j].ToString() == dsRawData.Tables[k].Rows[i][j].ToString());
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test, Order(6)]
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
            for (int i = 0; i < result.Tables.Count; i++)
            {
                Assert.That(result.Tables[i].Rows.Count == 1);
                Assert.That(result.Tables[i].Columns.Count == 3);
                Assert.That(result.Tables[i].Rows[0]["Name"].ToString() == "Jack");
                Assert.That(result.Tables[i].Rows[0]["StudentId"].ToString() == "10000");
                Assert.That(result.Tables[i].Rows[0]["Age"].ToString() == "15");
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test, Order(7)]
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

        [Test, Order(8)]
        public void Excel2DataSet_DataType()
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

        [Test, Order(9)]
        public void Excel2DataSet_DataType2()
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
                        Assert.That(result.Tables[k].Rows[i][j].ToString() == dsRawData.Tables[k].Rows[i][j].ToString());
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
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
            #endregion
        }

        [Test, Order(3)]
        public void Excel2DataModel()
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
                Assert.That(rawData[i].Name == result[i].Name);   
                Assert.That(rawData[i].StudentId == result[i].StudentId);
                Assert.That(rawData[i].Age == result[i].Age);
                Assert.That(rawData[i].Birth == result[i].Birth);
                Assert.That(rawData[i].TestTime.ToString() == result[i].TestTime.ToString());
                Assert.That(rawData[i].UpdateTime.ToString() == result[i].UpdateTime.ToString());
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test, Order(4)]
        public void Excel2DataModel_Anchor()
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
                Assert.That(rawData[i].Name == result[i].Name);
                Assert.That(rawData[i].StudentId == result[i].StudentId);
                Assert.That(rawData[i].Age == result[i].Age);
                Assert.That(rawData[i].Birth == result[i].Birth);
                Assert.That(rawData[i].TestTime.ToString() == result[i].TestTime.ToString());
                Assert.That(rawData[i].UpdateTime.ToString() == result[i].UpdateTime.ToString());
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }

        [Test, Order(5)]
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

        [Test, Order(6)]
        public void Excel2DataModel_DataType()
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

        [Test, Order(7)]
        public void Excel2DataModel_DataType2()
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
                Assert.That(rawData[i].Name == result[i].Name);
                Assert.That(rawData[i].StudentId == result[i].StudentId);
                Assert.That(rawData[i].Age == result[i].Age);
                Assert.That(rawData[i].Birth == result[i].Birth);
                Assert.That(rawData[i].TestTime.ToString() == result[i].TestTime.ToString());
                Assert.That(rawData[i].UpdateTime.ToString() == result[i].UpdateTime.ToString());
            }
            Console.WriteLine(JsonConvert.SerializeObject(result));
            #endregion
        }
    }
}
