using ExcelConverter;
using NPOI.SS.Formula.Functions;
using NUnit.Framework;
using Org.BouncyCastle.Asn1.BC;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using TableConverter;
using static ICSharpCode.SharpZipLib.Zip.ExtendedUnixData;

namespace TestMyLib
{
    public class Test1_ExportExcel
    {
        private DataTable dtStudent = new DataTable();
        private readonly string folder = @".\";
        private List<Student> Students = new List<Student>();
        [SetUp]
        public void Setup()
        {
            DataModelExtensions dmConvertor = new DataModelExtensions();
            Students = new List<Student>
            {
                new Student { Name = "Jack", Age = 15.00, StudentId = 10000 },
                new Student { Name = "Smith", Age = 17.02, StudentId = 10100 },
                new Student { Name = "Keroro", Age = 20.321, StudentId = 10200 }
            };
            dtStudent = dmConvertor.ToDataTable(Students);
        }

        [Test]
        public void Test00_DataTypeTest()
        {
            List<DataTypeTest> list = new List<DataTypeTest>
            {
                new DataTypeTest
                {
                    Int=0,
                    String="Test",
                    DateTime=DateTime.Now,
                    DateOnly=DateOnly.FromDateTime(DateTime.Now),
                    TimeOnly = TimeOnly.FromDateTime(DateTime.Now),
                    Int16 = 1,
                    Int32 = 2,
                    Int64 = 3,
                    UInt16 = 4,
                    UInt32 = 5,
                    UInt64 = 6,
                    Flag1 = false,
                    Flag2 = true,
                    Double = 0.1,
                    Float = 0.2F,
                    Decimal = 0.3M,
                    Single = 0.4F,
                    BigInteger = 10
                }
            };
            ExcelComponent myexcel = new ExcelComponent();
            byte[] data = myexcel.export(list);
            using (FileStream fs = File.Create(this.folder + "datatype1.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }

        [Test]
        public void Test000_DataTypeTest()
        {
            List<DataTypeTest> list = new List<DataTypeTest>
            {
                new DataTypeTest
                {
                    Int=0,
                    String="Test",
                    DateTime=DateTime.Now,
                    DateOnly=DateOnly.FromDateTime(DateTime.Now),
                    TimeOnly = TimeOnly.FromDateTime(DateTime.Now),
                    Int16 = 1,
                    Int32 = 2,
                    Int64 = 3,
                    UInt16 = 4,
                    UInt32 = 5,
                    UInt64 = 6,
                    Flag1 = false,
                    Flag2 = true,
                    Double = 0.1,
                    Float = 0.2F,
                    Decimal = 0.3M,
                    Single = 0.4F,
                    BigInteger = 10
                }
            };
            DataModelExtensions dmConvertor = new DataModelExtensions();
            ExcelComponent myexcel = new ExcelComponent();
            byte[] data = myexcel.export(dmConvertor.ToDataTable(list));
            using (FileStream fs = File.Create(this.folder + "datatype2.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }

        [Test]
        public void Test01_DataTable2Excel()
        {
            ExcelComponent myexcel = new ExcelComponent();
            byte[] data = myexcel.export(this.dtStudent);
            using (FileStream fs = File.Create(this.folder + "test1.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        [Test]
        public void Test02_DataTable2Excel_StartWith()
        {
            ExcelComponent myexcel = new ExcelComponent();
            myexcel.SetRange(new SheetRange
            {
                MinRowIndex = 2,
                MinColIndex = 2,
            });
            byte[] data = myexcel.export(this.dtStudent);
            using (FileStream fs = File.Create(this.folder + "test1_StartWith.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }

        [Test]
        public void Test03_DataSet2Excel()
        {
            DataTable dtstudent2 = this.dtStudent.Copy();
            dtstudent2.TableName = "hahaha";
            DataSet st = new DataSet();
            st.Tables.Add(this.dtStudent);
            st.Tables.Add(dtstudent2);
            ExcelComponent myexcel = new ExcelComponent();
            byte[] data = myexcel.export(st);
            using (FileStream fs = File.Create(this.folder + "test2.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }

        [Test]
        public void Test04_DataSet2Excel_StartWith()
        {
            DataTable dtstudent2 = this.dtStudent.Copy();
            dtstudent2.TableName = "hahaha";
            DataSet st = new DataSet();
            st.Tables.Add(this.dtStudent);
            st.Tables.Add(dtstudent2);
            ExcelComponent myexcel = new ExcelComponent();
            myexcel.SetRange(new SheetRange
            {
                MinRowIndex = 2,
                MinColIndex = 2,
            });
            byte[] data = myexcel.export(st);
            using (FileStream fs = File.Create(this.folder + "test2_StartWith.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }

        [Test]
        public void Test05_DataModel2Excel()
        {
            ExcelComponent myexcel = new ExcelComponent();
            var data = myexcel.export(this.Students);
            using (FileStream fs = File.Create(this.folder + "test3.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }
        [Test]
        public void Test06_DataModel2Excel_StartWith()
        {
            ExcelComponent myexcel = new ExcelComponent();
            myexcel.SetRange(new SheetRange
            {
                MinRowIndex = 2,
                MinColIndex = 2,
            });
            var data = myexcel.export(this.Students);
            using (FileStream fs = File.Create(this.folder + "test3_StartWith.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }

        [Test]
        public void Test07_SetDataTypeStyle()
        {
            ExcelComponent myexcel = new ExcelComponent();
            myexcel.setDataTypeStyle(new Dictionary<string, string> { { "Double", "#,##0.0000" } });
            var data = myexcel.export(this.Students);
            using (FileStream fs = File.Create(this.folder + "test4.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
        }
    }

    public class Test2_ReadFromExcel
    {
        private readonly string folder = @".\";

        [Test]
        public void Test00_ReadDataTypeToClassModel()
        {
            string filepath = this.folder + "datatype1.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            List<DataTypeTest> list =  myexcel.readFileDM<DataTypeTest>(fs);
            if(list.Count == 0)
            {
                Assert.Fail("Can not read correct data.");
            }
        }

        [Test]
        public void Test00_ReadDataTypeToDataTable()
        {
            string filepath = this.folder + "datatype2.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            DataTable dtData = myexcel.readFileDT(fs);
            if (dtData.Rows.Count == 0)
            {
                Assert.Fail("Can not read correct data.");
            }
        }

        [Test]
        public void Test01_ReadFile2DataTable()
        {
            string filepath = this.folder + "test1.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            DataTable dt = myexcel.readFileDT(fs);
            if (dt.Rows.Count == 0)
            {
                Assert.Fail("DataTable is empty");
            }
        }
        [Test]
        public void Test02_ReadFile2DataTable_StartWith()
        {
            string filepath = this.folder + "test1_StartWith.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            myexcel.SetRange(new SheetRange
            {
                MinRowIndex = 2,
                MinColIndex = 2,
            });
            DataTable dt = myexcel.readFileDT(fs);
            if (dt.Rows.Count == 0)
            {
                Assert.Fail("DataTable is empty");
            }
        }
        [Test]
        public void Test03_ReadFile2DataSet()
        {
            string filepath = this.folder + "test2.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            DataSet ds = myexcel.readFileDS(fs);
            if (ds.Tables.Count == 0)
            {
                Assert.Fail("DataSet is empty");
            }
        }
        [Test]
        public void Test04_ReadFile2DataSet_StartWith()
        {
            string filepath = this.folder + "test2_StartWith.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            myexcel.SetRange(new SheetRange
            {
                MinRowIndex = 2,
                MinColIndex = 2,
            });
            DataSet ds = myexcel.readFileDS(fs);
            if (ds.Tables.Count == 0)
            {
                Assert.Fail("DataSet is empty");
            }
        }

        [Test]
        public void Test05_ReadFile2DataModel()
        {
            string filepath = this.folder + "test3.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            List<Student> dm = myexcel.readFileDM<Student>(fs);
            if (dm.Count == 0)
            {
                Assert.Fail("DataModel is empty");
            }
        }

        [Test]
        public void Test06_ReadFile2DataModel_StartWith()
        {
            string filepath = this.folder + "test3_StartWith.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            myexcel.SetRange(new SheetRange
            {
                MinRowIndex = 2,
                MinColIndex = 2,
            });
            List<Student> dm = myexcel.readFileDM<Student>(fs);
            if (dm.Count == 0)
            {
                Assert.Fail("DataModel is empty");
            }
        }
    }
}