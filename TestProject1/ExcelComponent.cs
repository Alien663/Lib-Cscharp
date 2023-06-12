using System;
using System.Linq;
using System.IO;
using System.Data;
using System.Reflection;
using System.Collections.Generic;
using NUnit.Framework;
using MyLib;
using System.ComponentModel;

namespace TestMyLib
{
    public class TestExportExcel
    {
        private DataTable dtStudent = new DataTable();
        private readonly string folder = @"D:\Test\";
        private List<Student> Students = new List<Student>();
        [SetUp]
        public void Setup()
        {
            DataModelExtensions dmConvertor = new DataModelExtensions();
            Students.Add(new Student() { Name = "Jack", Age = 15.00, StudentId = 100 });
            Students.Add(new Student() { Name = "Smith", Age = 17.02, StudentId = 101 });
            Students.Add(new Student() { Name = "Karoro", Age = 20.321, StudentId = 102 });
            dtStudent = dmConvertor.ToDataTable(Students);
        }

        [Test]
        public void DataTable2Excel()
        {
            ExcelComponent myexcel = new ExcelComponent();
            byte[] data = myexcel.export(this.dtStudent);
            using (FileStream fs = File.Create(this.folder + "test.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
            Assert.Pass("DataTable export out excel file success");
        }

        [Test]
        public void DataSet2Excel()
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
            Assert.Pass(message:"DataSet export out excel file success");
        }

        [Test]
        public void DataModel2Excel()
        {
            ExcelComponent myexcel = new ExcelComponent();
            var data = myexcel.export(this.Students);
            using (FileStream fs = File.Create(this.folder + "test3.xlsx"))
            {
                fs.Write(data, 0, data.Length);
            }
            Assert.Pass("DataModel export out excel file success");
        }
    }

    public class TestReadFromExcel
    {
        private readonly string folder = @"D:\Test\";

        [Test]
        public void TestReadFile2DataTable()
        {
            string filepath = this.folder + "test.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            DataTable dt = myexcel.readFileDT(fs);
            if(dt.Rows.Count > 0)
            {
                Assert.Pass("It can read data from excel to DataTable");
            }
            else
            {
                Assert.Fail("DataTable is empty");
            }
        }

        [Test]
        public void TestReadFile2DataSet()
        {
            string filepath = this.folder + "test2.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            DataSet ds = myexcel.readFileDS(fs);
            if (ds.Tables.Count > 0)
            {
                Assert.Pass("It can read data from exlcel to DataSet");
            }
            else
            {
                Assert.Fail("DataSet is empty");
            }
            
        }

        [Test]
        public void TestReadFile2DataModel()
        {
            string filepath = this.folder + "test.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            List<Student> dm = myexcel.readFileDM<Student>(fs, 0, 0);
            if(dm.Count > 0)
            {
                Assert.Pass("It can read data from excel to DataModel list");
            }
            else
            {
                Assert.Fail("DataModel is empty");
            }
        }
    }
}