using System;
using System.Linq;
using System.IO;
using System.Data;
using System.Reflection;
using System.Collections.Generic;
using NUnit.Framework;
using MyLib;

namespace TestProject1
{
    public class Tests
    {
        private class Student
        {
            public string Name { get; set; }
            public int StudentId { get; set; }
            public int? Age { get; set; }
        }
        private DataTable dtStudent = new DataTable();

        private string folder = @"D:\Test\";

        private List<Student> Students = new List<Student>();
        [SetUp]
        public void Setup()
        {
            this.Students.Add(new Student() { Name = "Jack", Age = 15, StudentId = 100 });
            this.Students.Add(new Student() { Name = "Smith", Age = 15, StudentId = 101 });
            this.Students.Add(new Student() { Name = "Smit", Age = null, StudentId = 102 });
            dtStudent = this.ToDataTable(Students);
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
            Assert.Pass();
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
            Assert.Pass();
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
            Assert.Pass();
        }

        [Test]
        public void TestReadFile2DataTable()
        {
            string filepath = this.folder + "test.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            DataTable dt = myexcel.readFileDT(fs);
            Console.WriteLine(dt);
            Assert.Pass();
        }

        [Test]
        public void TestReadFile2DataSet()
        {
            string filepath = this.folder + "test2.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            DataSet dt = myexcel.readFileDS(fs);
            Console.WriteLine(dt);
            Assert.Pass();
        }

        [Test]
        public void TestReadFile2DataModel()
        {
            string filepath = this.folder + "test.xlsx";
            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            ExcelComponent myexcel = new ExcelComponent();
            List<Student> dm = myexcel.readFileDM<Student>(fs, 0, 0);
            Console.WriteLine(dm);
            Assert.Pass();
        }

        private DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }
    }
}