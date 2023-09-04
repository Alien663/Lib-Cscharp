using MathNet.Numerics.Distributions;
using MyLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using NUnit.Framework;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Org.BouncyCastle.Asn1.BC;

namespace TestMyLib
{
    public class TestDataExtensions
    {
        private DataTable dtStudent = new DataTable();
        private List<Student> dmStudent = new List<Student>();
        [SetUp]
        public void Setup()
        {
            dtStudent = new DataTable();
            dtStudent.Columns.AddRange(new DataColumn[3] {
                new DataColumn("Name"), 
                new DataColumn("StudentId", Type.GetType("System.Int32")), 
                new DataColumn("Age", Type.GetType("System.Int32")) });
            dtStudent.Rows.Add("Jack", 15, 100);
            dtStudent.Rows.Add("Smith", 17, 101);
            dtStudent.Rows.Add("Karoro", 20, 102);

            dmStudent.Add(new Student() { Name = "Jack", Age = 15, StudentId = 100 });
            dmStudent.Add(new Student() { Name = "Smith", Age = 17, StudentId = 101 });
            dmStudent.Add(new Student() { Name = "Karoro", Age = 20, StudentId = 102 });
        }

        [Test]
        public void Table2Model()
        {
            List<Student> dmData = (List<Student>)dtStudent.ToList<Student>();
            if(dmData.Count == 0)
            {
                Assert.Fail("DataModel is empty");
            }
        }


        [Test]
        public void Model2Table()
        {
            DataModelExtensions dmConvertor = new DataModelExtensions();
            DataTable dtData = dmConvertor.ToDataTable(dmStudent);
            if (dtData.Rows.Count == 0)
            {
                Assert.Fail("DataModel is empty");
            }
        }
    }
}
