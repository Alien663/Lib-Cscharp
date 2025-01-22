using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using System;
using System.ComponentModel;
using System.Numerics;

namespace TestMyLib
{
    public class StudentModel
    {
        [DisplayName("Student Name")]
        public string Name { get; set; }
        [DisplayName("Student ID")]
        public int StudentId { get; set; }
        [DisplayName("Student Age")]
        public double Age { get; set; }
        [DisplayName("Birth Date")]
        public DateOnly Birth { get; set; } = DateOnly.FromDateTime(DateTime.Now);
        [DisplayName("Test Time")]
        public TimeOnly TestTime { get; set; } = TimeOnly.FromDateTime(DateTime.Now);
        [DisplayName("Last Update")]
        public DateTime UpdateTime { get; set; } = DateTime.Now;
        
    }
    public class DataTypeTestModel
    {
        public Guid guid { get; set; } = Guid.NewGuid();
        public int Int {  get; set; }
        public string String { get; set; }
        public DateTime DateTime { get; set; }
        public DateOnly DateOnly { get; set; }
        public TimeOnly TimeOnly { get; set; }
        public Int16 Int16 { get; set; }
        public Int32 Int32 { get; set; }
        public Int64 Int64 { get; set; }
        public UInt16 UInt16 { get; set; }
        public UInt32 UInt32 { get; set; }
        public UInt64 UInt64 { get; set; }
        public Boolean Flag1 { get; set; }
        public bool Flag2 { get; set; }
        public double Double { get; set; }
        public float Float { get; set; }
        public Decimal Decimal { get; set; }
        public Single Single { get; set; }
        public BigInteger BigInteger { get; set; }
    }
}
