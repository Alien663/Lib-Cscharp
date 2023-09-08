using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestMyLib
{
    public class Student
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
}
