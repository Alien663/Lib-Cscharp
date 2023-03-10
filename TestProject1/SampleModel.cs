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
        public int Age { get; set; }
    }
}
