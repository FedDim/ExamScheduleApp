using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExamScheduleApp.Model
{
    public class Teacher
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Classroom { get; set; }
        public int AcademicBuilding { get; set; }

        public Teacher() { }
    }
}
