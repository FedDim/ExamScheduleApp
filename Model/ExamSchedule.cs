using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExamScheduleApp.Model
{
    public class ExamSchedule
    {
        public string Surname { get; set; }
        public string SecondSurname { get; set; }
        public DateTime? ExamDate { get; set; }
        public string Subject { get; set; }
        public string Group { get; set; }
        public string ExamTime { get; set; }
        public string Classroom { get; set; }
        public string ExamType { get; set; }

        public ExamSchedule(string surname, string secondSurname, DateTime? examDate, string subject, string group, string time, string classroom, string type)
        {
            Surname = surname;
            SecondSurname = secondSurname;
            ExamDate = examDate;
            Subject = subject;
            Group = group;
            ExamTime = time;
            Classroom = classroom;
            ExamType = type;
        }
    }
}
