namespace ExamScheduleApp.Model
{
    public class ExamSchedule
    {
        public string FirstTeacher { get; set; }
        public string SecondTeacher { get; set; }
        public string ExamDate { get; set; }
        public string Subject { get; set; }
        public string Group { get; set; }
        public string ExamTime { get; set; }
        public string Classroom { get; set; }
        public string ExamType { get; set; }

        public ExamSchedule(string firstTeacher, string secondTeacher, string examDate, string subject, string group, string time, string classroom, string type)
        {
            FirstTeacher = firstTeacher;
            SecondTeacher = secondTeacher;
            ExamDate = examDate;
            Subject = subject;
            Group = group;
            ExamTime = time;
            Classroom = classroom;
            ExamType = type;
        }
    }
}
