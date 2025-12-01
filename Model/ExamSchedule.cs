namespace ExamScheduleApp.Model
{
    public class ExamSchedule
    {
        public int Id { get; set; }
        public int Teacher1Id { get; set; }
        public int? Teacher2Id { get; set; }
        public string Teacher1Name { get; set; }
        public string Teacher2Name { get; set; }
        public int SubjectId { get; set; }
        public string SubjectName { get; set; }
        public int GroupId { get; set; }
        public string GroupName { get; set; }
        public string DepartmentName { get; set; }
        public string Classroom { get; set; }
        public string ExamDate { get; set; }
        public string ExamTime { get; set; }
        public string ExamType { get; set; }

        public ExamSchedule() { }

        // Конструктор для обратной совместимости
        public ExamSchedule(string surname, string secondSurname, string examDate,
                                  string subject, string group, string department,
                                  string time, string classroom, string type)
        {
            Teacher1Name = surname;
            Teacher2Name = secondSurname;
            ExamDate = examDate;
            SubjectName = subject;
            GroupName = group;
            DepartmentName = department;
            ExamTime = time;
            Classroom = classroom;
            ExamType = type;
        }
    }
}