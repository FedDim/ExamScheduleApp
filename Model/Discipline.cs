namespace ExamScheduleApp.Model
{
    public class Subject
    {
        public int Id { get; set; }
        public string FullName { get; set; }
        public string ShortName12 { get; set; }
        public string ShortName9 { get; set; }
        public string ShortName5 { get; set; }

        public string Name => ShortName9;
        public string Code => ShortName5;

        public Subject() { }
    }
}
