namespace VadaanyaTalentTest1
{
    public class StudentDetails
    {
        public long aadhaarNumber { get; set; }
        public string studentName { get; set; }
        public string fatherName { get; set; }
        public string gender { get; set; }

        public string mobileNumber { get; set; }

        public string district { get; set; }
        public string mandal { get; set; }
        public string testScore { get; set; }
        public string dob { get; set; }

        public string email { get; set; }
    }

    public enum _District
    {
        SriSathyaSai =1,
        Anantapur = 2,
    }
}
