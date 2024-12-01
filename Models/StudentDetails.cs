using System.ComponentModel.DataAnnotations;

namespace VadaanyaTalentTest1.Models
{
    public class StudentDetails
    {
        [Required]
        public long aadhaarNumber { get; set; }

        [Required]
        public string studentName { get; set; }

        [Required]
        public string fatherName { get; set; }

        [Required]
        public string gender { get; set; }

        [Required]
        public string mobileNumber { get; set; }

        [Required]
        public string district { get; set; }

        [Required]
        public string mandal { get; set; }
        public string testScore { get; set; }
        public long applicationNumber { get; set; }

        [Required]
        public string dob { get; set; }

        [Required]
        public string email { get; set; }

        [Required]
        public string caste { get; set; }
    }

    public enum _District
    {
        SriSathyaSai = 1,
        Anantapur = 2,
    }
}
