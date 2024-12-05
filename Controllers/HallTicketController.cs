using Microsoft.AspNetCore.Mvc;
using VadaanyaTalentTest1.Models;
using VadaanyaTalentTest1.Handlers;

namespace VadaanyaTalentTest1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class HallTicketController: BaseController
    {
        private readonly ILogger<HallTicketController> _logger;
        private readonly IConfiguration _configuration;
        private static readonly string excelSheetFileName = "DSCAPPS";
        private static readonly string templateHallTicketFileName = "HallTicket-DSC2024";
        //private string excelSheetPath = Path.Combine(Directory.GetCurrentDirectory(), "Static", excelSheetFileName + ".xlsx");
        private string excelSheetPath = Path.Combine(Directory.GetCurrentDirectory(), "Static", excelSheetFileName + ".xlsx").Replace(Path.DirectorySeparatorChar == '\\' ? '\\' : '/', Path.DirectorySeparatorChar);
        //private string inPath = Path.Combine(Directory.GetCurrentDirectory(), "Static", templateHallTicketFileName + ".pdf");
        private string inPath = Path.Combine(Directory.GetCurrentDirectory(), "Static", templateHallTicketFileName + ".pdf").Replace(Path.DirectorySeparatorChar == '\\' ? '\\' : '/', Path.DirectorySeparatorChar);
        //private string outPath = Path.Combine(Directory.GetCurrentDirectory(), "Static", templateHallTicketFileName + "_GENERATED.pdf");
        private string outPath = Path.Combine(Directory.GetCurrentDirectory(), "Static", templateHallTicketFileName + "_GENERATED.pdf").Replace(Path.DirectorySeparatorChar == '\\' ? '\\' : '/', Path.DirectorySeparatorChar);
        private HallticketHandler _hallticketHandler;


        public HallTicketController(ILogger<HallTicketController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
            _hallticketHandler = new HallticketHandler(excelSheetPath);
        }


        // ROUTES

        [HttpGet]
        [Route("DSC-2024")]
        public async Task<IActionResult> GetDSC2024HallTicket(
            [FromQuery] long applicationNumber, 
            [FromQuery] long aadhaarNumber, 
            [FromQuery] string dob
            )
        {
            try
            {                
                Dictionary<string, string> applicant = null;                

                applicant = _hallticketHandler.FilterDSC2024ApplicantDetails(applicationNumber, aadhaarNumber, dob);

                _hallticketHandler.GenerateDSC2024Hallticket(applicant, inPath, outPath);

                var memory = new MemoryStream();
                using (var stream = new FileStream(outPath, FileMode.Open))
                {
                    await stream.CopyToAsync(memory);
                }
                memory.Position = 0;

                var downloadFileName = applicant["ApplicationNumber"] + "_HALLTICKET" + ".pdf";

                Response.Headers.Append("X-File-Name", downloadFileName);

                return File(memory, "application/pdf");
            }
            catch(Exception ex)
            {
                return HandleError(ex);
            }
            
        }


        // GET THE APPLICANT DETAILS 

        private List<StudentDetails> GetAllStudentDetails()
        {
            //Dummy Student Details Data
            var studentDetails = new List<StudentDetails>
        {
            new StudentDetails
            {
                aadhaarNumber = 920471536585,
                studentName = "Chandra Mouli",
                fatherName = "Madhusudana",
                gender = "Male",
                mobileNumber = "9885533513",
                district = _District.SriSathyaSai.ToString(),
                mandal = "KOTHACHERUVU",
                dob = "2001-07-14",
                email = "chandramouli3636@gmail.com",
                testScore = "98",
                caste = "OC",
                applicationNumber = 2912240002
            },
            new StudentDetails
            {
                aadhaarNumber = 960130463168,
                studentName = "D Venkatanarayana",
                fatherName = "D Venkataramudu",
                gender = "Male",
                mobileNumber = "7893132930",
                district = _District.SriSathyaSai.ToString(),
                mandal = "NALLAMADA",
                dob = "1992-06-10",
                email = "daravenkat4@gmail.com",
                testScore = "96",
                caste = "SC",
                applicationNumber = 2912240003
            },
            new StudentDetails
            {
                aadhaarNumber = 711018435988,
                studentName = "Sanvi",
                fatherName = "Sankar",
                gender = "Female",
                mobileNumber = "8074349270",
                district = _District.Anantapur.ToString(),
                mandal = "PUTLURU",
                dob = "2000-12-23",
                email = "anithayellapagari@gmail.com",
                testScore = "135",
                caste = "OC",
                applicationNumber = 2912240004
            },
            new StudentDetails
            {
                aadhaarNumber = 415783323379,
                studentName = "Nunna Vikram",
                fatherName = "Nunna Gurubrahmam",
                gender = "Male",
                mobileNumber = "6281471609",
                district = _District.Anantapur.ToString(),
                mandal = "TADIPATRI",
                dob = "1996-11-14",
                email = "vikramsai959@gmail.com",
                testScore = "113",
                caste = "BC-A",
                applicationNumber = 2912240005
            },
            new StudentDetails
            {
                aadhaarNumber = 387864514893,
                studentName = "Makkisetty Maruthi Prasad",
                fatherName = "Makkisetty sreenivasulu",
                gender = "Male",
                mobileNumber = "7842448814",
                district = _District.SriSathyaSai.ToString(),
                mandal = "BUKKAPATNAM",
                dob = "1994-06-25",
                email = "maruthiprasad156@gmail.com",
                testScore = "123",
                caste = "OC",
                applicationNumber = 2912240006
            }
        };

            return studentDetails;
        }
    }
}
