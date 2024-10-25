using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Net;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using Twilio;
using Twilio.Rest.Api.V2010.Account;
using Twilio.Types;

namespace VadaanyaTalentTest1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class StudentDetailsController : ControllerBase
    {
        private readonly ILogger<StudentDetailsController> _logger;
        private readonly List<StudentDetails> _studentDetails = new List<StudentDetails>();

        public StudentDetailsController(ILogger<StudentDetailsController> logger)
        {
            _logger = logger;
        }

        [HttpGet("{studentId}")]
        public ActionResult<StudentDetails> Get(long studentId)
        {
            var student = GetStudentFromExcel(studentId);
            if (student == null)
            {
                return NotFound();
            }
            return Ok(student);

        }

        [HttpPost]
        public ActionResult<StudentDetails> Post(StudentDetails student)
        {
            _studentDetails.Add(student);
            AddStudentToExcel(student);
            SendEmail(student);
            //SendWhatsAppMessage(student);
            return Ok(student);
        }

        private void AddStudentToExcel(StudentDetails student)
        {
            var filePath = "StudentDetails.xlsx";
            FileInfo fileInfo = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet;
                if (fileInfo.Exists)
                {
                    worksheet = package.Workbook.Worksheets[0];
                }
                else
                {
                    worksheet = package.Workbook.Worksheets.Add("StudentDetails");
                    worksheet.Cells[1, 1].Value = "AadhaarNumber";
                    worksheet.Cells[1, 2].Value = "StudentName";
                    worksheet.Cells[1, 3].Value = "FatherName";
                    worksheet.Cells[1, 4].Value = "Gender";
                    worksheet.Cells[1, 5].Value = "MobileNumber";
                    worksheet.Cells[1, 6].Value = "District";
                    worksheet.Cells[1, 7].Value = "Mandal";
                    worksheet.Cells[1, 8].Value = "Dob";
                    worksheet.Cells[1, 9].Value = "Email";
                }

                int row = worksheet.Dimension?.Rows + 1 ?? 2;
                worksheet.Cells[row, 1].Value = student.aadhaarNumber;
                worksheet.Cells[row, 2].Value = student.studentName;
                worksheet.Cells[row, 3].Value = student.fatherName;
                worksheet.Cells[row, 4].Value = student.gender;
                worksheet.Cells[row, 5].Value = student.mobileNumber;
                worksheet.Cells[row, 6].Value = student.district;
                worksheet.Cells[row, 7].Value = student.mandal;
                worksheet.Cells[row, 8].Value = student.dob;
                worksheet.Cells[row, 9].Value = student.email;

                package.Save();
            }
        }

        private StudentDetails GetStudentFromExcel(long studentId)
        {
            var filePath = "StudentDetails.xlsx";
            FileInfo fileInfo = new FileInfo(filePath);

            if (!fileInfo.Exists)
            {
                return null;
            }

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                if (worksheet == null)
                {
                    return null;
                }

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    if (worksheet.Cells[row, 1].GetValue<long>() == studentId)
                    {
                        return new StudentDetails
                        {
                            aadhaarNumber = worksheet.Cells[row, 1].GetValue<long>(),
                            studentName = worksheet.Cells[row, 2].GetValue<string>(),
                            fatherName = worksheet.Cells[row, 3].GetValue<string>(),
                            gender = worksheet.Cells[row, 4].GetValue<string>(),
                            mobileNumber = worksheet.Cells[row, 5].GetValue<string>(),
                            district = worksheet.Cells[row, 5].GetValue<string>(),
                            mandal = worksheet.Cells[row, 7].GetValue<string>(),
                            dob = worksheet.Cells[row, 8].GetValue<string>(),
                            email = worksheet.Cells[row, 9].GetValue<string>(),
                        };
                    }
                }
            }

            return null;
        }

        private void SendEmail(StudentDetails student)
        {
            var fromAddress = new MailAddress("vadaanyatrial@gmail.com", "VadaanyaOrg");
            var toAddress = new MailAddress(student.email, student.studentName);
            const string fromPassword = "nfqqnwwveyaeebgy";
            const string subject = "Student Details Submission Confirmation";
            string body = $"Dear {student.studentName},\n\nYour details have been successfully stored.\n\nBest regards,\nVadaanyaTalentTest";

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new System.Net.NetworkCredential(fromAddress.Address, fromPassword)
            };

            using (var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = body
            })
            {
                smtp.Send(message);
            }
        }

       /* private void SendWhatsAppMessage(StudentDetails student)
        {
            const string accountSid = "";
            const string authToken = "";

            TwilioClient.Init(accountSid, authToken);

            var message = MessageResource.Create(
                body: $"Dear {student.Name},\n\nYour details have been successfully stored.\n\nBest regards,\nVadaanyaTalentTest",
                from: new PhoneNumber("whatsapp:+14155238886"),
                to: new PhoneNumber($"whatsapp:{student.PhoneNumber}")
            );

            _logger.LogInformation($"WhatsApp message sent to {student.PhoneNumber}: {message.Sid}");
        }*/
    }
}

