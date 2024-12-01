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
using Npgsql;
using System.Data;
using System.Windows.Input;
using VadaanyaTalentTest1.Models;

namespace VadaanyaTalentTest1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class StudentDetailsController : ControllerBase
    {
        private readonly ILogger<StudentDetailsController> _logger;
        private readonly IConfiguration _configuration;
        private readonly static List<StudentDetails> _studentDetails = new List<StudentDetails>();

        public StudentDetailsController(ILogger<StudentDetailsController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        [HttpGet]
        public ActionResult<StudentDetails> Get([FromQuery] long aadhaarNumber, [FromQuery] string mobileNumber, [FromQuery] string email)
        {
            var student = GetStudentFromDatabase(aadhaarNumber, mobileNumber, email);
            if (student == null)
            {
                return NotFound();
            }
            return Ok(student);


        }

        [HttpPost]
        public ActionResult<StudentDetails> Post(StudentDetails student)
        {
            int studentCount = GetStudentCountFromDatabase();
            if (studentCount >= Convert.ToInt32(Environment.GetEnvironmentVariable("STUDENT_LIMIT")))
            {
                return StatusCode(403, "Student limit reached. Cannot add more students.");
            }

            if (String.IsNullOrEmpty(student.testScore)) student.testScore = "0";

            // Check if the student already exists in the database
            var existingStudent = GetStudentFromDatabase(student.aadhaarNumber, student.mobileNumber, student.email);
            if (existingStudent != null)
            {
                return Conflict("Student with the same Aadhaar number, mobile number, or email already exists.");
            }

            student.applicationNumber = 2912240000 + studentCount + 1;
            _studentDetails.Add(student);
            AddStudentToExcel(student);
            AddStudentToDatabase(student);

            // Run the email sending task in the background
            Task.Run(async () =>
            {
                try
                {
                    await SendEmailAsync(student);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred while sending the email: {ex.Message}");
                }
            });

            //SendEmail(student);
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
                    worksheet.Cells[1, 10].Value = "TetScore";
                    worksheet.Cells[1, 11].Value = "Caste";
                    worksheet.Cells[1, 12].Value = "ApplicationNumber";
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
                worksheet.Cells[row, 10].Value = student.testScore;
                worksheet.Cells[row, 11].Value = student.caste;
                worksheet.Cells[row, 12].Value = student.applicationNumber;

                package.Save();
            }
        }

        private void SendEmail(StudentDetails student)
        {
            var fromAddress = new MailAddress(Environment.GetEnvironmentVariable("EMAIL_ID"), "Vadaanya Janaa Society");
            var toAddress = new MailAddress(student.email, student.studentName);
            var fromPassword = Environment.GetEnvironmentVariable("EMAIL_PASSWORD");
            string subject = $"Student Details Submission Confirmation-{student.applicationNumber}";
            string body = GetEmailbody(student);
            //string body = $"Dear { student.studentName},\n\nThank you for registering for Vadaanya’s DSC Talent Test 2024.\n Your application number is {student.applicationNumber}.\n\n******This is an auto-generated mail. Kindly do not reply to this email.*******\nKindly reach out to vadaanyadsc2024@gmail.com for any concerns.\n\nBest regards,\nTeam Vadaanya";

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
                Body = body,
                IsBodyHtml = true // Enable HTML formatting
            })
            {
                smtp.Send(message); ;
            }
        }

        private async Task SendEmailAsync(StudentDetails student)
        {
            var fromAddress = new MailAddress(Environment.GetEnvironmentVariable("EMAIL_ID"), "Vadaanya Janaa Society");
            var toAddress = new MailAddress(student.email, student.studentName);
            var fromPassword = Environment.GetEnvironmentVariable("EMAIL_PASSWORD");
            string subject = $"Student Details Submission Confirmation-{student.applicationNumber}";
            string body = GetEmailbody(student);

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
                Body = body,
                IsBodyHtml = true // Enable HTML formatting
            })
            {
                await smtp.SendMailAsync(message);
            }
        }

        private void AddStudentToDatabase(StudentDetails student)
        {
            //var connectionString = _configuration["ConnectionStrings:DefaultConnection"];
            var connectionString = Environment.GetEnvironmentVariable("CONNECTION_STRING");

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();

                using (var command = new NpgsqlCommand("INSERT INTO StudentDetails (aadhaarNumber, studentName, fatherName, gender, mobileNumber, district, mandal, dob, email,testscore,caste,applicationNumber) VALUES (@aadhaarNumber, @studentName, @fatherName, @gender, @mobileNumber, @district, @mandal, @dob, @email,@testscore,@caste,@applicationNumber)", connection))
                {
                    command.Parameters.AddWithValue("aadhaarNumber", student.aadhaarNumber);
                    command.Parameters.AddWithValue("studentName", student.studentName);
                    command.Parameters.AddWithValue("fatherName", student.fatherName);
                    command.Parameters.AddWithValue("gender", student.gender);
                    command.Parameters.AddWithValue("mobileNumber", student.mobileNumber);
                    command.Parameters.AddWithValue("district", student.district);
                    command.Parameters.AddWithValue("mandal", student.mandal);
                    command.Parameters.AddWithValue("dob", student.dob);
                    command.Parameters.AddWithValue("email", student.email);
                    command.Parameters.AddWithValue("testscore", student.testScore);
                    command.Parameters.AddWithValue("caste", student.caste);
                    command.Parameters.AddWithValue("applicationNumber", student.applicationNumber);


                    command.ExecuteNonQuery();
                }
            }
        }

        private StudentDetails GetStudentFromDatabase(long aadhaarNumber, string mobileNumber, string email)
        {
            //var connectionString = _configuration.GetConnectionString("DefaultConnection");
            var connectionString = Environment.GetEnvironmentVariable("CONNECTION_STRING");

            string dbcommand;
            string initialdbcommand = "SELECT aadhaarNumber, studentName, fatherName, gender, mobileNumber, district, mandal, dob, email, testscore, caste,applicationNumber FROM StudentDetails WHERE ";
            if (aadhaarNumber != 0 && String.IsNullOrEmpty(mobileNumber) && String.IsNullOrEmpty(email))
            {
                dbcommand = initialdbcommand + "aadhaarNumber = @aadhaarNumber";
            }
            else if (!String.IsNullOrEmpty(mobileNumber) && String.IsNullOrEmpty(email))
            {
                dbcommand = initialdbcommand + "mobileNumber = @mobileNumber";
            }
            else if (String.IsNullOrEmpty(mobileNumber) && !String.IsNullOrEmpty(email))
            {
                dbcommand = initialdbcommand + "email = @email";
            }
            else
            {
                dbcommand = initialdbcommand + "aadhaarNumber = @aadhaarNumber OR mobileNumber = @mobileNumber OR email = @Email";
            }

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();

                using (var command = new NpgsqlCommand(dbcommand, connection))
                {
                    if (aadhaarNumber != 0 && String.IsNullOrEmpty(mobileNumber) && String.IsNullOrEmpty(email))
                    {
                        command.Parameters.AddWithValue("aadhaarNumber", aadhaarNumber);
                    }
                    else if (!String.IsNullOrEmpty(mobileNumber) && String.IsNullOrEmpty(email))
                    {
                        command.Parameters.AddWithValue("mobileNumber", mobileNumber);
                    }
                    else if (String.IsNullOrEmpty(mobileNumber) && !String.IsNullOrEmpty(email))
                    {
                        command.Parameters.AddWithValue("email", email);
                    }
                    else
                    {
                        command.Parameters.AddWithValue("aadhaarNumber", aadhaarNumber);
                        command.Parameters.AddWithValue("mobileNumber", mobileNumber);
                        command.Parameters.AddWithValue("email", email);
                    }

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new StudentDetails
                            {
                                aadhaarNumber = reader.GetInt64(0),
                                studentName = reader.GetString(1),
                                fatherName = reader.GetString(2),
                                gender = reader.GetString(3),
                                mobileNumber = reader.GetString(4),
                                district = reader.GetString(5),
                                mandal = reader.GetString(6),
                                dob = reader.GetString(7),
                                email = reader.GetString(8),
                                testScore = reader.IsDBNull(9) ? null : reader.GetString(9),
                                caste = reader.GetString(10),
                                applicationNumber = reader.GetInt64(11),
                            };
                        }
                    }
                }
            }

            return null;
        }

        private int GetStudentCountFromDatabase()
        {
            //var connectionString = _configuration.GetConnectionString("DefaultConnection");
            var connectionString = Environment.GetEnvironmentVariable("CONNECTION_STRING");
            int count = 0;

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();

                using (var command = new NpgsqlCommand("SELECT COUNT(*) FROM StudentDetails", connection))
                {
                    count = Convert.ToInt32(command.ExecuteScalar());
                }
            }

            return count;
        }

        private string GetEmailbody(StudentDetails student)
        {
            var body = $@"
<html>
<head>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f4f4f4;
        }}
        .container {{
            width: 80%;
            margin: auto;
            background-color: #ffffff;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }}
        .header {{
            text-align: center;
            margin-bottom: 20px;
	    background-color: #36622b;
        }}
        .header img {{
            width: 50%;
            height: auto;
            background-color: #36622b;
            display: flex;
        }}
        .content p {{
            font-size: 12px;
            line-height: 1.6;
            margin:8px 0 8px 0;
        }}
        table {{
            width: 30%;
            border-collapse: collapse;
            margin-top: 1px;
        }}
        table, th, td {{
            border: 1px solid black;
        }}
        th, td {{
            padding: 3px;
            text-align: left;
            font-size: 11px;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        .footer {{
            margin-top: 20px;
            font-size: 11px;
            color: #555;
        }}
        .footer a {{
            color: #36622b;
            text-decoration: none;
        }}
    </style>
</head>
<body>
    <div class=""container"">
        <div class=""header"">
            <img src=""https://vadaanya.org/wp-content/uploads/2019/06/Logo.png"" alt=""Vadaanya Logo"">
        </div>
        <div class=""content"">
            <p>Dear {student.studentName},</p>
            <p>Thank you for registering for Vadaanya’s DSC Talent Test 2024
            <br>Your application number is {student.applicationNumber}.</p>
            <p>Your registration details are as follows:</p>
            <table>
                <tr>
                    <td>Aadhaar Number:</td>
                    <td>{student.aadhaarNumber}</td>
                </tr>
                <tr>
                    <td>Student Name:</td>
                    <td>{student.studentName}</td>
                </tr>
                <tr>
                    <td>Father Name:</td>
                    <td>{student.fatherName}</td>
                </tr>
                <tr>
                    <td>Gender:</td>
                    <td>{student.gender}</td>
                </tr>
                <tr>
                    <td>Mobile Number:</td>
                    <td>{student.mobileNumber}</td>
                </tr>
                <tr>
                    <td>Category:</td>
                    <td>{student.caste}</td>
                </tr>
                <tr>
                    <td>District:</td>
                    <td>{student.district}</td>
                </tr>
                <tr>
                    <td>Mandal:</td>
                    <td>{student.mandal}</td>
                </tr>
                <tr>
                    <td>DOB:</td>
                    <td>{student.dob}</td>
                </tr>
                <tr>
                    <td>Email ID:</td>
                    <td>{student.email}</td>
                </tr>
            </table>
            <p><b>******This is an auto-generated mail. Kindly do not reply to this email.*******</b>
            <br>For any inquiries or concerns,please contact <a href=""mailto:vadaanyadsc2024@gmail.com"">vadaanyadsc2024@gmail.com</a></p>
        </div>
        <div class=""footer"">
            <p>Best regards,<br>Team Vadaanya</p>
            <!--<p><a href='http://vadaanya.org/'>Visit our official website for details regarding exam and syllabus</a></p>-->
        </div>
    </div>
</body>
</html>";
            return body;
        }
    }
}

