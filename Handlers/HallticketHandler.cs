using Microsoft.IdentityModel.Tokens;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Globalization;
using System.Net;
using VadaanyaTalentTest1.Controllers;



namespace VadaanyaTalentTest1.Handlers
{
    public class HallticketHandler
    {
        private ExcelDataHandler _excelDH;

        public HallticketHandler(string excelSheetPath)
        {
            _excelDH = new ExcelDataHandler(excelSheetPath);
        }

        // DSC-2024 Exam

        public List<Dictionary<string, string>> GetDSC2024ApplicantDetails()
        {
            return _excelDH.GetExcelData();
        }

        public Dictionary<string, string> FilterDSC2024ApplicantDetails(long applicationNumber, long aadhaarNumber, string dob)
        {
            Dictionary<string, string> filterCriteria = new Dictionary<string, string>();

            if (dob.IsNullOrEmpty() || (applicationNumber == 0 && aadhaarNumber == 0))
                throw new StatusCodeException(HttpStatusCode.BadRequest, "Input data is invalid. Please try again.");

            if (applicationNumber != 0 && aadhaarNumber == 0)
                filterCriteria.Add("APPLICATION NUMBER", applicationNumber.ToString());            
            else if(applicationNumber == 0 && aadhaarNumber != 0)
                filterCriteria.Add("AADHAR NUMBER", aadhaarNumber.ToString());

            DateTime date = DateTime.ParseExact(dob, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            string formattedDob = date.ToString("M/d/yyyy");
            filterCriteria.Add("D-O-B", formattedDob);

            List<Dictionary<string, string>> filteredRows = _excelDH.FilterRowsByCriteria(filterCriteria);

            string tryAgainWithInput = applicationNumber == 0 ? "Application number" : "Aadhaar number";

            if (filteredRows.Count == 0)
                throw new StatusCodeException(HttpStatusCode.NotFound, $"No applicant found with the input details. Please try again. You may also try with {tryAgainWithInput}");

            if (filteredRows.Count > 1)
                throw new StatusCodeException(HttpStatusCode.Ambiguous, $"More than one applicant found with the same input details. Please try again. You may also try with {tryAgainWithInput}");

            Dictionary<string, string> applicant = filteredRows[0];

            return applicant;
        }


        public void GenerateDSC2024Hallticket(Dictionary<string,string> applicant, string inPath, string outPath)
        {
            if (!File.Exists(inPath))
                throw new StatusCodeException(HttpStatusCode.InternalServerError, "Cannot find hallticket template pdf.");

            PdfDocument document = PdfReader.Open(inPath, PdfDocumentOpenMode.Modify);
            PdfPage page = document.Pages[0];
            XGraphics gfx = XGraphics.FromPdfPage(page);

            XFont font1 = new XFont("Arial", 12, XFontStyleEx.Bold);
            XFont font2 = new XFont("Arial", 7.5, XFontStyleEx.Bold);
            long x = 190;

            gfx.DrawString(applicant["APPLICATION NUMBER"], 
                font1, 
                XBrushes.Black, 
                new XRect(x, 215, page.Width, page.Height), 
                XStringFormats.TopLeft);

            gfx.DrawString(!applicant["HALLTICKET NUMBER"].IsNullOrEmpty() ? applicant["HALLTICKET NUMBER"] : "", // hallticket
                font1,
                XBrushes.Black,
                new XRect(x, 252, page.Width, page.Height),
                XStringFormats.TopLeft);

            gfx.DrawString(applicant["NAME OF THE STUDENT"],
               font1,
               XBrushes.Black,
               new XRect(x, 285, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["FATHER NAME"],
               font1,
               XBrushes.Black,
               new XRect(x, 318, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["GENDER"],
               font1,
               XBrushes.Black,
               new XRect(x, 353, page.Width, page.Height),
               XStringFormats.TopLeft);

            DateTime date = DateTime.ParseExact(applicant["D-O-B"], "M/d/yyyy", CultureInfo.InvariantCulture);
            string formattedDate = date.ToString("dd-MM-yyyy");

            gfx.DrawString(formattedDate,
               font1,
               XBrushes.Black,
               new XRect(x, 385, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["MOBILE NUMBER"],
               font1,
               XBrushes.Black,
               new XRect(x, 420, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["DISTRICT"],
               font1,
               XBrushes.Black,
               new XRect(x, 455, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["MANDAL"],
               font1,
               XBrushes.Black,
               new XRect(x, 490, page.Width, page.Height),
               XStringFormats.TopLeft);

            string venue = !applicant["EXAM VENUE"].IsNullOrEmpty() ? applicant["EXAM VENUE"] : "";

            if (!venue.IsNullOrEmpty())
            {
                string venueLine1 = venue.Split("KOTHACHERUVU ,")[0] + "KOTHACHERUVU ,";
                string venueLine2 = venue.Split("KOTHACHERUVU ,")[1];

                gfx.DrawString(venueLine1,
                    font1,
                    XBrushes.Black,
                    new XRect(113, 543, page.Width, page.Height),
                    XStringFormats.TopLeft);

                gfx.DrawString(venueLine2,
                       font1,
                       XBrushes.Black,
                       new XRect(113, 563, page.Width, page.Height),
                       XStringFormats.TopLeft);
            }
            else
            {
                gfx.DrawString(venue,
                    font1,
                    XBrushes.Black,
                    new XRect(113, 543, page.Width, page.Height),
                    XStringFormats.TopLeft);
            }

            document.Save(outPath);
        }



        // --- NEXT EXAM ---
        // Create same methods here for next exam

    }
}
