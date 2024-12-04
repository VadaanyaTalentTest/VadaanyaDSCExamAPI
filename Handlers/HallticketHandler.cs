using Microsoft.IdentityModel.Tokens;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Globalization;



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
                throw new Exception("Input data is invalid.");

            if (applicationNumber != 0 && aadhaarNumber == 0)
                filterCriteria.Add("ApplicationNumber", applicationNumber.ToString());            
            else if(applicationNumber == 0 && aadhaarNumber != 0)
                filterCriteria.Add("AadhaarNumber", aadhaarNumber.ToString());
            
            filterCriteria.Add("Dob", dob);

            List<Dictionary<string, string>> filteredRows = _excelDH.FilterRowsByCriteria(filterCriteria);

            if (filteredRows.Count == 0)
                throw new Exception("No applicant found with the input details.");

            if (filteredRows.Count > 1)
                throw new Exception("More than one applicant found with the same input details.");

            Dictionary<string, string> applicant = filteredRows[0];

            return applicant;
        }


        public void GenerateDSC2024Hallticket(Dictionary<string,string> applicant, string inPath, string outPath)
        {
            if (!File.Exists(inPath))
                throw new Exception("Cannot find hallticket template pdf.");

            PdfDocument document = PdfReader.Open(inPath, PdfDocumentOpenMode.Modify);
            PdfPage page = document.Pages[0];
            XGraphics gfx = XGraphics.FromPdfPage(page);

            XFont font1 = new XFont("Arial", 12, XFontStyleEx.Bold);
            XFont font2 = new XFont("Arial", 7.5, XFontStyleEx.Bold);
            long x = 190;

            gfx.DrawString(applicant["ApplicationNumber"], 
                font1, 
                XBrushes.Black, 
                new XRect(x, 215, page.Width, page.Height), 
                XStringFormats.TopLeft);

            gfx.DrawString(!applicant["HallTicketNumber"].IsNullOrEmpty() ? applicant["HallTicketNumber"] : "TBD", // hallticket
                font1,
                XBrushes.Black,
                new XRect(x, 252, page.Width, page.Height),
                XStringFormats.TopLeft);

            gfx.DrawString(applicant["StudentName"],
               font1,
               XBrushes.Black,
               new XRect(x, 285, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["FatherName"],
               font1,
               XBrushes.Black,
               new XRect(x, 318, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["Gender"],
               font1,
               XBrushes.Black,
               new XRect(x, 353, page.Width, page.Height),
               XStringFormats.TopLeft);

            DateTime date = DateTime.ParseExact(applicant["Dob"], "yyyy-MM-dd", CultureInfo.InvariantCulture);
            string formattedDate = date.ToString("dd-MM-yyyy");

            gfx.DrawString(formattedDate,
               font1,
               XBrushes.Black,
               new XRect(x, 385, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["MobileNumber"],
               font1,
               XBrushes.Black,
               new XRect(x, 420, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["District"],
               font1,
               XBrushes.Black,
               new XRect(x, 455, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(applicant["Mandal"],
               font1,
               XBrushes.Black,
               new XRect(x, 490, page.Width, page.Height),
               XStringFormats.TopLeft);

            gfx.DrawString(!applicant["Venue"].IsNullOrEmpty() ? applicant["Venue"] : "TBD", 
               font1,
               XBrushes.Black,
               new XRect(113, 543, page.Width, page.Height),
               XStringFormats.TopLeft);

            document.Save(outPath);
        }



        // --- NEXT EXAM ---
        // Create same methods here for next exam

    }
}
