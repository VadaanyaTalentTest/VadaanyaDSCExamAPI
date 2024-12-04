using OfficeOpenXml;

namespace VadaanyaTalentTest1.Handlers
{
    public class ExcelDataHandler
    {
        private string _filePath;
        private List<Dictionary<string, string>> _excelData;

        public ExcelDataHandler(string filePath)
        {
            _filePath = filePath;
            _excelData = ReadExcelByColumnNames();
        }

        public List<Dictionary<string, string>> GetExcelData()
        {
            return _excelData;
        }

        public List<Dictionary<string, string>> FilterRowsByCriteria(Dictionary<string, string> filterCriteria)
        {
            List<Dictionary<string, string>> filteredRows = _excelData.Where(row => filterCriteria.All(criteria => row.ContainsKey(criteria.Key) && row[criteria.Key] == criteria.Value)).ToList();

            return filteredRows;
        }

        public List<Dictionary<string, string>> ReadExcelByColumnNames()
        {
            List<Dictionary<string, string>> result = new List<Dictionary<string, string>>();

            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var columnNames = new Dictionary<int, string>();

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    columnNames[col] = worksheet.Cells[1, col].Text;
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var rowData = new Dictionary<string, string>();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        rowData[columnNames[col]] = worksheet.Cells[row, col].Text;
                    }
                    result.Add(rowData);
                }
            }

            return result;
        }


    }
}
