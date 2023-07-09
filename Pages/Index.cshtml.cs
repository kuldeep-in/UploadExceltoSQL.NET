using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace H2Input.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private string dbtable = "dbo.H2Input1";
        private string connectionString;
        public DataTable capexData { get; set; }
        public DataTable opexData { get; set; }

        public IndexModel(ILogger<IndexModel> logger, IConfiguration configuration)
        {
            _logger = logger;
            //connectionString = $"Server={server};Database={database};User Id={user};Password={password};";
            connectionString = configuration.GetConnectionString("DefaultConnection");
        }

        [BindProperty]
        public Microsoft.AspNetCore.Http.IFormFile Upload { get; set; }

        public string Message { get; set; }


        public IActionResult OnGet()
        {
            return Page();
        }

        //public void OnGet()
        //{
        //    Message = TempData["Message"] as string;
        //}

        public async Task<IActionResult> OnPostAsync()
        {
            if (Upload != null)
            {
                try
                {
                    string fileName = Path.GetFileName(Upload.FileName);
                    string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);
                    string targetSheetName = "Summary";
                    // Define a custom encoding provider that supports the required encoding
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                    using (var stream = new MemoryStream())
                    {
                        TempData["Message"] = "File Uploading ...";

                        Upload.CopyTo(stream);
                        stream.Position = 0;

                        // Create an ExcelDataReader for the uploaded file
                        using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding(1252) // Specify the required encoding here
                        }))
                        {
                            // Configure the reader to read the first sheet of the workbook
                            var excelDataSetConfiguration = new ExcelDataSetConfiguration
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                                {
                                    UseHeaderRow = true
                                }
                            };

                            // Load the Excel data into a DataSet
                            var dataSet = reader.AsDataSet(excelDataSetConfiguration);

                            // Access the first table (sheet) in the DataSet
                            var dataTable = dataSet.Tables[targetSheetName];
                            List<KeyValuePair<string, string>> headerList = new List<KeyValuePair<string, string>>();

                            if (dataTable != null)
                            {
                                using (SqlConnection connection = new SqlConnection(connectionString))
                                {
                                    await connection.OpenAsync();
                                    using (SqlCommand truncateCommand = new SqlCommand($"Truncate table {dbtable}", connection))
                                    {
                                        await truncateCommand.ExecuteNonQueryAsync();
                                    }

                                    string categoryName = string.Empty;
                                    opexData = new DataTable();
                                    opexData.Columns.Add("-");
                                    capexData = new DataTable();
                                    capexData.Columns.Add("-");

                                    var rowCount = dataTable.Rows.Count;
                                    var colCount = dataTable.Columns.Count;

                                    string parameterName = "";
                                    string parameterKey = "";
                                    string parameterValue = "";
                                    //float parameterValue = 0;
                                    int casecount = 0;

                                    int col = 0;

                                    for (int row = 0; row < rowCount; row++)
                                    {
                                        col = 0;
                                        casecount = 0;

                                        if (categoryName.ToUpper() == "CAPEX" && headerList.Count() == 0)
                                        {
                                            col++;
                                            while (col < colCount && !(dataTable.Rows[row][col] == DBNull.Value || string.IsNullOrEmpty(dataTable.Rows[row][col].ToString())))
                                            {
                                                if (col % 3 == 0)
                                                {
                                                    casecount++;
                                                }

                                                capexData.Columns.Add($"Case {casecount} - {dataTable.Rows[row][col].ToString()}");
                                                opexData.Columns.Add($"Case {casecount} - {dataTable.Rows[row][col].ToString()}");

                                                KeyValuePair<string, string> pair = new KeyValuePair<string, string>(dataTable.Rows[row][col].ToString(), $"Case {casecount}");
                                                headerList.Add(pair);
                                                col++;
                                            }
                                        }

                                        if (dataTable.Rows[row][0] == DBNull.Value || string.IsNullOrEmpty(dataTable.Rows[row][0].ToString()))
                                        {
                                            continue;
                                        }
                                        else if (dataTable.Rows[row][0].ToString().ToUpper() == "CAPEX" || dataTable.Rows[row][0].ToString().ToUpper() == "OPEX")
                                        {
                                            categoryName = dataTable.Rows[row][0].ToString();
                                            continue;
                                        }
                                        else
                                        {
                                            parameterName = dataTable.Rows[row][0].ToString();
                                            col++;
                                            var capexRow = capexData.NewRow();
                                            var opexRow = opexData.NewRow();
                                            capexRow[0] = parameterName;
                                            opexRow[0] = parameterName;

                                            for (int i = 0; i < headerList.Count; i++)
                                            {
                                                KeyValuePair<string, string> pair = headerList[i];
                                                parameterValue = dataTable.Rows[row][i + 1].ToString();
                                                capexRow[i + 1] = parameterValue;
                                                opexRow[i + 1] = parameterValue;

                                                string insertQuery = $"INSERT INTO {dbtable} (CategoryName, ParameterName, CaseName, ParameterKey, ParameterValue) " +
                                                $"VALUES ('{categoryName}', '{parameterName}', '{pair.Value}', '{pair.Key}', '{parameterValue}')";

                                                using (SqlCommand command = new SqlCommand(insertQuery, connection))
                                                {
                                                    await command.ExecuteNonQueryAsync();
                                                }
                                                col++;

                                            }
                                            if (categoryName == "CAPEX")
                                                capexData.Rows.Add(capexRow);
                                            if (categoryName == "OPEX")
                                                opexData.Rows.Add(opexRow);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    TempData["Message"] = "Excel file uploaded successfully.";
                }
                catch (Exception ex)
                {
                    TempData["Message"] = "An error occurred: " + ex.Message;
                }
            }
            else
            {
                TempData["Message"] = "Please upload a valid Excel file.";
            }

            //return RedirectToPage("Index");
            return Page();
        }
    }
}