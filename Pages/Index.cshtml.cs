using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using H2Input.Models;
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
		private string dbtable = "dbo.KD01";
		private string connectionString;
		public DataTable cat01Data { get; set; }
		public DataTable cat02Data { get; set; }

		public DataTable sqlDataTable { get; set; }
		public DataTable viewDataTable { get; set; }

		public IndexModel(ILogger<IndexModel> logger, IConfiguration configuration)
		{
			_logger = logger;
			connectionString = configuration.GetConnectionString("DefaultConnection");
		}

		[BindProperty]
		public Microsoft.AspNetCore.Http.IFormFile Upload { get; set; }

		public string Message { get; set; }


		public async Task<IActionResult> OnGet()
		{
			HashSet<string> columnList = new HashSet<string>();
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();


				string sql = $"SELECT * FROM {dbtable}";
				SqlCommand command = new SqlCommand(sql, connection);

				SqlDataReader reader = command.ExecuteReader();

				sqlDataTable = new DataTable("UploadedData");
				sqlDataTable.Columns.Add("CategoryName", typeof(string));
				sqlDataTable.Columns.Add("ParameterName", typeof(string));
				sqlDataTable.Columns.Add("CaseName", typeof(string));
				sqlDataTable.Columns.Add("ParameterKey", typeof(string));
				sqlDataTable.Columns.Add("ParameterValue", typeof(string));

				while (reader.Read())
				{
					DataRow row = sqlDataTable.NewRow();
					row["CategoryName"] = reader["CategoryName"];
					row["ParameterName"] = reader["ParameterName"];
					row["CaseName"] = reader["CaseName"];
					row["ParameterKey"] = reader["ParameterKey"];
					row["ParameterValue"] = reader["ParameterValue"];

					if (reader["CaseName"] is null || string.IsNullOrEmpty(reader["CaseName"].ToString()))
					{
						columnList.Add($"{reader["ParameterKey"]}");
					}
					else
					{
						columnList.Add($"{reader["CaseName"]}-{reader["ParameterKey"]}");
					}

					sqlDataTable.Rows.Add(row);

				}
			}

			viewDataTable = new DataTable();
			viewDataTable.Columns.Add("Catg");
			viewDataTable.Columns.Add("-");

			foreach (var column in columnList)
			{
				viewDataTable.Columns.Add(column);
			}

			var viewDataRow = viewDataTable.NewRow();

			int colCount = 0;
			foreach (DataRow rowItem in sqlDataTable.Rows)
			{
				//if (rowItem["-"].ToString() == "Category01")
				//{
				if (colCount == 0)
				{ 
				viewDataRow[colCount] = rowItem["CategoryName"];
				colCount++;
				viewDataRow[colCount] = rowItem["ParameterName"];
				colCount++;
				viewDataRow[colCount] = rowItem["ParameterValue"];
				colCount++;
				}
				else if (colCount < columnList.Count() + 2) 
				{
					viewDataRow[colCount] = rowItem["ParameterValue"];
					colCount++;
				}
				else
				{
					viewDataTable.Rows.Add(viewDataRow);
					viewDataRow = viewDataTable.NewRow();
					colCount = 0;
				}

			}
			viewDataTable.Rows.Add(viewDataRow);

			return Page();
		}

		public async Task<List<KD01Model>> GetAllData()
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				await connection.OpenAsync();

				string sql = $"SELECT * FROM {dbtable}";
				SqlCommand command = new SqlCommand(sql, connection);

				List<KD01Model> rows = new List<KD01Model>();
				SqlDataReader reader = await command.ExecuteReaderAsync();

				while (reader.Read())
				{
					KD01Model row = new KD01Model
					{
						ID = (int)reader["ID"],
						CategoryName = (string)reader["CategoryName"],
						ParameterName = (string)reader["ParameterName"],
						CaseName = (string)reader["CaseName"],
						ParameterKey = (string)reader["ParameterKey"],
						ParameterValue = (string)reader["ParameterValue"]
					};

					rows.Add(row);
				}

				return rows;
			}
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
									cat02Data = new DataTable();
									cat02Data.Columns.Add("-");
									cat01Data = new DataTable();
									cat01Data.Columns.Add("-");

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

										if (categoryName == "Category01" && headerList.Count() == 0)
										{
											col++;
											while (col < colCount && !(dataTable.Rows[row][col] == DBNull.Value || string.IsNullOrEmpty(dataTable.Rows[row][col].ToString())))
											{
												if (col % 3 == 0)
												{
													casecount++;
												}

												if (casecount < 1)
												{
													cat01Data.Columns.Add($"{dataTable.Rows[row][col].ToString()}");
													cat02Data.Columns.Add($"{dataTable.Rows[row][col].ToString()}");
													KeyValuePair<string, string> pair = new KeyValuePair<string, string>(dataTable.Rows[row][col].ToString(), $"");
													headerList.Add(pair);
												}
												else
												{
													cat01Data.Columns.Add($"Case {casecount} - {dataTable.Rows[row][col].ToString()}");
													cat02Data.Columns.Add($"Case {casecount} - {dataTable.Rows[row][col].ToString()}");
													KeyValuePair<string, string> pair = new KeyValuePair<string, string>(dataTable.Rows[row][col].ToString(), $"Case {casecount}");
													headerList.Add(pair);
												}



												col++;
											}
										}

										if (dataTable.Rows[row][0] == DBNull.Value || string.IsNullOrEmpty(dataTable.Rows[row][0].ToString()))
										{
											continue;
										}
										else if (dataTable.Rows[row][0].ToString() == "Category01" || dataTable.Rows[row][0].ToString() == "Category02")
										{
											categoryName = dataTable.Rows[row][0].ToString();
											continue;
										}
										else
										{
											parameterName = dataTable.Rows[row][0].ToString();
											col++;
											var capexRow = cat01Data.NewRow();
											var opexRow = cat02Data.NewRow();
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
											if (categoryName == "Category01")
												cat01Data.Rows.Add(capexRow);
											if (categoryName == "Category02")
												cat02Data.Rows.Add(opexRow);
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

namespace H2Input.Models
{
	public class KD01Model
	{
		public int ID { get; set; }
		public string CategoryName { get; set; }
		public string ParameterName { get; set; }
		public string CaseName { get; set; }
		public string ParameterKey { get; set; }
		public string ParameterValue { get; set; }
	}
}