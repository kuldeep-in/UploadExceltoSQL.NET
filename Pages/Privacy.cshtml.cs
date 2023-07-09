using DocumentFormat.OpenXml.Office2013.Excel;
using H2Input.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data;
using System.Data.SqlClient;

namespace H2Input.Pages
{
	public class PrivacyModel : PageModel
	{
		private readonly ILogger<PrivacyModel> _logger;
		private string dbtable = "dbo.KD01";
		private string connectionString;
		public DataTable cat01Data { get; set; }
		public DataTable cat02Data { get; set; }

		public DataTable sqlDataTable { get; set; }
		public DataTable viewDataTable { get; set; }
		public List<KD01Model> kD01Models { get; set; }

		public PrivacyModel(ILogger<PrivacyModel> logger, IConfiguration configuration)
		{
			_logger = logger;
			connectionString = configuration.GetConnectionString("DefaultConnection");
		}

        public async Task<IActionResult> OnGet()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync();

                string sql = $"SELECT * FROM {dbtable}";
                SqlCommand command = new SqlCommand(sql, connection);

                kD01Models = new List<KD01Model>();
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

                    kD01Models.Add(row);
                }

                //return rows;
            }

            return Page();
        }

		public async void GetAllData()
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				await connection.OpenAsync();

				string sql = $"SELECT * FROM {dbtable}";
				SqlCommand command = new SqlCommand(sql, connection);

                kD01Models = new List<KD01Model>();
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

					kD01Models.Add(row);
				}

				//return rows;
			}
		}
	}
}