using System.Configuration;
using System.Data.SqlClient;

namespace ExcelLoader.Console
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			using (var cn = GetConnection())
			{
				var loader = new AerieEngineering.ExcelLoader();
				loader.Save(@"C:\Users\Adam\Downloads\Copy of gvl_charlene_priorityone_12 04 2018.xlsx", cn, "dbo", "MaximoDataV2", true);
			}
		}

		private static string GetConnectionString()
		{
			return ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
		}

		private static SqlConnection GetConnection()
		{
			return new SqlConnection(GetConnectionString());
		}
	}
}