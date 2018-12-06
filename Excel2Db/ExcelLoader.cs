using AerieEngineering.Util;
using ExcelDataReader;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace AerieEngineering
{
	public class ExcelLoader
	{
		public ExcelLoader()
		{
		}

		public int Save(string fileName, SqlConnection connection, string schemaName, string tableName, bool truncateFirst = false)
		{
			var ds = Read(fileName);
			SaveDataTable(connection, ds.Tables[0], schemaName, tableName, truncateFirst);
			return ds.Tables[0].Rows.Count;
		}

		public int Save(Stream stream, SqlConnection connection, string schemaName, string tableName, bool truncateFirst = false)
		{
			var ds = Read(stream);
			SaveDataTable(connection, ds.Tables[0], schemaName, tableName, truncateFirst);
			return ds.Tables[0].Rows.Count;
		}

		private void SaveDataTable(SqlConnection connection, DataTable table, string schemaName, string tableName, bool truncateFirst)
		{
			if (truncateFirst) connection.Execute($"TRUNCATE TABLE [{schemaName}].[{tableName}]");

			// thanks to https://stackoverflow.com/a/4582786/2023653
			foreach (DataRow row in table.Rows)
			{
				row.AcceptChanges();
				row.SetAdded();
			}

			using (SqlCommand select = BuildSelectCommand(table, connection, schemaName, tableName))
			{
				using (var adapter = new SqlDataAdapter(select))
				{
					using (var builder = new SqlCommandBuilder(adapter))
					{
						adapter.InsertCommand = builder.GetInsertCommand();
						adapter.Update(table);
					}
				}
			}
		}

		private SqlCommand BuildSelectCommand(DataTable table, SqlConnection connection, string schemaName, string tableName)
		{
			string[] columnNames = table.Columns.OfType<DataColumn>().Select(col => col.ColumnName).ToArray();
			string query = $"SELECT {string.Join(", ", columnNames.Select(col => $"[{col}]"))} FROM [{schemaName}].[{tableName}]";
			return new SqlCommand(query, connection);
		}

		public DataSet Read(string fileName)
		{
			using (var stream = File.OpenRead(fileName))
			{
				return Read(stream);
			}
		}

		public DataSet Read(Stream stream)
		{
			using (var reader = ExcelReaderFactory.CreateReader(stream))
			{
				return reader.AsDataSet(new ExcelDataSetConfiguration()
				{
					UseColumnDataType = true,
					ConfigureDataTable = (r) =>
					{
						return new ExcelDataTableConfiguration() { UseHeaderRow = true };
					}
				});
			}
		}
	}
}