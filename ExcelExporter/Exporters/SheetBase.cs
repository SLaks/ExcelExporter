using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;

namespace ExcelExporter.Exporters {
	///<summary>A reusable base class that exports data provided by a derived class to an excel spreadsheet.</summary>
	///<typeparam name="TRow">The type of the objects used by the implementation to represent each row.</typeparam>
	public abstract class SheetBase<TRow> : IExcelSheet {
		///<summary>Gets the name of the exported worksheet.</summary>
		public string Name { get; private set; }

		///<summary>Creates a SheetBase instance.</summary>
		protected SheetBase(string name) {
			Name = name;
		}

		///<summary>Gets the columns to put in the sheet.</summary>
		protected abstract IEnumerable<ColumnInfo> GetColumns();
		///<summary>Gets a collection of objects containing the data to show in each row.</summary>
		protected abstract IEnumerable<TRow> GetRows();
		///<summary>Gets the column values for a specific row object.  (in the order returned by GetColumns())</summary>
		protected abstract IEnumerable<object> GetValues(TRow row);

		///<summary>Exports this sheet to an open OleDb connection.</summary>
		public void Export(OleDbConnection connection) {
			var safeName = Name.Replace("]", "]]");

			StringBuilder tableBuilder = new StringBuilder();
			StringBuilder insertBuilder = new StringBuilder();

			int colIndex = 0;
			foreach (var prop in GetColumns()) {
				var safeColumnName = prop.Name.Replace("]", "]]");

				if (tableBuilder.Length == 0) {
					tableBuilder.AppendFormat("CREATE TABLE [{0}] (\r\n\t", safeName);
					insertBuilder.AppendFormat("INSERT INTO [{0}] VALUES (\r\n\t", safeName);
				} else {
					tableBuilder.Append(",\r\n\t");
					insertBuilder.Append(",\r\n\t");
				}
				tableBuilder.AppendFormat("[{0}] {1} NULL", safeColumnName, prop.DataType.GetSqlType());
				insertBuilder.AppendFormat("@Col{0}", colIndex++);
			}
			tableBuilder.Append("\r\n)");
			insertBuilder.Append("\r\n)");

			connection.ExecuteNonQuery(tableBuilder.ToString());


			foreach (var row in GetRows()) {
				using (var command = connection.CreateCommand(insertBuilder.ToString())) {
					colIndex = 0;
					foreach (var value in GetValues(row))
						command.Parameters.Add(new OleDbParameter("Col" + colIndex++, value ?? DBNull.Value));

					command.ExecuteNonQuery();
				}
			}
		}
	}
	///<summary>Stores information about a column in an exported sheet.</summary>
	public class ColumnInfo {
		///<summary>Gets the name of the column as displayed in the column header.</summary>
		public string Name { get; private set; }
		///<summary>Gets the type of the column.</summary>
		public Type DataType { get; private set; }

		///<summary>Creates a new ColumnInfo instance.</summary>
		public ColumnInfo(string name, Type columnType) {
			Name = name;
			DataType = columnType;
		}
	}
}
