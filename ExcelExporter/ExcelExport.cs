using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace ExcelExporter {
	///<summary>Exports a collection of tables to an Excel spreadsheet.</summary>
	public class ExcelExport {
		///<summary>Creates a new ExcelExport instance.</summary>
		public ExcelExport() { Sheets = new Collection<IExcelSheet>(); }

		///<summary>Gets the sheets which will be exported by this instance.</summary>
		public Collection<IExcelSheet> Sheets { get; private set; }

		///<summary>Adds a collection of strongly-typed objects to be exported.</summary>
		///<param sheetName="sheetName">The name of the sheet to generate.</param>
		///<param sheetName="items">The rows to export to the sheet.</param>
		///<returns>This instance, to allow chaining.kds</returns>
		public ExcelExport AddSheet<TRow>(string sheetName, IEnumerable<TRow> items) {
			Sheets.Add(new Exporters.TypedSheet<TRow>(sheetName, items));
			return this;
		}

		///<summary>Adds the contents of a DataTable instance to be exported, using the table's name as the worksheet name.</summary>
		public ExcelExport AddSheet(DataTable table) {
			if (table == null) throw new ArgumentNullException("table");
			return AddSheet(table.TableName, table);
		}
		///<summary>Adds the contents of a DataTable instance to be exported.</summary>
		public ExcelExport AddSheet(string sheetName, DataTable table) {
			Sheets.Add(new Exporters.DataTableSheet(sheetName, table));
			return this;
		}

		///<summary>Adds the contents returned by an open DataReader to be exported.</summary>
		///<param sheetName="sheetName">The name of the sheet to generate.</param>
		///<pparam name="reader">The reader to read rows from.  This reader must remain open when ExportTo() is called.</pparam>
		public ExcelExport AddSheet(string sheetName, DbDataReader reader) {
			Sheets.Add(new Exporters.DataReaderSheet(sheetName, reader));
			return this;
		}

		///<summary>Exports all of the added sheets to an Excel file.</summary>
		///<param sheetName="fileName">The filename to export to.  The file type is inferred from the extension.</param>
		public void ExportTo(string fileName) {
			ExportTo(fileName, GetDBType(Path.GetExtension(fileName)));
		}
		///<summary>Exports all of the added sheets to an Excel file.</summary>
		public void ExportTo(string fileName, ExcelFormat format) {
			using (var connection = new OleDbConnection(GetConnectionString(fileName, format))) {
				connection.Open();
				foreach (var sheet in Sheets) {
					sheet.Export(connection);
				}
			}
		}

		#region Excel Formats
		static readonly List<KeyValuePair<ExcelFormat, string>> FormatExtensions = new List<KeyValuePair<ExcelFormat, string>> {
			new KeyValuePair<ExcelFormat, string>(ExcelFormat.Excel2003,			".xls"),
			new KeyValuePair<ExcelFormat, string>(ExcelFormat.Excel2007,			".xlsx"),
			new KeyValuePair<ExcelFormat, string>(ExcelFormat.Excel2007Binary,		".xlsb"),
			new KeyValuePair<ExcelFormat, string>(ExcelFormat.Excel2007Macro,		".xlsm"),
		};

		///<summary>Gets the database format that uses the given extension.</summary>
		public static ExcelFormat GetDBType(string extension) {
			var pair = FormatExtensions.FirstOrDefault(kvp => kvp.Value.Equals(extension, StringComparison.OrdinalIgnoreCase));

			if (pair.Value == null)
				throw new ArgumentException("Unrecognized extension: " + extension, "extension");
			return pair.Key;
		}

		///<summary>Gets the file extension for a database format.</summary>
		public static string GetExtension(ExcelFormat format) { return FormatExtensions.First(kvp => kvp.Key == format).Value; }

		static string GetConnectionString(string filePath, ExcelFormat format) {
			if (String.IsNullOrEmpty(filePath)) throw new ArgumentNullException("filePath");

			var csBuilder = new OleDbConnectionStringBuilder { DataSource = filePath, PersistSecurityInfo = false };

			const string ExcelProperties = "IMEX=0;HDR=YES";
			switch (format) {
				case ExcelFormat.Excel2003:
					csBuilder.Provider = "Microsoft.Jet.OLEDB.4.0";
					csBuilder["Extended Properties"] = "Excel 8.0;" + ExcelProperties;
					break;
				case ExcelFormat.Excel2007:
					csBuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
					csBuilder["Extended Properties"] = "Excel 12.0 Xml;" + ExcelProperties;
					break;
				case ExcelFormat.Excel2007Binary:
					csBuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
					csBuilder["Extended Properties"] = "Excel 12.0;" + ExcelProperties;
					break;
				case ExcelFormat.Excel2007Macro:
					csBuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
					csBuilder["Extended Properties"] = "Excel 12.0 Macro;" + ExcelProperties;
					break;
			}

			return csBuilder.ToString();
		}
		#endregion
	}

	///<summary>A format for a database file.</summary>
	public enum ExcelFormat {
		///<summary>An Excel 97-2003 .xls file.</summary>
		Excel2003,
		///<summary>An Excel 2007 .xlsx file.</summary>
		Excel2007,
		///<summary>An Excel 2007 .xlsb binary file.</summary>
		Excel2007Binary,
		///<summary>An Excel 2007 .xlsm file with macros.</summary>
		Excel2007Macro
	}

	///<summary>Stores a single exportable worksheet.</summary>
	public interface IExcelSheet {
		///<summary>Gets the name of the worksheet.</summary>
		string Name { get; }
		///<summary>Exports the worksheet to the specified OleDb connection.</summary>
		void Export(OleDbConnection connection);
	}
}