using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelExporter.Exporters {
	class DataTableSheet : SheetBase<DataRow> {
		public DataTable Table { get; private set; }

		public DataTableSheet(string name, DataTable table)
			: base(name) {
			if (table == null) throw new ArgumentNullException("table");
			if (table.Columns.Count == 0)
				throw new ArgumentException("Table has no columns to export", "table");

			Table = table;
		}

		protected override IEnumerable<ColumnInfo> GetColumns() {
			return Table.Columns.Cast<DataColumn>().Select(c => new ColumnInfo(c.ColumnName, c.DataType));
		}

		protected override IEnumerable<DataRow> GetRows() {
			return Table.AsEnumerable();
		}

		protected override IEnumerable<object> GetValues(DataRow row) {
			return Table.Columns.Cast<DataColumn>().Select(c => row[c]);
		}
	}
}
