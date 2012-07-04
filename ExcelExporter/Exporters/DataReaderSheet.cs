using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;

namespace ExcelExporter.Exporters {
	class DataReaderSheet : SheetBase<IDataRecord> {
		readonly DbDataReader reader;
		public DataReaderSheet(string name, DbDataReader reader)
			: base(name) {
			this.reader = reader;
		}

		protected override IEnumerable<ColumnInfo> GetColumns() {
			return Enumerable.Range(0, reader.FieldCount)
							 .Select(i => new ColumnInfo(
								 reader.GetName(i),
								 reader.GetFieldType(i)
							 ));
		}

		protected override IEnumerable<IDataRecord> GetRows() {
			return reader.Cast<IDataRecord>();
		}

		protected override IEnumerable<object> GetValues(IDataRecord row) {
			return Enumerable.Range(0, reader.FieldCount)
							 .Select(row.GetValue);
		}
	}
}
