using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;

namespace ExcelExporter.Exporters {
	class DataReaderSheet : SheetBase<IDataRecord> {
		readonly IDataReader reader;
		public DataReaderSheet(string name, IDataReader reader)
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
			return new DbEnumerable(reader).Cast<IDataRecord>();
		}
		class DbEnumerable : IEnumerable {
			readonly IDataReader reader;
			public DbEnumerable(IDataReader reader) { this.reader = reader; }

			public IEnumerator GetEnumerator() { return new DbEnumerator(reader, closeReader: true); }
		}

		protected override IEnumerable<object> GetValues(IDataRecord row) {
			return Enumerable.Range(0, reader.FieldCount)
							 .Select(row.GetValue);
		}
	}
}
