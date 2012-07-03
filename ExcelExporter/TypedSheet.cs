using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace ExcelExporter {
	class TypedSheet<TRow> : Collection<TRow>, IExcelSheet {
		class Property {
			public string Name { get; private set; }
			public Func<TRow, object> GetValue { get; private set; }
			public Type Type;

			public Property(PropertyInfo prop) {
				Type = prop.PropertyType;
				Name = prop.Name.Replace('_', ' ');

				if (!prop.PropertyType.IsValueType)
					GetValue = (Func<TRow, object>)Delegate.CreateDelegate(typeof(Func<TRow, object>), prop.GetGetMethod());
				else {
					//If the property returns a value type, we need to compile an expression that boxes it.
					var param = Expression.Parameter(rowType, "row");
					GetValue = Expression.Lambda<Func<TRow, object>>(
						Expression.Convert(
							Expression.Property(param, prop),
							typeof(object)
						),
						param
					).Compile();
				}
			}
		}
		static readonly Type rowType = typeof(TRow);
		static readonly bool isAnonymousType = rowType.Namespace == null
											&& !rowType.IsPublic
											&& rowType.IsGenericType
											&& rowType.Name.Contains("AnonymousType");

		//If T is anonymous, get the declaration order from the ctor params.
		//http://msmvps.com/blogs/jon_skeet/archive/2009/12/09/quot-magic-quot-null-argument-testing.aspx
		static readonly List<Property> properties = (
			isAnonymousType ?
				rowType.GetConstructors()[0].GetParameters().Select(p => rowType.GetProperty(p.Name))
			  : rowType.GetProperties()
		).Select(p => new Property(p)).ToList();

		public string Name { get; set; }
		public TypedSheet(string name, IEnumerable<TRow> items) {
			if (properties.Count == 0)
				throw new InvalidOperationException("Type " + rowType + " has no properties to export");

			Name = name;
			((List<TRow>)base.Items).AddRange(items);
		}

		public void Export(IDbConnection connection) {
			var safeName = Name.Replace("]", "]]");

			StringBuilder tableBuilder = new StringBuilder();
			StringBuilder insertBuilder = new StringBuilder();

			int colIndex = 0;
			foreach (var prop in properties) {
				var safeColumnName = prop.Name.Replace("]", "]]");

				if (tableBuilder.Length == 0) {
					tableBuilder.AppendFormat("CREATE TABLE [{0}] (\r\n\t", safeName);
					insertBuilder.AppendFormat("INSERT INTO [{0}] VALUES (\r\n\t", safeName);
				} else {
					tableBuilder.Append(",\r\n\t");
					insertBuilder.Append(",\r\n\t");
				}
				tableBuilder.AppendFormat("[{0}] {1} NULL", safeColumnName, prop.Type.GetSqlType());
				insertBuilder.AppendFormat("@Col{0}", colIndex++);
			}
			tableBuilder.Append("\r\n)");
			insertBuilder.Append("\r\n)");

			connection.ExecuteNonQuery(tableBuilder.ToString());

			foreach (var row in this) {
				using (var command = connection.CreateCommand(insertBuilder.ToString())) {
					for (int c = 0; c < properties.Count; c++)
						command.Parameters.Add(new OleDbParameter("Col" + c, properties[c].GetValue(row) ?? DBNull.Value));

					command.ExecuteNonQuery();
				}
			}
		}
	}
}
