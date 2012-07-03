using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace ExcelExporter.Exporters {
	class TypedSheet<TRow> : SheetBase<TRow>, IExcelSheet {
		class Property : ColumnInfo {
			public Func<TRow, object> GetValue { get; private set; }

			public Property(PropertyInfo prop)
				: base(prop.Name.Replace('_', ' '), prop.PropertyType) {
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

		readonly IList<TRow> items;
		public TypedSheet(string name, IEnumerable<TRow> items)
			: base(name) {
			if (properties.Count == 0)
				throw new InvalidOperationException("Type " + rowType + " has no properties to export");

			this.items = items.ToList();
		}

		protected override IEnumerable<ColumnInfo> GetColumns() {
			return properties.Cast<ColumnInfo>();	//Since I want to support .Net 3.5, IEnumerable<T> is not covariant
		}

		protected override IEnumerable<TRow> GetRows() {
			return items;
		}

		protected override IEnumerable<object> GetValues(TRow row) {
			return properties.Select(p => p.GetValue(row));
		}
	}
}
