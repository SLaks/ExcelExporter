using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelExporter {
	static class DbExtensions {
		public static string GetSqlType(this Type type) {
			type = Nullable.GetUnderlyingType(type) ?? type;
			switch (type.Name) {
				case "String":
				case "Boolean":
					return "Text";
				case "DateTime":
					return "Date";
				case "Int32":
					return "Double";
				case "Decimal":
					return "Currency";
				default:
					return type.Name;
			}
		}


		///<summary>Creates a DbCommand.</summary>
		///<param name="connection">The connection to create the command for.</param>
		///<param name="sql">The SQL of the command.</param>
		public static IDbCommand CreateCommand(this IDbConnection connection, string sql) {
			if (connection == null) throw new ArgumentNullException("connection");

			var retVal = connection.CreateCommand();
			retVal.CommandText = sql;
			return retVal;
		}
		///<summary>Executes a SQL statement against a connection.</summary>
		///<param name="connection">The connection to the database.  The connection is not closed.</param>
		///<param name="sql">The SQL to execute.</param>
		///<returns>The number of rows affected by the statement.</returns>
		public static int ExecuteNonQuery(this IDbConnection connection, string sql) {
			using (var command = connection.CreateCommand(sql)) return command.ExecuteNonQuery();
		}
	}
}
