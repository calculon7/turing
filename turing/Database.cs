using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Turing {
    public class Database {
        public static object[][] Read(string query) {
            string conString = "redacted_connection_string";

            OracleConnection con = new OracleConnection(conString);
            con.Open();

            OracleCommand cmd = con.CreateCommand();
            cmd.CommandText = query;

            OracleDataReader reader = cmd.ExecuteReader();

            var records = new List<object[]>();

            while (reader.Read()) {
                var values = new object[reader.FieldCount];
                reader.GetOracleValues(values);
                records.Add(values);
            }

            return records.ToArray();
        }

        public static int Write(string nonquery) {
            string conString = "redacted_connection_string";

            OracleConnection con = new OracleConnection(conString);
            con.Open();

            OracleCommand cmd = con.CreateCommand();
            cmd.CommandText = nonquery;

            int rowsAffected = cmd.ExecuteNonQuery();

            return rowsAffected;
        }
    }
}
