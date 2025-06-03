using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Data;
using System.Data.SqlClient;


namespace Provvigioni_Agenti.Controllers
{
    internal class Database
    {
        private static string connectionString = string.Empty;

        public Database()
        {

            //connectionString = @"Data Source=LAPTOP-ACER-I7\SQLEXPRESS;Initial Catalog=cn8_rp;Integrated Security=True;Connect Timeout=30;Encrypt=False;";
        }

        static void stringConnection()
        {
            string computerName = System.Environment.MachineName;
            string databaseType = "DATABASE_AZIENDA";

            IniFile myIni = new IniFile("../settings.ini"); // leggo il file ini

            connectionString = myIni.Read("StringaConnessione", databaseType);
        }



        public static  List<T> SELECT_GET_LIST<T>(string sql, string database = "")
        {
            List<T> list = new List<T>();

            if (connectionString == string.Empty)
            {
                Database.stringConnection();
            }


            SqlConnection sdwDBConnection = new SqlConnection(connectionString.ToString());

            sdwDBConnection.Open();

            sql = "SET DATEFORMAT ymd " + sql;

            SqlCommand cmd = new SqlCommand(sql, sdwDBConnection);

            cmd.CommandTimeout = 600;

            SqlDataReader  reader = cmd.ExecuteReader();

            list = GetList<T>(reader);

            sdwDBConnection.Close();



            return list;

        }

        public static List<T> GetList<T>(IDataReader reader)
        {
            List<T> list = new List<T>();

            while (reader.Read())
            {
                var type = typeof(T);
                T obj = (T)Activator.CreateInstance(type);
                foreach (var prop in type.GetProperties())
                {
                    var protoType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(reader[prop.Name].ToString().Trim(), protoType));
                }
                list.Add(obj);
            }

            return list;
        }

    }

}
