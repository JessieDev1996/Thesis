using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using MySql.Data.MySqlClient;
namespace Project_Test1.Models
{
    public class ConnectDB
    {
        public DataSet select(string sql,string nameTable)
        {
            //string connectionString = "Data Source=JES-42COM\\SQLEXPRESS;Initial Catalog=Cooperative;Integrated Security=True;";

            string connectionString = "Data Source=localhost;username=root;password=;database=coperativermutt;"; //Connect Database

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                DataSet ds = new DataSet();
                try
                {
                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    da.Fill(ds, nameTable);
                    conn.Close();
                    conn.Dispose();
                    return ds;
                }
                catch(Exception)
                {
                    conn.Close();
                    conn.Dispose();
                    //ex.Message.ToString();
                    return null;
                }
            }
        }
        public string insert_update_delete(string sql)
        {
            //string connectionString = "Data Source=JES-42COM\\SQLEXPRESS;Initial Catalog=Cooperative;Integrated Security=True;";

            string connectionString = "Data Source=localhost;username=root;password=;database=coperativermutt;"; //Connect Database

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    MySqlCommand _insert_update_delete = new MySqlCommand(sql, conn);
                    _insert_update_delete.ExecuteNonQuery();
                    conn.Close();
                    conn.Dispose();
                    return "Success";
                }
                catch(Exception ex)
                {
                    conn.Close();
                    conn.Dispose();
                    return ex.ToString();
                }
            }
        }
    }
}