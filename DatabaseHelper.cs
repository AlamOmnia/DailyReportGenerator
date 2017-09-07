using System.Data;
using MySql.Data.MySqlClient;

namespace Demo_Excel_Export
{
    public class DatabaseHelper
    {
        public string ConnectionString { get; private set; }
        public MySqlConnection MySqlConnection { get; private set; }

        #region Singleton implementation
        private static DatabaseHelper databaseHelper;
        public static DatabaseHelper Instance
        {
            get
            {
                if (databaseHelper == null)
                    databaseHelper = new DatabaseHelper();
                return databaseHelper;
            }
        }
        #endregion

        private DatabaseHelper()
        {
            // Set the connection string from config file
            ConnectionString = @"Server=10.0.30.77;User Id=root;password=Takay1#$ane;Persist Security Info=True;default command timeout=3600;database=purple";
        }

        // Open database connection
        public void OpenConnection()
        {
            MySqlConnection = new MySqlConnection(ConnectionString);
            MySqlConnection.Open();
        }

        // Execute sql query from the parameter, 
        public int ExecuteNonQuery(string sql, int timeout = 0)
        {
            MySqlCommand cmd = MySqlConnection.CreateCommand();
            cmd.CommandText = sql;
            cmd.CommandTimeout = timeout;

            return cmd.ExecuteNonQuery();
        }

        // Execute SQL from the parameter and read from database
        public MySqlDataReader ExecuteReader(string sql, int timeout = 0)
        {
            MySqlCommand cmd = MySqlConnection.CreateCommand();
            cmd.CommandText = sql;
            cmd.CommandTimeout = timeout;
            //System.Console.WriteLine(timeout);
            return cmd.ExecuteReader();
        }
        /*
        // Execute SQL from the parameter and read from database
        public MySqlDataReader ExecuteReader(MySqlCommand cmd, int timeout = 0)
        {
            cmd.CommandTimeout = timeout;
            return cmd.ExecuteReader();
        }*/

        // Open database connection
        public void CloseConnection()
        {
            if (MySqlConnection.State == ConnectionState.Open)
                MySqlConnection.Close();
        }

        public MySqlCommand CreateCommand()
        {
            var command = MySqlConnection.CreateCommand();
            command.CommandType = CommandType.Text;
            return command;
        }
    }
}
