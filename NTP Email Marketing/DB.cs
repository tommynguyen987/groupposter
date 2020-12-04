using System;
using System.Data;
using System.Data.SQLite;
using System.Linq;

namespace NTPEmailMarketing
{
    public class DB : IDisposable
    {
        //public IDbConnection Connection =>
             //new SQLiteConnection(ConnectionString, true).OpenAndReturn();

        public DB()
        {
            CreateTry();
        }

        private string _database = null;
        private System.Data.SQLite.SQLiteConnection cnn = null;

        public string Database
        {
            get
            {
                //AppDomain.CurrentDomain.BaseDirectory                
                //Environment.CurrentDirectory
                //System.IO.Path.GetDirectory(Application.ExecutablePath)
                if (_database == null) _database = System.IO.Directory.GetCurrentDirectory() + "\\data.sqlite";
                return _database;
            }
            set
            {
                _database = value;
            }
        }
        // Get or set the connection string to the SQLite database
        private string ConnectionString
        {
            get
            {
                return "Data Source=" + Database + ";Version=3;";
            }
        }
        public bool Create()
        {
            if (System.IO.File.Exists(Database)) return false;
            System.Data.SQLite.SQLiteConnection.CreateFile(Database);
            return true;
        }
        public bool CreateTry()
        {
            try
            {
                return Create();
            }
            catch
            {
                return false;
            }
        }
        public System.Data.DataTable LoadTry(string SQL)
        {
            try
            {
                return Load(SQL);
            }
            catch
            {
                return null;
            }
        }
        public System.Data.DataTable Load(string SQL)
        {
            DataSet ds = new DataSet();
            using (System.Data.SQLite.SQLiteConnection connection = new System.Data.SQLite.SQLiteConnection(ConnectionString, true))
            {
                connection.Open();
                using (System.Data.SQLite.SQLiteCommand command = new System.Data.SQLite.SQLiteCommand(SQL, connection))
                {
                    using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(command))
                    {
                        da.Fill(ds);
                    }
                }
                connection.Close();
            }
            if (ds.Tables.Count > 0) return ds.Tables[0];
            return null;
        }
        public System.Data.SQLite.SQLiteDataReader GetReader(string SQL)
        {
            System.Data.SQLite.SQLiteConnection connection = new System.Data.SQLite.SQLiteConnection(ConnectionString, true);
            connection.Open();

            System.Data.SQLite.SQLiteCommand command = new System.Data.SQLite.SQLiteCommand(SQL, connection);            
            return command.ExecuteReader();
        }
        public long ExecuteTry(string SQL, SQLiteParameter[] parameters)
        {
            try
            {
                return Execute(SQL, parameters);
            }
            catch
            {
                return -1;
            }
        }
        public long Execute(string SQL, SQLiteParameter[] parameters)
        {
            System.Data.SQLite.SQLiteConnection connection = null;
            long ret = -1;
            if (cnn != null)
            {
                connection = cnn;
            }
            else
            {
                connection = new System.Data.SQLite.SQLiteConnection(ConnectionString, true);
                connection.Open();
            }

            using (System.Data.SQLite.SQLiteCommand command = new System.Data.SQLite.SQLiteCommand(SQL, connection))
            {
                if (parameters != null)
                {
                    command.Parameters.AddRange(parameters);
                }
                ret = (long)command.ExecuteNonQuery();
            }

            if (cnn != null)
            {
                connection = null;
            }
            else
            {
                connection.Close();
                connection.Dispose();
                connection = null;
            }
            return ret;
        }        
        public object ExecuteScalarTry(string SQL, SQLiteParameter[] parameters)
        {
            try
            {
                return ExecuteScalar(SQL, parameters);
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public object ExecuteScalar(string SQL, SQLiteParameter[] parameters)
        {
            System.Data.SQLite.SQLiteConnection connection = null;
            object ret = null;
            if (cnn != null)
            {
                connection = cnn;
            }
            else
            {
                connection = new System.Data.SQLite.SQLiteConnection(ConnectionString, true);
                connection.Open();
            }

            using (System.Data.SQLite.SQLiteCommand command = new System.Data.SQLite.SQLiteCommand(SQL, connection))
            {
                if (parameters != null)
                {
                    command.Parameters.AddRange(parameters);
                }                
                ret = command.ExecuteScalar();
            }

            if (cnn != null)
            {
                connection = null;
            }
            else
            {
                connection.Close();
                connection.Dispose();
                connection = null;
            }
            return ret;
        }
        public bool Update(string SQL)
        {
            System.Data.SQLite.SQLiteConnection connection = null;

            if (cnn != null)
            {
                connection = cnn;
            }
            else
            {
                connection = new System.Data.SQLite.SQLiteConnection(ConnectionString, true);
                connection.Open();
            }

            using (System.Data.SQLite.SQLiteCommand command = new System.Data.SQLite.SQLiteCommand(SQL, connection))
            {
                command.ExecuteNonQuery();
            }

            if (cnn != null)
            {
                connection = null;
            }
            else
            {
                connection.Close();
                connection.Dispose();
                connection = null;
            }
            return true;
        }
        public bool Update(string SQL, SQLiteParameter[] parameters)
        {
            System.Data.SQLite.SQLiteConnection connection = null;

            if (cnn != null)
            {
                connection = cnn;
            }
            else
            {
                connection = new System.Data.SQLite.SQLiteConnection(ConnectionString, true);
                connection.Open();
            }

            using (System.Data.SQLite.SQLiteCommand command = new System.Data.SQLite.SQLiteCommand(SQL, connection))
            {
                if (parameters != null)
                {
                    command.Parameters.AddRange(parameters);
                }
                command.ExecuteNonQuery();
            }

            if (cnn != null)
            {
                connection = null;
            }
            else
            {
                connection.Close();
                connection.Dispose();
                connection = null;
            }
            return true;
        }        
        public bool UpdateTry(string SQL)
        {
            try
            {
                return Update(SQL);
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public bool UpdateTry(string SQL, SQLiteParameter[] parameters)
        {
            try
            {
                return Update(SQL, parameters);
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public bool isTableExists(string tablename)
        {
            Int32 count = 0;
            string SQL = "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name = '" + tablename + "'";
            try
            {
                using (System.Data.SQLite.SQLiteConnection connection = new System.Data.SQLite.SQLiteConnection(ConnectionString, true))
                {
                    connection.Open();
                    using (System.Data.SQLite.SQLiteCommand command = new System.Data.SQLite.SQLiteCommand(SQL, connection))
                    {
                        using (IDataReader reader = command.ExecuteReader())
                        {
                            if ((reader != null) && (reader.Read())) count = reader.GetInt32(0);
                        }
                    }
                    connection.Close();
                }
                return count > 0;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public string SafeValue(string value)
        {
            if (value == null) return string.Empty;
            return value.Replace('"', '\'');
        }
        public void Dispose()
        {
            if (cnn != null) cnn.Close();
            if (cnn != null) cnn.Dispose();
            cnn = null;
        }
    }
}
