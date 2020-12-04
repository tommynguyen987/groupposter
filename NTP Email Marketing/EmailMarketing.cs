using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Data.SQLite;
using System.Text;

namespace NTPEmailMarketing
{
    public class EmailMarketing
    {
        public const string TableName = "tb_EmailMarketing";
        public long ID { get; set; }
        public string Email { get; set; }
        public string Title { get; set; }
        public DateTime CreatedDate { get; set; }
        //public int TimesPosted { get; set; }
        public bool Rejected { get; set; }      
        
        public EmailMarketing Get(SQLiteDataReader reader)
        {
            this.ID = (long)reader["ID"];
            this.Email = reader["Email"].ToString();
            this.Title = reader["Title"].ToString();
            this.CreatedDate = reader["CreatedDate"]==null?DateTime.Now: DateTime.Parse(reader["CreatedDate"].ToString());
            this.Rejected = reader["Rejected"] == null ? false : bool.Parse(reader["Rejected"].ToString());                          
            return this;
        }
        public IList<EmailMarketing> GetList(long Id)
        {            
            return Get("select * from " + TableName + " where Id="+Id);
        }
        public EmailMarketing GetById(long Id)
        {
            return GetList(Id).FirstOrDefault();
        }
        public IList<EmailMarketing> Get(string SQL)
        {
            IList<EmailMarketing> list = new List<EmailMarketing>();
            using (DB db = new DB())
            {
                try
                {
                    using (var reader = db.GetReader(SQL))
                    {
                        if (reader.HasRows)
                            while (reader.Read())
                                list.Add(new EmailMarketing().Get(reader));
                    }
                }
                catch
                { }
            }
            return list;
        }
        public static bool CreateTable(DB db)
        {
            bool ret = true;
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("create table if not exists {0}(Id INTEGER PRIMARY KEY AUTOINCREMENT, Email TEXT NULL, Title TEXT NULL, CreatedDate TEXT NULL, Rejected INTEGER NULL); ", TableName);
            ret = db.UpdateTry(sb.ToString());
            return ret;
        }
        public static bool IsExistName(string email, string title)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("SELECT COUNT(*) FROM {0} WHERE Email=@Email AND Title=@Title;", TableName);
            long ret = 0;
            using (DB db = new DB())
            {
                CreateTable(db);
                SQLiteParameter[] parameters = new SQLiteParameter[2];
                parameters[0] = new SQLiteParameter("@Email", email);
                parameters[1] = new SQLiteParameter("@Title", title);
                ret = (long)db.ExecuteScalarTry(sb.ToString(), parameters);
            }
            return ret > 0;
        }
        public static bool IsExistTitle(string title, string createdDate)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("SELECT COUNT(*) FROM {0} WHERE Title=@Title AND CreatedDate=@CreatedDate;", TableName);
            long ret = 0;
            using (DB db = new DB())
            {
                CreateTable(db);
                SQLiteParameter[] parameters = new SQLiteParameter[2];
                parameters[0] = new SQLiteParameter("@Title", title);
                parameters[1] = new SQLiteParameter("@CreatedDate", createdDate);
                ret = (long)db.ExecuteScalarTry(sb.ToString(), parameters);
            }
            return ret > 0;
        }
        public static bool IsExistName(string title)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("SELECT COUNT(*) FROM {0} WHERE Title=@Title;", TableName);
            long ret = 0;
            using (DB db = new DB())
            {
                CreateTable(db);
                SQLiteParameter[] parameters = new SQLiteParameter[1];
                parameters[0] = new SQLiteParameter("@Title", title);
                ret = (long)db.ExecuteScalarTry(sb.ToString(), parameters);
            }
            return ret > 0;
        }
        public static bool IsRejected(string email)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("SELECT COUNT(*) FROM {0} WHERE Email=@Email AND Rejected=1;", TableName);
            long ret = 0;            
            using (DB db = new DB())
            {
                CreateTable(db);
                SQLiteParameter[] parameters = new SQLiteParameter[1];
                parameters[0] = new SQLiteParameter("@Email", email);
                ret = (long)db.ExecuteScalarTry(sb.ToString(), parameters);
            }
            return ret > 0;
        }        
        public static bool Insert(string email, string title, string createdDate, int rejected)
        {
            bool ret = false;
            using (DB db = new DB())
            {
                CreateTable(db);       
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("INSERT INTO {0}(Email, Title, CreatedDate, Rejected) VALUES (@Email, @Title, @CreatedDate, @Rejected);", TableName);
                SQLiteParameter[] parameters = new SQLiteParameter[4];
                parameters[0] = new SQLiteParameter("@Email", email);
                parameters[1] = new SQLiteParameter("@Title", title);
                parameters[2] = new SQLiteParameter("@CreatedDate", createdDate);
                parameters[3] = new SQLiteParameter("@Rejected", rejected);
                ret = db.UpdateTry(sb.ToString(), parameters);
            }                        
            return ret;
        }
        public static bool Update(string email, string title, string createdDate, int rejected, long Id)
        {
            bool ret = false;
            using (DB db = new DB())
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("UPDATE {0} SET Email=@Email, Title=@Title, CreatedDate=@CreatedDate, Rejected=@Rejected WHERE Id=@Id", TableName);
                SQLiteParameter[] parameters = new SQLiteParameter[5];
                parameters[0] = new SQLiteParameter("@Email", email);
                parameters[1] = new SQLiteParameter("@Title", title);
                parameters[2] = new SQLiteParameter("@CreatedDate", createdDate);
                parameters[3] = new SQLiteParameter("@Rejected", rejected);
                parameters[4] = new SQLiteParameter("@Id", Id);
                ret = db.UpdateTry(sb.ToString(), parameters);
            }
            return ret;
        }
        public static bool Update(string email, int rejected)
        {
            bool ret = false;
            using (DB db = new DB())
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("UPDATE {0} SET Rejected=@Rejected WHERE Email=@Email", TableName);
                SQLiteParameter[] parameters = new SQLiteParameter[2];
                parameters[0] = new SQLiteParameter("@Rejected", rejected);
                parameters[1] = new SQLiteParameter("@Email", email);
                ret = db.UpdateTry(sb.ToString(), parameters);
            }
            return ret;
        }
        public static bool Delete(long Id)
        {
            long ret = -1;
            using (DB db = new DB())
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("delete from {0} WHERE Id=@Id", TableName);
                SQLiteParameter[] parameters = new SQLiteParameter[1];
                parameters[0] = new SQLiteParameter("@Id", Id);
                ret = db.ExecuteTry(sb.ToString(), parameters);
            }
            return ret > -1;
        }
    }
}
