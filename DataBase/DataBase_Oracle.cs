using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.Linq;
using System.Xml;

//using System.Data.OracleClient;

using Oracle.ManagedDataAccess.Client;

namespace DataBase.Connections1
{
    public class DBConnection
    {

        public string DB = "ORACLE";

        public string ConnectionString = "";

        public OracleConnection Connection = null;
        public OracleTransaction Transaction = null;


        public Dictionary<string, string> userInfo = null;

        private DateTime DT_START;

        private string FolderName = "";
        private string Instanace_ID = "";
        private string UserName = "";

        public DBConnection()
        {
            FolderName = getSettings("LogFolder");
        }

        public string getConnectionstring()
        {
            return System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            //return System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString1"].ConnectionString;
        }
        public OracleConnection OpenConnection()
        {
            if (Connection == null)
                Connection = new OracleConnection(getConnectionstring());
            if (Connection.State != ConnectionState.Open)
                Connection.Open();
            return Connection;
        }
        public void CloseConnection()
        {
            if (Connection != null)
            {
                if (Connection.State == ConnectionState.Open)
                    Connection.Close();
            }
        }
        public void BeginTransaction()
        {
            //Connection = new ORacleConnection(getConnectionstring());
            //Connection.Open();
            OpenConnection();
            Transaction = Connection.BeginTransaction();
        }
        public void CommitTransaction()
        {
            if (Transaction != null)
                Transaction.Commit();
            Connection.Close();
        }
        public void RollbackTransaction()
        {
            if (Transaction != null)
                Transaction.Rollback();
            if (Connection != null)
            {
                if (Connection.State == ConnectionState.Open)
                    Connection.Close();
            }
        }

        // Start CRUD Functions
        public object ExecuteScalar(string sql)
        {
            DT_START = DateTime.Now;
            OpenConnection();
            OracleCommand Cmd = Connection.CreateCommand();
            Cmd.CommandText = sql;
            Cmd.CommandType = CommandType.Text;
            if (Transaction != null)
                Cmd.Transaction = Transaction;
            object mData = Cmd.ExecuteScalar();
            CreateLog(sql);
            //return Cmd.ExecuteScalar();
            return mData;
        }
        public int ExecuteNonQuery(string sql)
        {
            DT_START = DateTime.Now;
            OpenConnection();
            OracleCommand Cmd = Connection.CreateCommand();
            if (Transaction != null)
                Cmd.Transaction = Transaction;
            Cmd.CommandText = sql;
            Cmd.CommandType = CommandType.Text;
            int mData = Cmd.ExecuteNonQuery();
            CreateLog(sql);
            return mData;
        }
        public DataTable ExecuteQuery(string sql)
        {

            DT_START = DateTime.Now;
            OpenConnection();
            DataTable DataTable = new DataTable();
            OracleCommand Cmd = new OracleCommand(sql, Connection);
            if (Transaction != null)
                Cmd.Transaction = Transaction;
            OracleDataAdapter dataAdapter = new OracleDataAdapter();
            dataAdapter.SelectCommand = Cmd;
            dataAdapter.Fill(DataTable);
            CreateLog(sql);
            return DataTable;
        }

        

        private void CreateLog(string str)
        {
            try
            {
                TaskLog(str);
                /*
                if (userInfo != null && FolderName.Length > 0)
                {
                    Instanace_ID = userInfo["INSTANCE_ID"].ToString();
                    UserName = userInfo["USR_NAME"].ToString();
                    if (Instanace_ID.Length > 0)
                    {
                        TaskLog(str);
                    }
                }
                */
            }
            catch (Exception) { }
        }


        public void TaskLog(string str)
        {
            DateTime Dt = DateTime.Now;
            int Seconds = Dt.Subtract(DT_START).Seconds;
            string totSeconds = Dt.Subtract(DT_START).TotalSeconds.ToString("0.##");
            string FileName = FolderName + Dt.ToString("yyyy-MM-dd") + ".txt";
            if ( Seconds >=1)
                FileName = FolderName + Dt.ToString("yyyy-MM-dd") + "-long.txt";
            string sData = Dt.ToString("yyyy-MM-dd:HH:mm:ss tt") + "," + Seconds.ToString() + "," + totSeconds +  "," + "\"" + str + "\"";
            StreamWriter sw = new StreamWriter(FileName, true);
            sw.WriteLine(sData);
            sw.Flush();
            sw.Close();
        }
        public void old_TaskLog(string str)
        {
            DateTime Dt = DateTime.Now;
            int Seconds = Dt.Subtract(DT_START).Seconds;
            string totSeconds = Dt.Subtract(DT_START).TotalSeconds.ToString("#.##");
            string FileName = FolderName + Dt.ToString("yyyy-MM-dd") + "-" + Instanace_ID + ".txt";
            string sData = Dt.ToString("yyyy-MM-dd:HH:mm:ss tt") + "," + Seconds.ToString() + "," + totSeconds + "," + "\"" + UserName + "\"" + "," + "\"" + str + "\"";
            StreamWriter sw = new StreamWriter(FileName, true);
            sw.WriteLine(sData);
            sw.Flush();
            sw.Close();
        }

        public void CreateErrorLog(string str)
        { 
            //string FileName = FolderName +  "errorlog.txt";

            string FileName = @"c:\\log\errorlog.txt";

            StreamWriter sw = new StreamWriter(FileName, true);
            sw.WriteLine(str);
            sw.Flush();
            sw.Close();
        }


        public string getSettings(string str)
        {
            string sRetVal = "";
            try
            {
                sRetVal = System.Configuration.ConfigurationManager.AppSettings[str].ToString();
            }
            catch (Exception)
            {
                sRetVal = "";
            }
            return sRetVal;
        }


        // End CRUD Functions

        public XmlDocument ExecuteXmlQuery(string sql)
        {
            OpenConnection();
            DataSet Ds_Test = new DataSet();
            OracleCommand Cmd = new OracleCommand(sql, Connection);
            if (Transaction != null)
                Cmd.Transaction = Transaction;
            OracleDataAdapter dataAdapter = new OracleDataAdapter();
            dataAdapter.SelectCommand = Cmd;
            dataAdapter.Fill(Ds_Test);
            XmlDocument objXMLDoc = new XmlDocument();
            objXMLDoc.LoadXml(Ds_Test.GetXml());
            return objXMLDoc;
        }
        public string TestConnection()
        {
            string sql = "select 'Hello' as code ";
            OpenConnection();
            OracleCommand Cmd = new OracleCommand(sql, Connection);
            return Cmd.ExecuteScalar().ToString();
        }

        public DataContext CreateDBContext()
        {
            DataContext DB = new DataContext(getConnectionstring());
            return DB;
        }

        public Boolean IsRowExists(string sql)
        {
            object oData = ExecuteScalar(sql);
            if (oData == null)
                return false;
            else  if ( oData == DBNull.Value)
                return false;
            else
                return true;
        }

        public int LastNo(string TableName, string Branch_Code, string Year)
        {
            string sql = "";
            int iNextNo = 0;
            if (TableName == "CARGO_IMP_MASTERM")
                sql = "select max(mbl_cfno) as cfno from cargo_imp_masterm ";
            if (sql != "")
            {
                object oData = ExecuteScalar(sql);
                if (oData.Equals(DBNull.Value))
                    iNextNo = 1000;
                else
                    iNextNo = int.Parse(oData.ToString());
            }
            return iNextNo;
        }


        public Boolean IsValidUser()
        {
            string sql = "";
            Boolean bRet = false;

            string usr_code = userInfo["USR_CODE"].ToString();
            string macaddress = userInfo["MACADDRESS"].ToString();
            string LOGIN_MULTIPLE_SYSETM = userInfo["ALLOW_LOGIN_FROM_MULTIPLE_SYSTEM"].ToString();

            //string LOGIN_MULTIPLE_SYSETM = getSettings("ALLOW_LOGIN_FROM_MULTIPLE_SYSTEM");



            if (LOGIN_MULTIPLE_SYSETM == "Y" || LOGIN_MULTIPLE_SYSETM == "")
                return true;

            sql = "select usr_code from user_userm  where usr_code = '" + usr_code + "'  and usr_macaddress = '" + macaddress + "'";
            object oData = ExecuteScalar(sql);
            if (oData == null)
                bRet = false;
            else
                bRet = true;
            return bRet;
        }


    }
}
