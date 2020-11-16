using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace DataBase
{
    public class DBRecord
    {

        private string SQL1 = "";
        private string SQL2 = "";
        public  string WHERE = "";
        private string Table_Name = "";
        private string Mode = "";
        private string PK_COLUMN = "";
        private string PK_VALUE = "";

        public  void CreateRow(string lTable_Name, string lMode, string PKey_Col, string pKey_Data)
        {
            Init();
            Table_Name = lTable_Name;
            Mode = lMode;
            SetPrimaryKey(PKey_Col, pKey_Data);
        }
        private  void Init()
        {
            SQL1 = "";
            SQL2 = "";
            WHERE = "";
            Table_Name = "";
            Mode = "";
            PK_COLUMN = "";
            PK_VALUE = "";
        }
        public  string UpdateRow()
        {
            string str = "";
            if (Mode == "ADD")
            {
                str = "INSERT INTO {TNAME} ({COLUMNS}) VALUES ({VALUES})";
                str = str.Replace("{TNAME}", Table_Name);
                str = str.Replace("{COLUMNS}", SQL1);
                str = str.Replace("{VALUES}", SQL2);
            }
            if (Mode == "EDIT")
            {
                str = "UPDATE {TNAME} SET {VALUES} WHERE {WHERE} ";
                str = str.Replace("{TNAME}", Table_Name);
                str = str.Replace("{VALUES}", SQL1);
                str = str.Replace("{WHERE}", WHERE);
            }
            return str;
        }
        public  void SetPrimaryKey(string FldName, string Data)
        {
            PK_COLUMN = FldName;
            PK_VALUE = Data;
            if (Mode == "ADD")
                InsertData("STRING", FldName, Data);
            else
            {
                if (WHERE != "")
                    WHERE += " AND ";
                WHERE += PK_COLUMN += "='" + PK_VALUE + "'";
            }
        }

      


        public void InsertString(string FldName, object Data, string CharacterCase = "")
        {
            if (Data == null)
            {
                Data = "";
            }
            if (string.IsNullOrEmpty(Data.ToString()) || string.IsNullOrWhiteSpace(Data.ToString()))
            {
                Data = "";
            }
            else
            {
                if (CharacterCase != "P")
                {
                    if (CharacterCase == "T")
                        Data = Data.ToString().Replace("'", "''");
                    else
                        Data = Data.ToString().Trim().Replace("'", "''");
                }

                if (CharacterCase.Trim().ToUpper() == "" || CharacterCase.Trim().ToUpper() == "U")
                    Data = Data.ToString().ToUpper();
                else if (CharacterCase.Trim().ToUpper() == "L")
                    Data = Data.ToString().ToLower();
            }
            InsertData("STRING", FldName.ToString(), Data.ToString());
        }


        public void InsertNumerID(string FldName, Object Data)
        {
            if (Data == null || Data.ToString() == "" || Data.ToString() == "0")
                Data = "NULL";
            InsertData("NUMERIC", FldName, Data.ToString());
        }



        public  void InsertNumeric(string FldName, string Data)
        {
            if (Data.Trim() == "")
                Data = "0";
            InsertData("NUMERIC", FldName, Data);
        }

        public  void InsertFunction(string FldName, string Data)
        {
            InsertData("FUNCTION", FldName, Data);
        }

        public  void InsertDate(string FldName, Object Data)
        {
            string sData = "";
            DateTime Dt;
            if (Data == null || Data.ToString() == "")
                sData = "NULL";
            else
            {
               // Dt = (DateTime)Data;
                Dt = DateTime.Parse(Data.ToString());
                sData = "'{DATE}'";
                sData = sData.Replace("{DATE}", Dt.ToString(Lib.BACK_END_DATE_FORMAT));


            }
            InsertData("DATE", FldName, sData);
        }

        public void InsertDateString(string FldName, object Data)
        {
            string sData = "";
            
            if (Data == null || Data.ToString() == "")
                sData = "NULL";
            else
            {
                sData = "'" + Data + "'";
            }
            InsertData("DATE", FldName, sData);
        }



        public void InsertDateAndTime(string FldName, Object Data)
        {
            string sData = "";
            DateTime Dt;
            if (sData == null)
                sData = "NULL";
            else
            {
                Dt = (DateTime)Data;
                sData = "'{DATE}'";
                sData = sData.Replace("{DATE}", Dt.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            InsertData("DATE", FldName, sData);
        }
        private  void InsertData(string sType, string FldName, string Data)
        {
            if (Mode == "ADD")
            {
                if (SQL1 != "")
                    SQL1 += ",";
                SQL1 += FldName;
                if (SQL2 != "")
                    SQL2 += ",";
                if (sType == "STRING")
                {
                    if ( Data =="")
                        SQL2 += "NULL";
                    else 
                        SQL2 += "'" + Data + "'";
                }
                if (sType == "NUMERIC")
                    SQL2 += Data;
                if (sType == "FUNCTION")
                    SQL2 += "(" + Data + ")";
                if (sType == "DATE")
                    SQL2 += Data;
            }
            else
            {
                if (SQL1 != "")
                    SQL1 += ",";
                if (sType == "STRING")
                {
                    if (Data == "")
                    {
                        SQL1 += FldName + "=" + "NULL";
                    }
                    else
                        SQL1 += FldName + "=" + "'" + Data + "'";
                }  
                if (sType == "NUMERIC")
                    SQL1 += FldName + "=" + Data;
                if (sType == "FUNCTION")
                    SQL1 += FldName + "=" + "(" + Data + ")";
                if (sType == "DATE")
                    SQL1 += FldName + "=" + Data;
            }
        }
        public   void AddGeneralColumns(string DataMode, Dictionary<string, string> userInfo)
        {
            if (DataMode == "ADD")
            {
                InsertString("rec_locked", "N");
                InsertString("rec_created_by", userInfo["USR_CODE"].ToString());
                InsertDateAndTime("rec_created_date", System.DateTime.Now);
            }
            else
            {
                InsertString("rec_edited_by", userInfo["USR_CODE"].ToString());
                InsertDateAndTime("rec_edited_date", System.DateTime.Now);
            }
        }

        public void AddGeneralColumns(string DataMode, Dictionary<string, string> userInfo, DateTime mDate)
        {
            if (DataMode == "ADD")
            {
                InsertString("rec_locked", "N");
                InsertString("rec_created_id", userInfo["USR_PKID"].ToString());
                InsertString("rec_created_by", userInfo["USR_CODE"].ToString());
                InsertString("rec_company_code", userInfo["REC_COMPANY_CODE"].ToString());
                InsertString("rec_branch_code", userInfo["REC_BRANCH_CODE"].ToString());
                InsertDateAndTime("rec_created_date", mDate);
            }
            else
            {
                InsertString("rec_edited_by", userInfo["USR_CODE"].ToString());
                InsertDateAndTime("rec_edited_date", mDate);
            }
        }
        //END OF CODE//
    }


    public class DBFunctions
    {
        public Boolean LogAudit(Dictionary<string, string> LogData, Dictionary<string, string> userInfo, DateTime Date_Created, out string ErrorMessage)
        {
            Boolean bRet = false;
            ErrorMessage = "";

            DataBase.Connections.DBConnection DB = null;
            try
            {
                DB = new DataBase.Connections.DBConnection();
                DB.BeginTransaction();

                DBRecord mRec = new DBRecord();
                mRec.CreateRow("user_audit", "ADD", "AUDIT_ID", userInfo["AUDIT_ID"].ToString());
                mRec.InsertString("AUDIT_USER_ID", GetString( userInfo["USR_PKID"].ToString() , 40));
                mRec.InsertString("AUDIT_USER_CODE", GetString( userInfo["USR_CODE"].ToString(),15));
                mRec.InsertString("AUDIT_SCREEN",GetString(  LogData["AUDIT_SCREEN"].ToString(),60));
                mRec.InsertString("AUDIT_ACTION", GetString( LogData["AUDIT_ACTION"].ToString(),60));
                mRec.InsertString("AUDIT_REMARKS", GetString( LogData["AUDIT_REMARKS"].ToString(),250));
                mRec.InsertString("AUDIT_KEY", GetString( LogData["AUDIT_KEY"].ToString(),40));
                mRec.InsertString("AUDIT_REFNO", GetString( LogData["AUDIT_REFNO"].ToString(),60));
                mRec.InsertString("AUDIT_COMPANY_CODE", GetString( userInfo["REC_COMPANY_CODE"].ToString(),4));
                mRec.InsertString("AUDIT_BRANCH_CODE", GetString( userInfo["REC_BRANCH_CODE"].ToString(),4));
                mRec.InsertString("AUDIT_FYCODE", GetString( userInfo["YEAR_CODE"].ToString(),4));
                mRec.InsertDateAndTime("AUDIT_CLIENT_DATE", Date_Created);

                string Sql = mRec.UpdateRow();
                DB.ExecuteNonQuery(Sql);

                if (LogData["AUDIT_ACTION"].ToString() == "LOGOUT")
                {
                    Sql = " update user_userm set usr_logged = 'N' where usr_pkid = '" + userInfo["USR_PKID"].ToString() + "'";
                    DB.ExecuteNonQuery(Sql); 
                }

                DB.CommitTransaction();
                bRet = true;
            }
            catch (Exception Ex)
            {
                bRet = false;
                DB.RollbackTransaction();
                ErrorMessage = Ex.Message.ToString();
            }
            return bRet;
        }

        public Boolean UserLogBook(string Log_id, string Log_Source, string Log_Type, Object Log_Doc_Date, string Log_HBL_Id, string Log_CUST_Id, string Log_User_Code, DateTime Date_Created, string Log_From_Amt, string Log_To_Amt, string Log_Description, out string ErrorMessage)
        {
            Boolean bRet = false;
            ErrorMessage = "";
            DataBase.Connections.DBConnection DB = null;
            try
            {
                DB = new DataBase.Connections.DBConnection();
                DB.BeginTransaction();

                DBRecord mRec = new DBRecord();
                mRec.CreateRow("user_logbook", "ADD", "LOG_ID", Log_id);
                mRec.InsertString("LOG_SOURCE", GetString(Log_Source, 30));
                mRec.InsertString("LOG_TYPE", GetString(Log_Type, 10));
                mRec.InsertString("LOG_HBL_ID", Log_HBL_Id);
                mRec.InsertString("LOG_CUST_ID", Log_CUST_Id);
                mRec.InsertNumeric("LOG_FROM_AMT", Lib.Conv2Decimal(Log_From_Amt).ToString());
                mRec.InsertNumeric("LOG_TO_AMT", Lib.Conv2Decimal(Log_To_Amt).ToString());
                mRec.InsertString("LOG_USER_CODE", GetString(Log_User_Code, 15));
                mRec.InsertDate("LOG_DOC_DATE", Log_Doc_Date);
                mRec.InsertDateAndTime("LOG_DATE", Date_Created);

                string Sql = mRec.UpdateRow();
                DB.ExecuteNonQuery(Sql);

                DB.CommitTransaction();
                bRet = true;
            }
            catch (Exception Ex)
            {
                bRet = false;
                DB.RollbackTransaction();
                ErrorMessage = Ex.Message.ToString();
            }
            return bRet;
        }




        private string GetString(string str, int iLen)
        {
            if (str.Length > iLen)
                return str.Substring(0, iLen);
            else
                return str;

        }

        


        public Boolean IsShipmentClosed(string sID, string sTYPE, string BranchID)
        {
            Boolean bRet = false;
            int LockDays = 0;

            DateTime Dt_Now;
            DateTime REF_DATE = DateTime.Now;
            double Days = 0;

            string sql = "";
            string LOCK_STATUS = null;
            string OPR_MODE = "";
            DataBase.Connections.DBConnection DB = null;
            DataTable Dt_temp = new DataTable();
            DataTable Dt_param = new DataTable();

            try
            {
                DB = new DataBase.Connections.DBConnection();


                sql = "";
                if (sTYPE == "MASTER")
                    sql = "select mbl_pkid,mbl_mode,mbl_lock,mbl_unlock_date,mbl_ref_date from cargo_masterm where mbl_pkid = '" + sID + "'";
                else if (sTYPE == "HOUSE")
                    sql = "select mbl_pkid,mbl_mode,mbl_lock,mbl_unlock_date,mbl_ref_date from cargo_masterm a inner join cargo_housem b on a.mbl_pkid = b.hbl_mbl_id where b.hbl_pkid ='" + sID + "'";
                else if (sTYPE == "INVOICE")
                    sql = "select mbl_pkid,mbl_mode,mbl_lock,mbl_unlock_date,mbl_ref_date from cargo_masterm a inner join cargo_invoicem b on a.mbl_pkid = b.inv_mbl_id where b.inv_pkid ='" + sID + "'";

                if (sql != "")
                {
                    Dt_temp = DB.ExecuteQuery(sql);

                    if (Dt_temp.Rows.Count > 0)
                    {
                        LOCK_STATUS = Dt_temp.Rows[0]["mbl_lock"].ToString();
                        OPR_MODE = Dt_temp.Rows[0]["mbl_mode"].ToString();
                        REF_DATE = (DateTime)Dt_temp.Rows[0]["mbl_ref_date"];
                        if (OPR_MODE == "GE" || OPR_MODE == "PS" || OPR_MODE == "PR" || OPR_MODE == "CM")
                            OPR_MODE = "ADMIN";
                        if (OPR_MODE == "OTHERS" || OPR_MODE == "FA" || OPR_MODE == "EXTRA")
                            OPR_MODE = "OTHERS";

                        sql = "";
                        if (OPR_MODE == "SEA EXPORT" || OPR_MODE == "SEA IMPORT")
                            sql = "select param_name3 from mast_param where param_name1 = 'LOCK-DAYS-SEA'    and param_name2 = '" + BranchID + "'";
                        if (OPR_MODE == "AIR EXPORT" || OPR_MODE == "AIR IMPORT")
                            sql = "select param_name3 from mast_param where param_name1 = 'LOCK-DAYS-AIR'   and param_name2 = '" + BranchID + "'";
                        if (OPR_MODE == "OTHERS")
                            sql = "select param_name3 from mast_param where param_name1 = 'LOCK-DAYS-OTHERS' and param_name2 = '" + BranchID + "'";
                        if (OPR_MODE == "ADMIN")
                            sql = "select param_name3 from mast_param where param_name1 = 'LOCK-DAYS-ADMIN' and param_name2 = '" + BranchID + "'";
                        if (sql != "")
                        {
                            Dt_param = DB.ExecuteQuery(sql);
                            if (Dt_param.Rows.Count > 0)
                            {
                                LockDays = Lib.Conv2Integer(Dt_param.Rows[0]["param_name3"].ToString());
                            }
                        }

                        if (LOCK_STATUS == null || LOCK_STATUS.Trim() == "")
                        {
                            Dt_Now = DateTime.Now;
                            Days = Dt_Now.Subtract(REF_DATE).TotalDays;
                            if ((OPR_MODE == "SEA EXPORT" || OPR_MODE == "SEA IMPORT") && Days > LockDays && LockDays > 0)
                                bRet = true;
                            if ((OPR_MODE == "AIR EXPORT" || OPR_MODE == "AIR IMPORT") && Days > LockDays && LockDays > 0)
                                bRet = true;
                            if ((OPR_MODE == "OTHERS") && Days > LockDays && LockDays > 0)
                                bRet = true;
                            if ((OPR_MODE == "ADMIN") && Days > LockDays && LockDays > 0)
                                bRet = true;
                        }
                        else if (LOCK_STATUS == "L")
                            bRet = true;
                        else if (LOCK_STATUS == "U")
                        {
                            bRet = false;
                            if (Dt_temp.Rows[0]["mbl_unlock_date"] != DBNull.Value)
                            {
                                Dt_Now = DateTime.Now;
                                Days = Dt_Now.Subtract((DateTime)Dt_temp.Rows[0]["mbl_unlock_date"]).TotalDays;

                                if (Days >= 2)
                                    bRet = true;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                bRet = true;
            }
            if (DB != null)
                DB.CloseConnection();
            return bRet;
        }


        public Boolean UpdateDate(string type, string ID, DateTime  Dt, out string ErrorMessage)
        {
            Boolean bRet = false;
            ErrorMessage = "";
            DataBase.Connections.DBConnection DB = null;
            string sql = "";
            try
            {
                DB = new DataBase.Connections.DBConnection();
                DB.BeginTransaction();
                    sql = "";
                    sql += " update acc_ledgerd set jv_master_date = '" + Dt.ToString("yyyy-MM-dd") + "' ";
                    sql += " from  cargo_invoicem a inner join acc_ledgerd b on a.inv_pkid = b.jv_header_id ";
                    sql += " where inv_mbl_id = '" +  ID +"' ";
                    DB.ExecuteNonQuery(sql);
                DB.CommitTransaction();
                bRet = true;
            }
            catch (Exception Ex)
            {
                bRet = false;
                DB.RollbackTransaction();
                ErrorMessage = Ex.Message.ToString();
            }
            return bRet;
        }






        public  Boolean UpdateProfit(string MBLID, string CATEGORY)
        {
            // VERSION - 101
            Boolean bRet = false;
            Boolean bDetails = true;
            string sql = "";
            decimal MAR_Total = 0, MAP_Total = 0;

            decimal HAR_Total_wt = 0, HAP_Total_wt = 0;
            decimal HAR_Total_cbm = 0, HAP_Total_cbm = 0;
            decimal AR = 0, AP = 0;
            decimal mFactor_wt = 0, hFactor_wt = 0;
            decimal mFactor_cbm = 0, hFactor_cbm = 0;
            decimal nProfit = 0;
            decimal nProfit_wt = 0;
            decimal nProfit_cbm = 0;
            decimal EXP_TOTAL = 0, INC_TOTAL = 0, PER = 0;

            int ZeroWt = 0;
            int ZeroCbmOrWt = 0;

            DataRow DR_HEXP = null;
            DataTable DT_MEXP = new DataTable();
            DataTable DT_HEXP = new DataTable();
            DataTable DT_HOUSE = new DataTable();

            int iHouseTot = 0;

            string mCol_1 = "";
            string hCol_1 = "";

            string mCol_2 = "";
            string hCol_2 = "";

            bDetails = false;
            DataBase.Connections.DBConnection DB = null;

            try
            {
                DB = new DataBase.Connections.DBConnection();
                if (CATEGORY == "OI" || CATEGORY == "OE")
                {
                    mCol_1 = "MBL_WEIGHT";
                    mCol_2 = "MBL_CBM";

                    hCol_1 = "HBL_WEIGHT";
                    hCol_2 = "HBL_CBM";
                    bDetails = true;
                }
                if (CATEGORY == "AI" || CATEGORY == "AE") 
                {
                    mCol_1 = "MBL_WEIGHT";
                    mCol_2 = "MBL_CHWT";
                    
                    hCol_1 = "HBL_WEIGHT";
                    hCol_2 = "HBL_CHWT";
                    bDetails = true;
                }

                if (CATEGORY == "OT" || CATEGORY == "EX")
                {
                    mCol_1 = "MBL_CBM";
                    hCol_1 = "HBL_CBM";
                    
                    mCol_2 = "MBL_CHWT";
                    hCol_2 = "HBL_CHWT";
                    bDetails = true;
                }


                // Sum AR and AP Total For Master
                sql = " select  ";
                sql += " sum(case when inv_arap = 'AR' then inv_total else 0 end) as income, ";
                sql += " sum(case when inv_arap = 'AP' then inv_total else 0 end) as expense, ";
                sql += " sum(case when inv_arap = 'AR' and inv_cost_type = 'M' then inv_total else 0 end) as ar_total, ";
                sql += " sum(case when inv_arap = 'AP' and inv_cost_type = 'M' then inv_total else 0 end) as ap_total ";
                sql += " from cargo_invoicem a where  ";
                sql += " inv_mbl_id = '" + MBLID + "' and rec_deleted = 'N'";
                DT_MEXP = DB.ExecuteQuery(sql);

                // Sum AR and AP Total For House
                sql = " select  inv_hbl_id, ";
                sql += " sum(case when inv_arap = 'AR' then inv_total else 0 end) as ar_total, ";
                sql += " sum(case when inv_arap = 'AP' then inv_total else 0 end) as ap_total ";
                sql += " from cargo_invoicem a  where   ";
                sql += " inv_mbl_id = '" + MBLID + "' and inv_cost_type = 'H' and rec_deleted = 'N'";
                sql += " group by inv_hbl_id ";
                DT_HEXP = DB.ExecuteQuery(sql);
                sql = "";

                sql = " select mbl_pkid,hbl_pkid,mbl_zero_wt,mbl_zero_cbm,mbl_zero_chwt,mbl_weight,mbl_cbm,mbl_chwt, hbl_weight,hbl_cbm,hbl_chwt ";
                sql += " from cargo_masterm a inner join cargo_housem b on a.mbl_pkid = b.hbl_mbl_id ";
                sql += " where mbl_pkid = '" + MBLID + "'";
                // House Wise Weight And CBM For Master And House
                DT_HOUSE = DB.ExecuteQuery(sql);

                
                iHouseTot = DT_HOUSE.Rows.Count;
                ZeroWt = 0;
                ZeroCbmOrWt = 0;

                if (iHouseTot > 0)
                {
                    ZeroWt = Lib.Conv2Integer(DT_HOUSE.Rows[0]["mbl_zero_wt"].ToString());
                    if (CATEGORY == "OI" || CATEGORY == "OE")
                        ZeroCbmOrWt = Lib.Conv2Integer(DT_HOUSE.Rows[0]["mbl_zero_cbm"].ToString());
                    else
                        ZeroCbmOrWt = Lib.Conv2Integer(DT_HOUSE.Rows[0]["mbl_zero_chwt"].ToString());
                }

                string Profit_Type_Wt = "E"; // Equal
                string Profit_Type_CbmOrWt = "E"; // Divide

                 


                if (iHouseTot > 0)
                {
                    Profit_Type_Wt  = (ZeroWt  == 0) ? "D" : "E";
                    Profit_Type_CbmOrWt = (ZeroCbmOrWt == 0) ? "D" : "E";
                }

                // Master AR total And AP total
                MAR_Total = 0; MAP_Total = 0;
                if (DT_MEXP.Rows.Count > 0)
                {
                    INC_TOTAL = Lib.Conv2Decimal(DT_MEXP.Rows[0]["income"].ToString());
                    EXP_TOTAL = Lib.Conv2Decimal(DT_MEXP.Rows[0]["expense"].ToString());
                
                    MAR_Total = Lib.Conv2Decimal(DT_MEXP.Rows[0]["ar_total"].ToString());
                    MAP_Total = Lib.Conv2Decimal(DT_MEXP.Rows[0]["ap_total"].ToString());
                }
                // Master Expense and Income
                DataColumn[] Pkey = new DataColumn[1];
                Pkey[0] = DT_HEXP.Columns["inv_hbl_id"];
                DT_HEXP.PrimaryKey = Pkey;

                DB.BeginTransaction();

                if (bDetails == true)
                {
                    foreach (DataRow Dr in DT_HOUSE.Rows)
                    {
                        mFactor_wt = Lib.Conv2Decimal(Dr[mCol_1].ToString());
                        hFactor_wt = Lib.Conv2Decimal(Dr[hCol_1].ToString());

                        mFactor_cbm = Lib.Conv2Decimal(Dr[mCol_2].ToString());
                        hFactor_cbm = Lib.Conv2Decimal(Dr[hCol_2].ToString());
                        if (Profit_Type_Wt == "E" || iHouseTot ==1)
                        {
                            mFactor_wt = iHouseTot;
                            hFactor_wt = 1;
                        }
                        if (Profit_Type_CbmOrWt == "E" || iHouseTot ==1)
                        {
                            mFactor_cbm = iHouseTot;
                            hFactor_cbm = 1;
                        }

                        DR_HEXP = DT_HEXP.Rows.Find(Dr["HBL_PKID"].ToString());

                        HAR_Total_wt = 0; HAP_Total_wt = 0;
                        HAR_Total_cbm = 0; HAP_Total_cbm = 0;


                        //if (mFactor_wt <= 0)
                        //    mFactor_wt = 1;
                        //if (iHouseTot ==1 && hFactor_wt <= 0)
                        //    hFactor_wt = 1;
                        //if (mFactor_cbm <= 0)
                        //    mFactor_cbm = 1;
                        //if (iHouseTot == 1 && hFactor_cbm <= 0)
                        //    hFactor_cbm = 1;
                        //if (CATEGORY == "OT") // Only single row
                        //{
                        //    mFactor_wt = 1;
                        //    hFactor_wt = 1;
                        //    mFactor_cbm = 1;
                        //    hFactor_cbm = 1;
                        //}

                        
                        if (DR_HEXP != null)
                        {
                            HAR_Total_wt = Lib.Conv2Decimal(DR_HEXP["ar_total"].ToString());
                            HAP_Total_wt = Lib.Conv2Decimal(DR_HEXP["ap_total"].ToString());
                            HAR_Total_cbm = HAR_Total_wt;
                            HAP_Total_cbm = HAP_Total_wt;
                        }


                        AR = 0; AP = 0;
                        if (MAR_Total > 0)
                            AR = MAR_Total * hFactor_wt / mFactor_wt;
                        if (MAP_Total > 0)
                            AP = MAP_Total * hFactor_wt / mFactor_wt;
                        HAR_Total_wt += AR; HAP_Total_wt += AP;
                        nProfit_wt = HAR_Total_wt - HAP_Total_wt;

                        AR = 0; AP = 0;
                        if (MAR_Total > 0)
                            AR = MAR_Total * hFactor_cbm / mFactor_cbm;
                        if (MAP_Total > 0)
                            AP = MAP_Total * hFactor_cbm / mFactor_cbm;
                        HAR_Total_cbm += AR; HAP_Total_cbm += AP;
                        nProfit_cbm = HAR_Total_cbm - HAP_Total_cbm;

                        sql = " update cargo_housem set ";
                        sql += " hbl_inc_total_wt = " + HAR_Total_wt.ToString() + ",";
                        sql += " hbl_exp_total_wt = " + HAP_Total_wt.ToString() + ",";
                        sql += " hbl_inc_total_cbm = " + HAR_Total_cbm.ToString() + ",";
                        sql += " hbl_exp_total_cbm = " + HAP_Total_cbm.ToString() + ",";
                        sql += " hbl_revenue_wt = " + nProfit_wt.ToString() + ",";
                        sql += " hbl_revenue_cbm = " + nProfit_cbm.ToString() + "";
                        sql += " where hbl_pkid = '" + Dr["HBL_PKID"].ToString() + "'";
                        DB.ExecuteNonQuery(sql);
                    }
                }

                PER = 0;
                nProfit = INC_TOTAL - EXP_TOTAL;
                if (INC_TOTAL != 0)
                {
                    PER = nProfit / INC_TOTAL * 100;
                    PER = Lib.Conv2Decimal(Lib.NumFormat(PER.ToString(), 2));
                }
                sql = "";
                sql = " update cargo_masterm set ";
                sql += " mbl_profit_type ='" + Profit_Type_Wt + Profit_Type_CbmOrWt + "',";
                sql += " mbl_inc_total = " + INC_TOTAL.ToString() + ", ";
                sql += " mbl_exp_total = " + EXP_TOTAL + ",";
                sql += " mbl_revenue = " + nProfit.ToString() + ",";
                sql += " mbl_per = " + PER.ToString();
                sql += " where mbl_pkid = '" + MBLID + "'";
                DB.ExecuteNonQuery( sql);
                DB.CommitTransaction();
                bRet = true;
            }
            catch (Exception)
            {
                bRet = false;
                DB.RollbackTransaction();
            }
            return bRet;
        }








    }

}
