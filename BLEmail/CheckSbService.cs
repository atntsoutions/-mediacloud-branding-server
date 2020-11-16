using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using Message = OpenPop.Mime.Message;

using XL.XSheet;

namespace BLEmail
{
    public class CheckSbService : BL_Base
    {
        private string type = "";
        private string rowtype = "";
        private string pkid = "";
        private string searchstring = "";
        private string company_code = "";
        private string branch_code = "";
        private string year_code = "";
        private string br_icegate_email = "";
        private string br_icegate_email_pwd = "";
        private string br_custom_locations = "";
        private int br_start_index = 0;
        private string savemsg = "";
        private string ManualUpdtSB_Reason = "";
        private string from_date = "";
        private string ErrorMessage = "";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            ErrorMessage = "";
            string[] sdata = null;
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<mailsb> mList = new List<mailsb>();
            mailsb mRow;

            type = SearchData["type"].ToString();
            rowtype = SearchData["rowtype"].ToString();
            searchstring = SearchData["searchstring"].ToString().ToUpper();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();
            from_date = SearchData["from_date"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                from_date = Lib.StringToDate(from_date);

                if (from_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                sWhere = " where a.rec_category = '{CATEGORY}'";
                sWhere += " and a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.rec_branch_code = '{BRCODE}'";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.sb_no) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.sb_no2) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.sb_job_no) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }
                else
                    sWhere += "  and a.sb_msg_date >= to_date('{FDATE}','DD-MON-YYYY') ";

                sWhere = sWhere.Replace("{CATEGORY}", rowtype);
                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);
                sWhere = sWhere.Replace("{FDATE}", from_date);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM mailsb  a ";
                    sql += sWhere;
                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        page_rowcount = Lib.Conv2Integer(Dt_Temp.Rows[0]["total"].ToString());
                        page_count = Lib.Conv2Integer(Dt_Temp.Rows[0]["page_total"].ToString());
                    }
                    page_current = 1;
                    //page_current = page_count;
                }
                else
                {
                    if (type == "FIRST")
                        page_current = 1;
                    if (type == "PREV" && page_current > 1)
                        page_current--;
                    if (type == "NEXT" && page_current < page_count)
                        page_current++;
                    if (type == "LAST")
                        page_current = page_count;
                }

                startrow = (page_current - 1) * page_rows + 1;
                endrow = (startrow + page_rows) - 1;

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select * from ( ";
                sql += " select sb_pkid ,sb_id,sb_from,sb_subject,";
                sql += " sb_msg_date, sb_msg_type, sb_job_no, ";
                sql += " sb_job_date ,sb_no,sb_date,sb_no2,sb_reason, ";
                sql += " row_number() over(order by sb_msg_date,sb_id) rn ";
                sql += " from mailsb a ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by sb_msg_date,sb_id";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
       
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new mailsb();
                    mRow.sb_pkid = Dr["sb_pkid"].ToString();
                    if (Dr["sb_from"].ToString().Contains("@"))
                    {
                        sdata = Dr["sb_from"].ToString().Split('@');
                        mRow.sb_from = sdata[0];
                    }
                    else
                        mRow.sb_from = Dr["sb_from"].ToString();
                    mRow.sb_subject = Dr["sb_subject"].ToString();
                    mRow.sb_msg_date = Lib.DatetoStringDisplayformat(Dr["sb_msg_date"]);
                    mRow.sb_msg_type = Dr["sb_msg_type"].ToString();
                    mRow.sb_job_no = Lib.Conv2Integer(Dr["sb_job_no"].ToString());
                    mRow.sb_job_date = Lib.DatetoStringDisplayformat(Dr["sb_job_date"]);
                    mRow.sb_no = Lib.Conv2Integer(Dr["sb_no"].ToString());
                    mRow.sb_date = Lib.DatetoStringDisplayformat(Dr["sb_date"]);
                    mRow.sb_no2 = Dr["sb_no2"].ToString();
                    mRow.sb_reason = Dr["sb_reason"].ToString();
                    if (Dr["sb_msg_type"].ToString() == "P")
                        mRow.row_colour = "GREEN";
                    else
                        mRow.row_colour = "RED";
                    mList.Add(mRow);
                }
                Dt_List.Rows.Clear();
                Con_Oracle.CloseConnection();

                GetSettings();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);
            RetData.Add("email", br_icegate_email);
            RetData.Add("emailpwd", br_icegate_email_pwd);
            RetData.Add("locations", br_custom_locations);
            return RetData;
        }

        private void GetSettings()
        {
            sql = " select name,caption from settings ";
            sql += "   where parentid = '" + branch_code + "' ";
            sql += "   and code = '" + rowtype + "' ";
            sql += "   and caption in ('BR_ICEGATE_EMAIL_PWD','BR_CUSTOM_LOCATIONS','BR_ICEGATE_EMAIL')";
            DataTable Dt_Temp = new DataTable();
            Con_Oracle = new DBConnection();
            Dt_Temp = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow Dr in Dt_Temp.Rows)
            {
                if(Dr["caption"].ToString()== "BR_ICEGATE_EMAIL")
                    br_icegate_email = Dr["name"].ToString();
                else if (Dr["caption"].ToString() == "BR_ICEGATE_EMAIL_PWD")
                    br_icegate_email_pwd = Dr["name"].ToString();
                else if (Dr["caption"].ToString() == "BR_CUSTOM_LOCATIONS")
                    br_custom_locations = Dr["name"].ToString();
            }
            Dt_Temp.Rows.Clear();

            //LovService lov = new LovService();
            //DataRow lovRow_Icegate_Email = lov.getSettings(branch_code, "BR_ICEGATE_EMAIL");
            //if (lovRow_Icegate_Email != null)
            //    br_icegate_email = lovRow_Icegate_Email["name"].ToString();
            //DataRow lovRow_Icegate_Email_Pwd = lov.getSettings(branch_code, "BR_ICEGATE_EMAIL_PWD");
            //if (lovRow_Icegate_Email_Pwd != null)
            //    br_icegate_email_pwd = lovRow_Icegate_Email_Pwd["name"].ToString();
            //DataRow lovRow_Custom_Locations = lov.getSettings(branch_code, "BR_CUSTOM_LOCATIONS");
            //if (lovRow_Custom_Locations != null)
            //    br_custom_locations = lovRow_Custom_Locations["name"].ToString();
        }
        public Dictionary<string, object> SaveSettings(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            savemsg = "";
            type = SearchData["type"].ToString();
            rowtype = SearchData["rowtype"].ToString();
            pkid = SearchData["pkid"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();
            br_icegate_email = SearchData["br_icegate_email"].ToString();
            br_icegate_email_pwd = SearchData["br_icegate_email_pwd"].ToString();
            br_custom_locations = SearchData["br_custom_locations"].ToString();
            br_start_index = Lib.Conv2Integer(SearchData["br_start_index"].ToString());

            if (type == "SAVE")
            {
                SaveData(branch_code, "BR_ICEGATE_EMAIL", br_icegate_email,rowtype);
                SaveData(branch_code, "BR_ICEGATE_EMAIL_PWD", br_icegate_email_pwd, rowtype);
                SaveData(branch_code, "BR_CUSTOM_LOCATIONS", br_custom_locations, rowtype);
            }
            else if (type == "DOWNLOAD")
            {
                ProcessSB();
            }
            else if (type == "UPDATESB")
            {
                string JobNo = "";
                string JobBookDate = "";
                string SbNo = "";
                string SbDate = "";
                string SID = "";
                string SMailID = "";

                sql = "select * from mailsb where sb_pkid ='" + pkid + "'";
                DataTable Dt_Temp = new DataTable();
                Con_Oracle = new DBConnection();
                Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                if (Dt_Temp.Rows.Count > 0)
                {
                    JobNo = Dt_Temp.Rows[0]["SB_JOB_NO"].ToString();
                    if (!Dt_Temp.Rows[0]["SB_JOB_DATE"].Equals(DBNull.Value))
                        JobBookDate = Lib.StringToDate(Dt_Temp.Rows[0]["SB_JOB_DATE"]);

                    SbNo = Dt_Temp.Rows[0]["SB_NO"].ToString();
                    if (!Dt_Temp.Rows[0]["SB_DATE"].Equals(DBNull.Value))
                        SbDate = Lib.StringToDate(Dt_Temp.Rows[0]["SB_DATE"]);

                    SID = Dt_Temp.Rows[0]["SB_ID"].ToString();
                    SMailID = Dt_Temp.Rows[0]["SB_MAILID"].ToString();
                }

                if (SbNo.Trim().Length > 0 && SbDate.Trim().Length > 0)
                {
                    if (UpdateSBNO(JobNo, JobBookDate, SbNo, SbDate, SID, SMailID, "Y"))
                        savemsg = "SUCCESSFULLY UPDATED";
                    else
                        savemsg = "INVALID EDI JOB# " + JobNo + ", CANNOT UPDATE";
                }
                else
                    savemsg = "INVALID SB#";
            }

            RetData.Add("savemsg", savemsg);
            RetData.Add("sbreason", ManualUpdtSB_Reason);
            return RetData;
        }

        private void SaveData(string sParentid, string sCaption, string sName,string sCode)
        {
            try
            {
                if (sParentid == "" || sCaption == "" || sCode == "")
                    return;

                Con_Oracle = new DBConnection();
                Con_Oracle.BeginTransaction();
                sql = "delete from settings where parentid = '" + sParentid + "' and caption = '" + sCaption + "'";
                sql += " and code = '" + sCode + "'";
                Con_Oracle.ExecuteQuery(sql);
                DBRecord Rec = new DBRecord();
                Rec.CreateRow("settings", "ADD", "caption", sCaption);
                Rec.InsertString("parentid", sParentid);
                Rec.InsertString("tablename", "TEXT");
                Rec.InsertString("id", "");
                Rec.InsertString("code", sCode);
                Rec.InsertString("name", sName, "P");
                sql = Rec.UpdateRow();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
            Con_Oracle.CloseConnection();
        }

        private void ProcessSB()
        {
            string sql = "";
            int LastRdIndex = 0;
            string UserName = "";
            string UserPwd = "";


            GetSettings();
            UserName = br_icegate_email;
            UserPwd = br_icegate_email_pwd;

            if (UserName.Trim() == "" || UserPwd.Trim() == "")
            {
                throw new Exception("INVALID CREDENTIALS.....[DOWNLOAD]");
            }

            LastRdIndex = GetLastReadIndex(UserName);
            if (LastRdIndex <= 0)
                LastRdIndex = br_start_index; //Value From front end

            if (LastRdIndex <= 0)
            {
                throw new Exception("DOWNLOAD SEQUENCE NOT SET.....[DOWNLOAD MESSAGE]");
            }

            if (br_custom_locations.Trim() == "")
            {
                throw new Exception("CUSTOMS LOCATIONS NOT SET.....[DOWNLOAD MESSAGE]");
            }

            //UserName = "softwaresupport@cargomar.in";
            //UserPwd = "CPLCSPSUP55#8";

            DownloadMessages("mail.eximusmail.com", 995, true, UserName, UserPwd, LastRdIndex);
        }

        private int GetLastReadIndex(string Usr_Name)
        {
            string sql = "";
            int iLast = 0;
            try
            {   
                sql = " select nvl(max(sb_id),0) as slno from mailsb ";
                sql += " where rec_company_code='" + company_code + "'";
                sql += " and rec_branch_code='" + branch_code + "'";
                sql += " and rec_category ='" + rowtype + "'";
                sql += " and sb_mailid='" + Usr_Name + "'";
                sql += " and sb_old is null ";

                Con_Oracle = new DBConnection();
                DataTable dt_temp = new DataTable();
                dt_temp = Con_Oracle.ExecuteQuery(sql);
                if (dt_temp.Rows.Count > 0)
                    iLast = Lib.Conv2Integer(dt_temp.Rows[0]["slno"].ToString());
                Con_Oracle.CloseConnection();

                dt_temp.Rows.Clear();
            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw ex;
            }
            return iLast;
        }

        private void DownloadMessages(string hostname, int port, bool useSsl, string username, string password, int LastSeenIndex)
        {
            try
            {
                string TxtMessage = "";
                string MsgAckType = "";
                string currentUid = "";
                string[] sdata = null;
                bool UnReadMessage = false;
                Message newMessage = null;
                int nUID = 0;
                int count = 0;

                using (Pop3Client client = new Pop3Client())
                {
                    client.Connect(hostname, port, useSsl);
                    client.Authenticate(username, password);

                    List<string> uids = client.GetMessageUids();
                    count = client.GetMessageCount();

                    for (int i = 0; i < count; i++)
                    {
                        currentUid = uids[i];

                        if (UnReadMessage == false)
                        {
                            sdata = currentUid.Split('.');
                            nUID = Convert.ToInt32(sdata[0]);
                            if (nUID > LastSeenIndex)
                                UnReadMessage = true;
                        }
                        if (UnReadMessage)
                        {
                            if (client.GetMessageHeaders(i + 1).From.Address.ToUpper().Contains("ICEGATE") && client.GetMessageHeaders(i + 1).Subject.ToUpper().StartsWith("CHCAE02"))
                            {
                                newMessage = client.GetMessage(i + 1);
                                TxtMessage = "";
                                MsgAckType = "";
                                List<MessagePart> attachments = newMessage.FindAllAttachments();
                                foreach (MessagePart attachment in attachments)
                                {
                                    if (attachment.FileName.ToUpper().EndsWith(".ACK"))
                                    {
                                        MsgAckType = "P";
                                        TxtMessage = attachment.GetBodyAsText();
                                    }
                                    else if (attachment.FileName.ToUpper().EndsWith(".NAK"))
                                    {
                                        MsgAckType = "N";
                                        TxtMessage = attachment.GetBodyAsText();
                                    }
                                }
                                if (TxtMessage.Trim().Length > 0)
                                {
                                    InsertMailSB(currentUid, newMessage.Headers.From.Address.ToString(), newMessage.Headers.Subject.ToString(), newMessage.Headers.DateSent, TxtMessage, MsgAckType, username);
                                }
                            }
                        }
                    }
                }

            }
            catch (InvalidLoginException)
            {
                throw new Exception("The server did not accept the user credentials!.....[POP3 Server Authentication]");
            }
            catch (PopServerNotFoundException)
            {
                throw new Exception("The server could not be found.....[POP3 Retrieval]");
            }
            catch (PopServerLockedException)
            {
                throw new Exception("The mailbox is locked. It might be in use or under maintenance. Are you connected elsewhere?.....[POP3 Account Locked]");
            }
            catch (LoginDelayException)
            {
                throw new Exception("Login not allowed. Server enforces delay between logins. Have you connected recently?.....[POP3 Account Login Delay]");
            }
            catch (Exception e)
            {
                throw new Exception("Error occurred retrieving mail " + e.Message + "......[POP3 Retrieval]");
            }
            finally
            {

            }
        }

        private bool InsertMailSB(string MsgUID, string MsgFrom, string MsgSubject, DateTime MsgDate, string TxtMessage, string MsgAckType, string MsgEmailID)
        {
            string SQL = "";
            bool bRet = false;
            int nUID = 0;
            string[] sdata = null;
            char SEP_CHAR = Convert.ToChar(29);
            string MessageCategory = "";
            string JobNo = "";
            string JobBookDate = "";
            string SbNo = "";
            string SbDate = "";
            string MessageLocation = "";

            sdata = MsgUID.Split('.');
            nUID = Convert.ToInt32(sdata[0]);
            MsgFrom = MsgFrom.Replace("'", "''").ToString().ToUpper();
            MsgSubject = MsgSubject.Replace("'", "''").ToString().ToUpper();

            TxtMessage = TxtMessage.Replace("\r", "");
            string[] sLins = TxtMessage.Split('\n');
            string[] sCols = null;

            if (sLins.Length > 0)
                sCols = sLins[0].Split(SEP_CHAR);

            if (sCols != null && sCols.Length > 2)
                MessageLocation = sCols[2];

            if (sCols != null && sCols.Length > 8)
                MessageCategory = sCols[8];
            if (MessageCategory.Trim() == "CHCAE02" && br_custom_locations.Contains(MessageLocation))// if (MessageCategory.Trim() == "CHCAE02")
            {
                sCols = null;
                if (sLins.Length > 1)
                    sCols = sLins[1].Split(SEP_CHAR);

                if (MsgAckType.Trim() == "")
                    MsgAckType = "N";

                Con_Oracle = new DBConnection();
                try
                {
                    sql = "select sb_id from mailsb ";
                    sql += " where rec_company_code ='" + company_code + "'";
                    sql += " and rec_branch_code ='" + branch_code + "'";
                    sql += " and rec_category ='" + rowtype + "'";
                    sql += " and sb_mailid ='" + MsgEmailID + "'";
                    sql += " and sb_id = " + nUID.ToString();

                    DataTable Dt_temp = new DataTable();
                    Dt_temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_temp.Rows.Count <= 0)
                    {
                        SQL = " INSERT INTO MAILSB ( ";
                        SQL += "   SB_PKID,SB_MAILID,SB_ID,SB_FROM,SB_SUBJECT,SB_MSG_DATE ";
                        SQL += "   ,SB_MSG_TYPE,SB_JOB_NO,SB_JOB_DATE ";
                        SQL += "   ,SB_NO,SB_DATE,SB_REASON,REC_COMPANY_CODE,REC_BRANCH_CODE,REC_CATEGORY) ";
                        SQL += "  SELECT ";
                        SQL += "   [SB_PKID],[SB_MAILID],[SB_ID],[SB_FROM],[SB_SUBJECT],[SB_MSG_DATE] ";
                        SQL += "   ,[SB_MSG_TYPE],[SB_JOB_NO],[SB_JOB_DATE] ";
                        SQL += "   ,[SB_NO],[SB_DATE],[SB_REASON],[REC_COMPANY_CODE],[REC_BRANCH_CODE],[REC_CATEGORY] ";
                        SQL += "   FROM DUAL WHERE NOT EXISTS (SELECT SB_ID FROM MAILSB WHERE SB_ID = [SB_ID] AND SB_MAILID = [SB_MAILID] ";
                        SQL += "   AND REC_BRANCH_CODE = [REC_BRANCH_CODE] AND REC_CATEGORY = [REC_CATEGORY] )";

                        SQL = SQL.Replace("[SB_PKID]", "'" + Guid.NewGuid().ToString().ToUpper() + "'");
                        SQL = SQL.Replace("[SB_MAILID]", "'" + GetSubStr(MsgEmailID, 150) + "'");
                        SQL = SQL.Replace("[SB_ID]", nUID.ToString());
                        SQL = SQL.Replace("[SB_FROM]", "'" + GetSubStr(MsgFrom, 100) + "'");
                        SQL = SQL.Replace("[SB_SUBJECT]", "'" + GetSubStr(MsgSubject, 100) + "'");
                        if (MsgDate != null)
                            SQL = SQL.Replace("[SB_MSG_DATE]", "'" + Lib.StringToDate(MsgDate) + "'");
                        else
                            SQL = SQL.Replace("[SB_MSG_DATE]", "Null");

                        SQL = SQL.Replace("[SB_MSG_TYPE]", "'" + MsgAckType + "'");
                        SQL = SQL.Replace("[REC_COMPANY_CODE]", "'" + company_code + "'");
                        SQL = SQL.Replace("[REC_BRANCH_CODE]", "'" + branch_code + "'");
                        SQL = SQL.Replace("[REC_CATEGORY]", "'" + rowtype + "'");

                        if (MsgAckType == "P")
                        {
                            if (sCols != null && sCols.Length > 1)
                            {
                                JobNo = sCols[1];
                                SQL = SQL.Replace("[SB_JOB_NO]", JobNo);
                            }
                            else
                                SQL = SQL.Replace("[SB_JOB_NO]", "NULL");

                            if (sCols != null && sCols.Length > 2)
                            {
                                JobBookDate = GetBackEndFormatedDate(sCols[2], "JOB");
                                SQL = SQL.Replace("[SB_JOB_DATE]", "'" + JobBookDate + "'");
                            }
                            else
                                SQL = SQL.Replace("[SB_JOB_DATE]", "NULL");

                            if (sCols != null && sCols.Length > 3)
                            {
                                SbNo = sCols[3];
                                SQL = SQL.Replace("[SB_NO]", SbNo);
                            }
                            else
                                SQL = SQL.Replace("[SB_NO]", "NULL");

                            if (sCols != null && sCols.Length > 4)
                            {
                                SbDate = GetBackEndFormatedDate(sCols[4], "SB");
                                SQL = SQL.Replace("[SB_DATE]", "'" + SbDate + "'");
                            }
                            else
                                SQL = SQL.Replace("[SB_DATE]", "NULL");

                            SQL = SQL.Replace("[SB_REASON]", "NULL");
                        }
                        else
                        {
                            if (sCols != null && sCols.Length > 1)
                                SQL = SQL.Replace("[SB_JOB_NO]", sCols[1]);
                            else
                                SQL = SQL.Replace("[SB_JOB_NO]", "NULL");

                            if (sCols != null && sCols.Length > 2)
                            {
                                SQL = SQL.Replace("[SB_JOB_DATE]", "'" + GetBackEndFormatedDate(sCols[2], "JOB") + "'");
                            }
                            else
                                SQL = SQL.Replace("[SB_JOB_DATE]", "NULL");
                            SQL = SQL.Replace("[SB_NO]", "NULL");
                            SQL = SQL.Replace("[SB_DATE]", "NULL");

                            SQL = SQL.Replace("[SB_REASON]", "'" + GetSubStr(sCols[3], 250).Replace("'", "''").ToUpper() + "'");
                        }

                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteNonQuery(SQL);
                        Con_Oracle.CommitTransaction();

                        if (MsgAckType == "P" && JobNo.Trim().Length > 0 && SbNo.Trim().Length > 0)
                            UpdateSBNO(JobNo, JobBookDate, SbNo, SbDate, nUID.ToString(), MsgEmailID);
                    }

                }
                catch (Exception Ex)
                {
                    bRet = false;
                    if (Con_Oracle != null)
                    {
                        Con_Oracle.RollbackTransaction();
                        Con_Oracle.CloseConnection();
                    }
                    throw Ex;
                }
                Con_Oracle.CloseConnection();
            }
            bRet = true;
            return bRet;
        }

        private string GetSubStr(string str, int len)
        {
            if (str.Length > len)
                str = str.Substring(0, len);
            return str;
        }

        private string GetBackEndFormatedDate(string sDate, string Stype)
        {
            int d = 0, m = 0, y = 0;
            if (sDate.Length < 8)
                return null;
            try
            {
                if (Stype == "JOB")
                {
                    d = Lib.Conv2Integer(sDate.Substring(0, 2));
                    m = Lib.Conv2Integer(sDate.Substring(2, 2));
                    y = Lib.Conv2Integer(sDate.Substring(4, 4));
                }
                else if (Stype == "SB")
                {
                    y = Lib.Conv2Integer(sDate.Substring(0, 4));
                    m = Lib.Conv2Integer(sDate.Substring(4, 2));
                    d = Lib.Conv2Integer(sDate.Substring(6, 2));
                }
            }
            catch (Exception)
            {
                return null;
            }
            return Lib.StringToDate(new DateTime(y, m, d));
        }

        private bool UpdateSBNO(string JobNo, string JobBookDate, string SbNo, string SbDate, string SID, string MsgEmailID, string ManualUpdt = "")
        {
            string sql = "";
            DataTable Dt_temp;
            string JOB_PKID = "";
            string JOB_NO = "";
            string SB_REASON = "";
            ManualUpdtSB_Reason = "";
            string PRE_SB_NO = "";
            bool bRet = false;
            Con_Oracle = new DBConnection();

            sql = "select job_pkid,job_docno from jobm ";
            sql += " where rec_company_code ='" + company_code + "'";
            sql += " and rec_branch_code ='" + branch_code + "'";
            sql += " and rec_category ='" + rowtype + "'";
            sql += " and job_date ='" + JobBookDate + "'";
            sql += " and job_edi_no = " + JobNo.ToString();
            Dt_temp = new DataTable();
            Dt_temp = Con_Oracle.ExecuteQuery(sql);
            if(Dt_temp.Rows.Count <=0 )
            {
                sql = "select job_pkid,job_docno from jobm ";
                sql += " where rec_company_code ='" + company_code + "'";
                sql += " and rec_branch_code ='" + branch_code + "'";
                sql += " and rec_category ='" + rowtype + "'";
                sql += " and job_date ='" + JobBookDate + "'";
                sql += " and job_docno = " + JobNo.ToString();
                Dt_temp = new DataTable();
                Dt_temp = Con_Oracle.ExecuteQuery(sql);
            }
            if (Dt_temp.Rows.Count > 0)
            {
                JOB_PKID = Dt_temp.Rows[0]["job_pkid"].ToString();
                JOB_NO = Dt_temp.Rows[0]["job_docno"].ToString();
            }

            if (JOB_PKID.Trim().Length > 0)
            {
                try
                {
                    sql = "select opr_job_id,opr_sbill_no from joboperationsm where opr_job_id = '{JOB_PKID}'"; //nvl(length(opr_sbill_no),0) <= 0 and
                    sql = sql.Replace("{JOB_PKID}", JOB_PKID);
                    Dt_temp = new DataTable();
                    Dt_temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_temp.Rows.Count <= 0)
                    {
                        sql = " INSERT INTO JOBOPERATIONSM ( ";
                        sql += "  OPR_JOB_ID,OPR_DRAWBACK_AMT,OPR_CRS,OPR_EGM_STATUS  ";
                        sql += "  ,OPR_SBILL_NO,OPR_SBILL_DATE ";
                        sql += "  ) VALUES (";
                        sql += "  [OPR_JOB_ID],[OPR_DRAWBACK_AMT],[OPR_CRS],[OPR_EGM_STATUS] ";
                        sql += "  ,[OPR_SBILL_NO],[OPR_SBILL_DATE] ";
                        sql += "  )";

                        sql = sql.Replace("[OPR_JOB_ID]", "'" + JOB_PKID + "'");
                        sql = sql.Replace("[OPR_DRAWBACK_AMT]", "0");
                        sql = sql.Replace("[OPR_CRS]", "'NIL'");
                        sql = sql.Replace("[OPR_EGM_STATUS]", "'EGM'");
                        sql = sql.Replace("[OPR_SBILL_NO]", "'" + SbNo + "'");
                        sql = sql.Replace("[OPR_SBILL_DATE]", "'" + SbDate + "'");

                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();

                        SB_REASON = "JOB# " + JOB_NO + ", SB# " + SbNo + ", SB DATE " + SbDate.ToUpper() + "  UPDATED";
                    }
                    else
                    {

                        if (Dt_temp.Rows[0]["opr_sbill_no"].ToString().Trim().Length <= 0 || ManualUpdt == "Y")
                        {
                            sql = "update joboperationsm set opr_sbill_no='{SBILL_NO}' ,opr_sbill_date='{SBILL_DATE}' where opr_job_id = '{JOB_PKID}'";
                            if (ManualUpdt != "Y")
                                sql += " and nvl(length(opr_sbill_no),0) <= 0 ";

                            sql = sql.Replace("{SBILL_NO}", SbNo);
                            sql = sql.Replace("{SBILL_DATE}", SbDate);
                            sql = sql.Replace("{JOB_PKID}", JOB_PKID);

                            Con_Oracle.BeginTransaction();
                            Con_Oracle.ExecuteNonQuery(sql);
                            Con_Oracle.CommitTransaction();

                            SB_REASON = "JOB# " + JOB_NO + ", SB# " + SbNo + ", SB DATE " + SbDate.ToUpper() + "  UPDATED";
                            if (ManualUpdt == "Y")
                                ManualUpdtSB_Reason = SB_REASON;
                        }
                        else
                            PRE_SB_NO = Dt_temp.Rows[0]["opr_sbill_no"].ToString().Trim();
                    }

                    sql = "update mailsb set  sb_no2 = '{SB_NO2}', sb_reason = '{REASON}' ";
                    sql += " where rec_company_code ='" + company_code + "'";
                    sql += " and rec_branch_code ='" + branch_code + "'";
                    sql += " and rec_category ='" + rowtype + "'";
                    sql += " and sb_mailid = '" + MsgEmailID + "'";
                    sql += " and sb_id = " + SID.ToString();

                    sql = sql.Replace("{REASON}", SB_REASON.ToUpper());
                    sql = sql.Replace("{JOBNO}", JOB_NO);
                    sql = sql.Replace("{SB_NO2}", PRE_SB_NO);

                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();

                    bRet = true;
                }
                catch (Exception Ex)
                {
                    bRet = false;
                    if (Con_Oracle != null)
                    {
                        Con_Oracle.RollbackTransaction();
                        Con_Oracle.CloseConnection();
                    }
                    throw Ex;
                }
            }
            Con_Oracle.CloseConnection();
            return bRet;
        }

    }

}

