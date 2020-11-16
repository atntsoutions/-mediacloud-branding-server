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
    public class EsanchitDownloadService : BL_Base
    {
        private string type = "";
        private string rowtype = "";
        private string pkid = "";
        private string searchstring = "";
        private string company_code = "";
        private string branch_code = "";
        private string year_code = "";
        private string br_esanchit_email = "";
        private string br_esanchit_email_pwd = "";
        private string br_esanchit_locations = "";
        private int br_start_index = 0;
        private string savemsg = "";
        private string ManualUpdtSB_Reason = "";
        private string from_date = "";
        private string ErrorMessage = "";
        private string user_code = "";
        private string jobid = "";
        private string jobno = "";
        private string linkid = "";
        private string linktype = "";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            ErrorMessage = "";
            string[] sdata = null;
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Esanchit> mList = new List<Esanchit>();
            Esanchit mRow;

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

                sWhere = " where a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.rec_branch_code = '{BRCODE}'";
                sWhere += " and a.doc_mail_id is not null ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.doc_drn) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.doc_irn) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }
                else
                    sWhere += "  and to_date(to_char(a.doc_upload_date,'DD-MON-YYYY') ,'DD-MON-YYYY') >= '{FDATE}' ";

                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);
                sWhere = sWhere.Replace("{FDATE}", from_date);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM edisupdocs  a ";
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
                sql += " select doc_pkid,doc_drn,doc_irn,doc_type_code,doc_name,doc_file_name,doc_upload_date ";
                sql += " ,c.param_name as doc_type_name";
                sql += " ,row_number() over(order by doc_upload_date,doc_drn,doc_irn) rn ";
                sql += " from edisupdocs a ";
                sql += " left join param c on (a.doc_type_code = c.param_code and c.param_type = 'ESANCHITDOC')";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow} ";
                sql += " order by doc_upload_date,doc_drn,doc_irn";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Esanchit();
                    mRow.doc_pkid = Dr["doc_pkid"].ToString();
                    mRow.doc_upload_date = Lib.DatetoStringDisplayformat(Dr["doc_upload_date"]);
                    mRow.doc_drn = Dr["doc_drn"].ToString();
                    mRow.doc_irn = Dr["doc_irn"].ToString();
                    mRow.doc_type_code = Dr["doc_type_code"].ToString();
                    mRow.doc_file_name = Dr["doc_file_name"].ToString();
                    if (Dr["doc_type_name"].ToString().Length > 50)
                        mRow.doc_name = Dr["doc_type_name"].ToString().Substring(0, 50);
                    else
                        mRow.doc_name = Dr["doc_type_name"].ToString();
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
            RetData.Add("email", br_esanchit_email);
            RetData.Add("emailpwd", br_esanchit_email_pwd);
            RetData.Add("locations", br_esanchit_locations);
            return RetData;
        }

        private void GetSettings()
        {
            sql = " select name,caption from settings ";
            sql += "   where parentid = '" + branch_code + "' ";
            sql += "   and code = '" + rowtype + "' ";
            sql += "   and caption in ('BR_ESANCHIT_LOCATIONS','BR_ESANCHIT_EMAIL','BR_ESANCHIT_EMAIL_PWD')";
            DataTable Dt_Temp = new DataTable();
            Con_Oracle = new DBConnection();
            Dt_Temp = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow Dr in Dt_Temp.Rows)
            {
                if (Dr["caption"].ToString() == "BR_ESANCHIT_EMAIL")
                    br_esanchit_email = Dr["name"].ToString();
                else if (Dr["caption"].ToString() == "BR_ESANCHIT_EMAIL_PWD")
                    br_esanchit_email_pwd = Dr["name"].ToString();
                else if (Dr["caption"].ToString() == "BR_ESANCHIT_LOCATIONS")
                    br_esanchit_locations = Dr["name"].ToString();
            }
            Dt_Temp.Rows.Clear();

        }
        public Dictionary<string, object> SaveSettings(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string sMsgError = "";

            savemsg = "";
            string cbdata = "";
            type = SearchData["type"].ToString();
            rowtype = SearchData["rowtype"].ToString();
            pkid = SearchData["pkid"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();
            user_code = SearchData["user_code"].ToString();
            br_esanchit_email = SearchData["br_esanchit_email"].ToString();
            br_esanchit_email_pwd = SearchData["br_esanchit_email_pwd"].ToString();
            br_esanchit_locations = SearchData["br_esanchit_locations"].ToString();
            br_start_index = Lib.Conv2Integer(SearchData["br_start_index"].ToString());
            if (SearchData.ContainsKey("cbdata"))
                cbdata = SearchData["cbdata"].ToString();

            if (type == "SAVE")
            {
                SaveData(branch_code, "BR_ESANCHIT_EMAIL", br_esanchit_email, rowtype);
                SaveData(branch_code, "BR_ESANCHIT_EMAIL_PWD", br_esanchit_email_pwd, rowtype);
                SaveData(branch_code, "BR_ESANCHIT_LOCATIONS", br_esanchit_locations, rowtype);
            }
            else if (type == "DOWNLOAD")
            {
                sMsgError = DownloadEsanchit();
            }
            else if (type == "PASTE-DATA")
            {
                sMsgError = SavePasteData(cbdata);
            }

            RetData.Add("errormsg", sMsgError);
            RetData.Add("savemsg", savemsg);
            return RetData;
        }

        private void SaveData(string sParentid, string sCaption, string sName, string sCode)
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

        private string DownloadEsanchit()
        {
            string sql = "";
            int LastRdIndex = 0;
            string UserName = "";
            string UserPwd = "";
            string sMsgError = "";

            GetSettings();
            UserName = br_esanchit_email;
            UserPwd = br_esanchit_email_pwd;

            if (UserName.Trim() == "" || UserPwd.Trim() == "")
            {
                throw new Exception("INVALID CREDENTIALS.....[DOWNLOAD]");
            }

            LastRdIndex = GetLastReadIndex(UserName);
            if (LastRdIndex <= 0)
                LastRdIndex = br_start_index; //Value From front end for the first time

            if (LastRdIndex <= 0)
            {
                throw new Exception("DOWNLOAD SEQUENCE NOT SET.....[DOWNLOAD MESSAGE]");
            }

            //if (br_custom_locations.Trim() == "")
            //{
            //    throw new Exception("CUSTOMS LOCATIONS NOT SET.....[DOWNLOAD MESSAGE]");
            //}

            //UserName = "softwaresupport@cargomar.in";
            //UserPwd = "CPLCSPSUP55#8";

            sMsgError = DownloadMessages("mail.eximusmail.com", 995, true, UserName, UserPwd, LastRdIndex);

            return sMsgError;
        }

        private int GetLastReadIndex(string Usr_Name)
        {
            string sql = "";
            int iLast = 0;
            try
            {
                sql = " select nvl(max(doc_email_slno),0) as slno from edisupdocs ";
                sql += " where rec_company_code='" + company_code + "'";
                sql += " and rec_branch_code='" + branch_code + "'";
                sql += " and doc_mail_id = '" + Usr_Name.ToUpper() + "'";
                sql += " and doc_old is null ";

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

        private string DownloadMessages(string hostname, int port, bool useSsl, string username, string password, int LastSeenIndex)
        {
            string MsgError = "";
            try
            {
                string TxtMessage = "";
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
                            try
                            {
                                if (client.GetMessageHeaders(i + 1).From.Address.ToUpper().Contains("ESANCHIT@ICEGATE") && client.GetMessageHeaders(i + 1).Subject.ToUpper().Contains("DOCUMENT UPLOAD CONFIRMATION"))
                                {
                                    newMessage = client.GetMessage(i + 1);
                                    TxtMessage = "";
                                    List<MessagePart> attachments = newMessage.FindAllAttachments();
                                    foreach (MessagePart attachment in attachments)
                                    {
                                        if (attachment.FileName.ToUpper().EndsWith(".DMS_"))
                                        {
                                            TxtMessage = attachment.GetBodyAsText();
                                        }
                                    }
                                    if (TxtMessage.Trim().Length > 0)
                                    {
                                        InsertEsanchit(currentUid, newMessage.Headers.From.Address.ToString(), newMessage.Headers.Subject.ToString(), newMessage.Headers.DateSent, TxtMessage, username);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (MsgError.Length > 0)
                                    MsgError += "\n" + ex.Message.ToString();
                                else
                                    MsgError += ex.Message.ToString();

                                if (Con_Oracle != null)
                                    Con_Oracle.CreateErrorLog("Esanchit Download " + ex.Message.ToString());
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
            return MsgError;
        }

        private bool InsertEsanchit(string MsgUID, string MsgFrom, string MsgSubject, DateTime MsgDate, string TxtMessage, string MsgEmailID)
        {
            string SQL = "";
            string SQL2 = "";
            bool bRet = false;
            int nUID = 0;
            string[] sdata = null;
            char SEP_CHAR = Convert.ToChar(29);
            string MessageCategory = "";
            string MessageType = "";
            string DocIRN = "";

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
                MessageCategory = sCols[2];
            if (sCols != null && sCols.Length > 5)
                MessageType = sCols[5];

            if (MessageCategory.Trim().ToUpper() == "ESANCHIT" && MessageType.Trim().ToUpper() == "ICES1_5")
            {
                Con_Oracle = new DBConnection();

                SQL = "INSERT INTO EDISUPDOCS ( ";
                SQL += " DOC_PKID,DOC_EMAIL_SLNO,REC_COMPANY_CODE,REC_BRANCH_CODE ,REC_CATEGORY,";
                SQL += " DOC_DRN,DOC_IRN,DOC_TYPE_CODE,DOC_FILE_TYPE,";
                SQL += " DOC_FILE_NAME,DOC_NAME,";
                SQL += " REC_CREATED_BY,REC_CREATED_DATE,DOC_UPLOAD_DATE,DOC_MAIL_ID";
                SQL += " )";
                SQL += " SELECT ";
                SQL += " [DOC_PKID],[DOC_EMAIL_SLNO],[REC_COMPANY_CODE],[REC_BRANCH_CODE],[REC_CATEGORY],";
                SQL += " [DOC_DRN],[DOC_IRN],[DOC_TYPE_CODE],[DOC_FILE_TYPE],";
                SQL += " [DOC_FILE_NAME],[DOC_NAME], ";
                SQL += " [REC_CREATED_BY],[REC_CREATED_DATE],[DOC_UPLOAD_DATE],[DOC_MAIL_ID]";
                SQL += " FROM DUAL WHERE NOT EXISTS (SELECT DOC_EMAIL_SLNO FROM EDISUPDOCS ";
                SQL += " WHERE REC_COMPANY_CODE = [REC_COMPANY_CODE]";
                SQL += " AND REC_BRANCH_CODE = [REC_BRANCH_CODE]";
                SQL += " AND DOC_IRN = [DOC_IRN]";
                SQL += " AND DOC_MAIL_ID = [DOC_MAIL_ID]";
                SQL += " AND DOC_EMAIL_SLNO = [DOC_EMAIL_SLNO]";
                SQL += " )";

                for (int LineIndex = 1; LineIndex < sLins.Length; LineIndex++)
                {
                    try
                    {
                        DocIRN = "";
                        sCols = sLins[LineIndex].Split(SEP_CHAR);
                        sql = SQL;

                        sql = sql.Replace("[DOC_PKID]", "'" + Guid.NewGuid().ToString().ToUpper() + "'");
                        sql = sql.Replace("[DOC_EMAIL_SLNO]", nUID.ToString());
                        sql = sql.Replace("[REC_COMPANY_CODE]", "'" + company_code + "'");
                        sql = sql.Replace("[REC_BRANCH_CODE]", "'" + branch_code + "'");
                        sql = sql.Replace("[REC_CATEGORY]", "NULL");

                        if (sCols != null && sCols.Length > 1)
                            sql = sql.Replace("[DOC_FILE_NAME]", "'" + GetSubStr(sCols[1], 200) + "'");
                        else
                            sql = sql.Replace("[DOC_FILE_NAME]", "NULL");

                        if (sCols != null && sCols.Length > 2)
                        {
                            DocIRN = GetSubStr(sCols[2], 16);
                            sql = sql.Replace("[DOC_IRN]", "'" + DocIRN + "'");
                        }
                        else
                            sql = sql.Replace("[DOC_IRN]", "NULL");

                        if (sCols != null && sCols.Length > 3)
                            sql = sql.Replace("[DOC_DRN]", "'" + GetSubStr(sCols[3], 16) + "'");
                        else
                            sql = sql.Replace("[DOC_DRN]", "NULL");

                        if (sCols != null && sCols.Length > 4)
                            sql = sql.Replace("[DOC_UPLOAD_DATE]", "to_date('" + GetBackEndDateTime(sCols[4]) + "','DD-MON-YYYY HH12:MI:SS')");
                        else
                            sql = sql.Replace("[DOC_UPLOAD_DATE]", "NULL");

                        if (sCols != null && sCols.Length > 5)
                            sql = sql.Replace("[DOC_TYPE_CODE]", "'" + GetSubStr(sCols[5], 6) + "'");
                        else
                            sql = sql.Replace("[DOC_TYPE_CODE]", "NULL");

                        sql = sql.Replace("[DOC_FILE_TYPE]", "'" + "PDF" + "'");

                        if (sCols != null && sCols.Length > 6)
                            sql = sql.Replace("[DOC_NAME]", "'" + GetSubStr(sCols[6], 200) + "'");
                        else
                            sql = sql.Replace("[DOC_NAME]", "NULL");

                        sql = sql.Replace("[REC_CREATED_BY]", "'" + user_code + "'");
                        sql = sql.Replace("[REC_CREATED_DATE]", "sysdate");

                        sql = sql.Replace("[DOC_MAIL_ID]", "'" + GetSubStr(MsgEmailID, 150) + "'");

                        if (DocIRN.Trim() != "")
                        {
                            SQL2 = "select doc_pkid from edisupdocs where rec_company_code ='" + company_code + "' and rec_branch_code ='" + branch_code + "' and doc_irn ='" + DocIRN + "'";
                            if (!Con_Oracle.IsRowExists(SQL2))
                            {
                                Con_Oracle.BeginTransaction();
                                Con_Oracle.ExecuteNonQuery(sql);
                                Con_Oracle.CommitTransaction();
                            }
                        }
                    }
                    catch (Exception Ex)
                    {
                        bRet = false;
                        if (Con_Oracle != null)
                        {
                            Con_Oracle.RollbackTransaction();
                            //  Con_Oracle.CloseConnection();
                            Con_Oracle.CreateErrorLog("EsanchitDownload" + Ex.Message.ToString());
                        }
                       // throw Ex;
                    }
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
            return str.ToUpper();
        }

        private string GetBackEndDateTime(string sDate)
        {
            int d = 0, m = 0, y = 0, hh = 0, mm = 0, ss = 0;
            string AMPM = "";
            if (sDate.Length < 8)
                return null;
            string[] strData = null;
            DateTime NewDt = DateTime.Now;
            try
            {
                if (sDate.Contains("-"))
                {
                    string[] sdata = sDate.Split('-'); //2019 - 09 - 16 16:04:05

                    if (sdata.Length > 0)
                        y = Lib.Conv2Integer(sdata[0].Trim());
                    if (sdata.Length > 1)
                        m = Lib.Conv2Integer(sdata[1].Trim());
                    if (sdata.Length > 2)
                    {
                        strData = sdata[2].Trim().Split(' ');
                        if (strData.Length > 0)
                            d = Lib.Conv2Integer(strData[0].Trim());
                        if (strData.Length > 1)
                        {
                            sdata = strData[1].Split(':');
                            hh = Lib.Conv2Integer(sdata[0]);
                            mm = Lib.Conv2Integer(sdata[1]);
                            ss = Lib.Conv2Integer(sdata[2]);
                        }
                    }
                }
                else
                {
                    string[] sdata = sDate.Split(' ');  

                    if (sdata.Length > 0)
                    {
                        strData = sdata[0].Split('/');
                        m = Lib.Conv2Integer(strData[0]);
                        d = Lib.Conv2Integer(strData[1]);
                        y = Lib.Conv2Integer(strData[2]);
                    }
                    if (sdata.Length > 1)
                    {
                        strData = sdata[1].Split(':');
                        hh = Lib.Conv2Integer(strData[0]);
                        mm = Lib.Conv2Integer(strData[1]);
                        ss = Lib.Conv2Integer(strData[2]);
                    }
                    if (sdata.Length > 2)
                        AMPM = sdata[2];
                }

                NewDt = new DateTime(y, m, d, hh, mm, ss);

            }
            catch (Exception)
            {
                return null;
            }
            return NewDt.ToString("dd-MMM-yyyy hh:mm:ss").ToUpper();
        }


        public IDictionary<string, object> LinkList(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            ErrorMessage = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Esanchit> mList = new List<Esanchit>();
            Esanchit mRow;
            try
            {

                type = SearchData["type"].ToString();
                linkid = SearchData["linkid"].ToString();
                jobid = SearchData["jobid"].ToString();
                jobno = SearchData["jobno"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                from_date = SearchData["from_date"].ToString();

                from_date = Lib.StringToDate(from_date);
                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                sWhere = " where a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.rec_branch_code = '{BRCODE}'";
                sWhere += " and (";
                sWhere += " a.doc_job_id = '{JOBID}'";
                //if (searchstring != "")
                //{
                //    sWhere += " or ";
                //    sWhere += "  upper(a.doc_drn) like '%" + searchstring.ToUpper() + "%'";
                //    sWhere += " or ";
                //    sWhere += "  upper(a.doc_irn) like '%" + searchstring.ToUpper() + "%'";
                //}
                sWhere += " )";
                //sWhere += " and ((doc_link_id is null ) or doc_link_id = '{LINKID}' )";

                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);
                sWhere = sWhere.Replace("{JOBID}", jobid);
                sWhere = sWhere.Replace("{LINKID}", linkid);


                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select doc_pkid,doc_drn,doc_irn,doc_type_code,doc_name,doc_file_name,doc_upload_date  ";
                sql += " ,doc_link_id,doc_link_type,b.job_pkid as doc_job_id,b.job_docno as doc_job_no ";
                sql += " ,c.param_name as doc_type_name ";
                sql += " from edisupdocs a ";
                sql += " left join jobm b on a.doc_job_id = b.job_pkid ";
                sql += " left join param c on (a.doc_type_code = c.param_code and c.param_type = 'ESANCHITDOC')";
                sql += sWhere;
                sql += " order by doc_upload_date,doc_drn,doc_irn";

                Dt_List = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Esanchit();
                    mRow.doc_pkid = Dr["doc_pkid"].ToString();
                    mRow.doc_upload_date = Lib.DatetoStringDisplayformat(Dr["doc_upload_date"]);
                    mRow.doc_drn = Dr["doc_drn"].ToString();
                    mRow.doc_irn = Dr["doc_irn"].ToString();
                    mRow.doc_type_code = Dr["doc_type_code"].ToString();
                    mRow.doc_name = Dr["doc_type_name"].ToString();
                    if (Dr["doc_type_name"].ToString().Length > 50)
                        mRow.doc_name = Dr["doc_type_name"].ToString().Substring(0, 50);
                    mRow.doc_file_name = Dr["doc_file_name"].ToString();
                    if (Dr["doc_file_name"].ToString().Length > 50)
                        mRow.doc_file_name = Dr["doc_file_name"].ToString().Substring(0, 50);
                    mRow.doc_job_id = Dr["doc_job_id"].ToString();
                    mRow.doc_job_no = Dr["doc_job_no"].ToString();
                    mRow.doc_link_type = Dr["doc_link_type"].ToString();
                    mRow.doc_selected = false;
                    //if (Dr["doc_link_id"].ToString() == linkid)
                    //    mRow.doc_selected = true;
                    mList.Add(mRow);
                }
                Dt_List.Rows.Clear();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("list", mList);
            return RetData;
        }
        public Dictionary<string, object> SaveLink(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            savemsg = "";
            DataTable Dt_job = new DataTable();
            Con_Oracle = new DBConnection();
            try
            {
                jobno = "";
                pkid = SearchData["pkid"].ToString();
                linktype = SearchData["linktype"].ToString();
                linkid = SearchData["linkid"].ToString();
                jobid = SearchData["jobid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();


                if (pkid.Contains(","))
                    pkid = pkid.Replace(",", "','");

                //if (jobid != "")
                //{
                //    sql = "select job_docno,job_year,rec_category from jobm where job_pkid='" + jobid + "'";
                //    Dt_job = Con_Oracle.ExecuteQuery(sql);
                //    if (Dt_job.Rows.Count > 0)
                //        jobno = Dt_job.Rows[0]["job_docno"].ToString();
                //}

                Con_Oracle.BeginTransaction();
                //sql = "update edisupdocs set doc_link_id = null,doc_link_type = null where doc_link_id = '" + linkid + "'";
                //Con_Oracle.ExecuteNonQuery(sql);

                if (pkid.Trim() != "")
                {
                    sql = "update edisupdocs set doc_link_id = '" + linkid + "'";
                    sql += " ,doc_link_type ='" + linktype + "'";
                    /*
                    sql += " ,doc_job_id ='" + jobid + "'";
                    if (Dt_job.Rows.Count > 0)
                    {
                        sql += " ,doc_job_no = " + Lib.Conv2Integer(Dt_job.Rows[0]["job_docno"].ToString());
                        sql += " ,doc_job_year = " + Lib.Conv2Integer(Dt_job.Rows[0]["job_year"].ToString());
                        sql += " ,rec_category = '" + Dt_job.Rows[0]["rec_category"].ToString() + "'";
                    }
                    */
                    sql += " where doc_pkid in ('" + pkid + "')";
                    Con_Oracle.ExecuteNonQuery(sql);
                }

                Con_Oracle.CommitTransaction();
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
            RetData.Add("jobno", jobno);
            RetData.Add("savemsg", savemsg);
            return RetData;
        }

        private string SavePasteData(string cbdata)
        {
            string[] Arry = null;
            string[] ArryData = null;

            int col_fname = -1;
            int col_drn = -1;
            int col_irn = -1;
            int col_update = -1;
            int col_type = -1;
            string sMsgError = "";

            if (cbdata == null || cbdata.Trim() == "")
            {
                throw new Exception("No Data Found.....[PASTEDATA]");
            }

            Arry = cbdata.Trim().Split('\n');
            if (Arry.Length > 0)
            {
                ArryData = Arry[0].Split('\t');
                if (ArryData.Length < 5)
                {
                    throw new Exception("Please select all Columns and continue.....[PASTEDATA]");
                }

                for (int i = 0; i < ArryData.Length; i++)
                {
                    if (ArryData[i].ToUpper().Trim() == "FILE NAME")
                        col_fname = i;
                    else if (ArryData[i].ToUpper().Trim() == "DRN")
                        col_drn = i;
                    else if (ArryData[i].ToUpper().Trim() == "IRN")
                        col_irn = i;
                    else if (ArryData[i].ToUpper().Trim() == "UPLOAD DATE")
                        col_update = i;
                    else if (ArryData[i].ToUpper().Trim() == "DOCUMENT TYPE")
                        col_type = i;
                }

                if (col_fname == -1 || col_drn == -1 || col_irn == -1 || col_update == -1 || col_type == -1)
                {
                    throw new Exception("Column Heading Not selected or Some Columns Missing or Column Name Changed.....[PASTEDATA]");
                }

                if (Arry.Length > 1)
                {
                    ArryData = Arry[1].Split('\t');//this method for ie and firefox browsers
                    if (ArryData.Length >= 5)
                    {
                        for (int i = 1; i < Arry.Length; i++)
                        {
                            if (Arry[i] != null)
                            {
                                ArryData = Arry[i].Split('\t');
                                if (ArryData != null)
                                    if (ArryData.Length >= 5)
                                    {
                                        InsertData(ArryData[col_fname], ArryData[col_drn], ArryData[col_irn], ArryData[col_update], ArryData[col_type]);
                                    }
                            }
                        }

                    }
                    else //this is for chrome browser
                    {

                        if ((Arry.Length - 1) % 5 != 0)
                        {
                            throw new Exception("Some Columns Data not Selected or Missing......[PASTEDATA]");
                        }

                        string sLine = "";
                        for (int i = 1; i < Arry.Length; i++)
                        {
                            if (Arry[i] != null)
                            {
                                if (sLine != "")
                                    sLine += "\t";

                                sLine += Arry[i];
                            }

                            if (i % 5 == 0)
                            {
                                ArryData = sLine.Split('\t');
                                if (ArryData != null)
                                    if (ArryData.Length >= 5)
                                    {
                                        InsertData(ArryData[col_fname], ArryData[col_drn], ArryData[col_irn], ArryData[col_update], ArryData[col_type]);
                                    }
                                sLine = "";
                            }
                        }
                    }
                }
            }

            return sMsgError;
        }

        private string OLD_SavePasteData(string cbdata)
        {
            string[] Arry = null;
            string[] ArryData = null;

            int col_fname = -1;
            int col_drn = -1;
            int col_irn = -1;
            int col_update = -1;
            int col_type = -1;
            string sMsgError = "";

            if (cbdata == null || cbdata.Trim() == "")
            {
                throw new Exception("No Data Found.....[PASTEDATA]");
            }

            Arry = cbdata.Split('\n');
            if (Arry.Length > 0)
            {
                ArryData = Arry[0].Split('\t');
                if (ArryData.Length < 5)
                {
                    throw new Exception("Please select all Columns and continue.....[PASTEDATA]");
                }

                for (int i = 0; i < ArryData.Length; i++)
                {
                    if (ArryData[i].ToUpper().Trim() == "FILE NAME")
                        col_fname = i;
                    else if (ArryData[i].ToUpper().Trim() == "DRN")
                        col_drn = i;
                    else if (ArryData[i].ToUpper().Trim() == "IRN")
                        col_irn = i;
                    else if (ArryData[i].ToUpper().Trim() == "UPLOAD DATE")
                        col_update = i;
                    else if (ArryData[i].ToUpper().Trim() == "DOCUMENT TYPE")
                        col_type = i;
                }

                if (col_fname == -1 || col_drn == -1 || col_irn == -1 || col_update == -1 || col_type == -1)
                {
                    throw new Exception("Column Heading Not selected or Some Columns Missing or Column Name Changed.....[PASTEDATA]");
                }

                for (int i = 1; i < Arry.Length; i++)
                {
                    if (Arry[i] != null)
                    {
                        ArryData = Arry[i].Split('\t');
                        if (ArryData != null)
                            if (ArryData.Length >= 5)
                            {
                                InsertData(ArryData[col_fname], ArryData[col_drn], ArryData[col_irn], ArryData[col_update], ArryData[col_type]);
                            }
                    }
                }
            }

            return sMsgError;
        }
        private bool InsertData(string doc_FileName, string doc_drn, string doc_irn, string doc_uploaddate, string doc_type)
        {
            string SQL = "";
            bool bRet = false;
            string[] sdata = null;
            string DocIRN = "";
            Con_Oracle = new DBConnection();
            string doc_typecode = "";
            try
            {
                sdata = doc_type.Split('-');
                if (sdata.Length > 0)
                    doc_typecode = sdata[0];
                doc_type = doc_type.Replace(doc_typecode + "-", "");

                SQL = "select doc_pkid from edisupdocs where rec_company_code ='" + company_code + "' and rec_branch_code ='" + branch_code + "' and doc_irn ='" + doc_irn + "'";
                if (!Con_Oracle.IsRowExists(SQL))
                {

                    sql = "INSERT INTO EDISUPDOCS ( ";
                    sql += " DOC_PKID,REC_COMPANY_CODE,REC_BRANCH_CODE ,REC_CATEGORY,";
                    sql += " DOC_DRN,DOC_IRN,DOC_TYPE_CODE,DOC_FILE_TYPE,";
                    sql += " DOC_FILE_NAME,DOC_NAME,";
                    sql += " REC_CREATED_BY,REC_CREATED_DATE,DOC_UPLOAD_DATE,DOC_MAIL_ID";
                    sql += ") values (";
                    sql += " [DOC_PKID],[REC_COMPANY_CODE],[REC_BRANCH_CODE],[REC_CATEGORY],";
                    sql += " [DOC_DRN],[DOC_IRN],[DOC_TYPE_CODE],[DOC_FILE_TYPE],";
                    sql += " [DOC_FILE_NAME],[DOC_NAME], ";
                    sql += " [REC_CREATED_BY],[REC_CREATED_DATE],[DOC_UPLOAD_DATE],[DOC_MAIL_ID]";
                    sql += " )";

                    sql = sql.Replace("[DOC_PKID]", "'" + Guid.NewGuid().ToString().ToUpper() + "'");
                    sql = sql.Replace("[REC_COMPANY_CODE]", "'" + company_code + "'");
                    sql = sql.Replace("[REC_BRANCH_CODE]", "'" + branch_code + "'");
                    sql = sql.Replace("[REC_CATEGORY]", "NULL");

                    if (doc_FileName != null && doc_FileName.Length > 1)
                        sql = sql.Replace("[DOC_FILE_NAME]", "'" + GetSubStr(doc_FileName, 200) + "'");
                    else
                        sql = sql.Replace("[DOC_FILE_NAME]", "NULL");

                    if (doc_irn != null && doc_irn.Length > 1)
                    {
                        DocIRN = GetSubStr(doc_irn, 16);
                        sql = sql.Replace("[DOC_IRN]", "'" + DocIRN + "'");
                    }
                    else
                        sql = sql.Replace("[DOC_IRN]", "NULL");

                    if (doc_drn != null && doc_drn.Length > 1)
                        sql = sql.Replace("[DOC_DRN]", "'" + GetSubStr(doc_drn, 16) + "'");
                    else
                        sql = sql.Replace("[DOC_DRN]", "NULL");

                    if (doc_uploaddate != null && doc_uploaddate.Length > 1)
                        sql = sql.Replace("[DOC_UPLOAD_DATE]", "to_date('" + GetBackEndDateTime(doc_uploaddate) + "','DD-MON-YYYY HH12:MI:SS')");
                    else
                        sql = sql.Replace("[DOC_UPLOAD_DATE]", "NULL");

                    if (doc_typecode != null && doc_typecode.Length > 1)
                        sql = sql.Replace("[DOC_TYPE_CODE]", "'" + GetSubStr(doc_typecode, 6) + "'");
                    else
                        sql = sql.Replace("[DOC_TYPE_CODE]", "NULL");

                    sql = sql.Replace("[DOC_FILE_TYPE]", "'" + "PDF" + "'");

                    if (doc_type != null && doc_type.Length > 1)
                        sql = sql.Replace("[DOC_NAME]", "'" + GetSubStr(doc_type, 200) + "'");
                    else
                        sql = sql.Replace("[DOC_NAME]", "NULL");

                    sql = sql.Replace("[REC_CREATED_BY]", "'" + user_code + "'");
                    sql = sql.Replace("[REC_CREATED_DATE]", "sysdate");

                    sql = sql.Replace("[DOC_MAIL_ID]", "'PASTEDATA'");

                    if (DocIRN.Trim() != "")
                    {
                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                    }
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

            bRet = true;
            return bRet;
        }
    }
}

