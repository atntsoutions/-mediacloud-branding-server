using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DataBase;
using DataBase_Oracle.Connections;

namespace BLOperations
{
    public class TrackOrderService : BL_Base
    {
        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Joborderm mRow = new Joborderm();

            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select ord_pkid,ord_booking_date,ord_rnd_insp_date,ord_po_rel_date,";
                sql += " ord_cargo_ready_date,ord_fcr_date,ord_insp_date, ";
                sql += " ord_stuf_date,ord_whd_date,ord_dlv_pol_date,ord_dlv_pod_date,ord_cargo_status  ";
                sql += " from joborderm a ";
                sql += " where  a.ord_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Joborderm();
                    mRow.ord_pkid = Dr["ord_pkid"].ToString();
                    mRow.ord_booking_date = Lib.DatetoString(Dr["ord_booking_date"]);
                    mRow.ord_rnd_insp_date = Lib.DatetoString(Dr["ord_rnd_insp_date"]);
                    mRow.ord_po_rel_date = Lib.DatetoString(Dr["ord_po_rel_date"]);
                    mRow.ord_cargo_ready_date = Lib.DatetoString(Dr["ord_cargo_ready_date"]);
                    mRow.ord_fcr_date = Lib.DatetoString(Dr["ord_fcr_date"]);
                    mRow.ord_insp_date = Lib.DatetoString(Dr["ord_insp_date"]);
                    mRow.ord_stuf_date = Lib.DatetoString(Dr["ord_stuf_date"]);
                    mRow.ord_whd_date = Lib.DatetoString(Dr["ord_whd_date"]);
                    mRow.ord_dlv_pol_date = Lib.DatetoString(Dr["ord_dlv_pol_date"]);
                    mRow.ord_dlv_pod_date = Lib.DatetoString(Dr["ord_dlv_pod_date"]);
                    mRow.ord_cargo_status = Dr["ord_cargo_status"].ToString();
                    break;
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("record", mRow);
            return RetData;
        }


        public string AllValid(Joborderm Record)
        {
            string str = "";
            try
            {
                /*
                if (Record.ord_po.Length <= 0)
                    Lib.AddError(ref str, " PO No Cannot Be Blank");

                sql = "";
                sql = "select ord_pkid from (";
                sql += " select ord_pkid from joborderm ";
                sql += " where rec_company_code = '{COMPANY_CODE}' ";
                sql += " and rec_branch_code = '{BRANCH_CODE}'";
                sql += " and ord_exp_id = '{EXP_ID}' ";
                sql += " and ord_imp_id = '{IMP_ID}' ";
                sql += " and ord_po = '{PO}' ";

                if (Record.ord_style != "")
                    sql += " and ord_style = '{STYLE}' ";
                else
                    sql += " and ord_style  is null ";

                if (Record.ord_color != "")
                    sql += " and ord_color = '{COLOR}' ";
                else
                    sql += " and ord_color is null";

                sql += ") a where ord_pkid <> '{PKID}'";
                
                sql = sql.Replace("{COMPANY_CODE}", Record._globalvariables.comp_code);
                sql = sql.Replace("{BRANCH_CODE}", Record._globalvariables.branch_code);
                sql = sql.Replace("{EXP_ID}", Record.ord_exp_id);
                sql = sql.Replace("{IMP_ID}", Record.ord_imp_id);
                sql = sql.Replace("{PO}", Record.ord_po);
                sql = sql.Replace("{STYLE}", Record.ord_style);
                sql = sql.Replace("{COLOR}", Record.ord_color);
                sql = sql.Replace("{PKID}", Record.ord_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " | This PO No Already Exists");
            */
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


      

        public Dictionary<string, object> Update(Joborderm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            string strDate = "";
            if (Record.ord_pkid.Contains(",") == false)//Single tracking data save
            {
                RetData = Save(Record);
            }
            else //Multiple tracking data save
            {
                try
                {
                    Con_Oracle = new DBConnection();
                    Record.ord_pkid = Record.ord_pkid.Replace(" ", "");
                    Record.ord_pkid = Record.ord_pkid.Replace(",", "','");

                    sql = "";
                    strDate = Lib.StringToDate(Record.ord_booking_date);
                    if (strDate != "NULL")
                        sql += " ord_booking_date ='" + strDate + "'";

                    strDate = Lib.StringToDate(Record.ord_rnd_insp_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_rnd_insp_date ='" + strDate + "'";
                    }

                    strDate = Lib.StringToDate(Record.ord_po_rel_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_po_rel_date ='" + strDate + "'";
                    }

                    strDate = Lib.StringToDate(Record.ord_cargo_ready_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_cargo_ready_date ='" + strDate + "'";
                    }

                    strDate = Lib.StringToDate(Record.ord_fcr_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_fcr_date ='" + strDate + "'";
                    }

                    strDate = Lib.StringToDate(Record.ord_insp_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_insp_date ='" + strDate + "'";
                    }

                    strDate = Lib.StringToDate(Record.ord_stuf_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_stuf_date ='" + strDate + "'";
                    }

                    strDate = Lib.StringToDate(Record.ord_whd_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_whd_date ='" + strDate + "'";
                    }
                    
                    strDate = Lib.StringToDate(Record.ord_dlv_pol_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_dlv_pol_date ='" + strDate + "'";
                    }
                    strDate = Lib.StringToDate(Record.ord_dlv_pod_date);
                    if (strDate != "NULL")
                    {
                        if (sql != "")
                            sql += ",";
                        sql += " ord_dlv_pod_date ='" + strDate + "'";
                    }

                    if (sql != "")
                    {
                        string sql2 = "";
                        sql2 = "update joborderm set " + sql;
                        sql2 += "  where ord_pkid in ('" + Record.ord_pkid + "')";
                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteNonQuery(sql2);
                        Con_Oracle.CommitTransaction();

                        sql = "update joborderm set ";
                        sql += " ord_track_status =";
                        sql += " case when nvl(length(ord_booking_date),0)>0 then 'BKD,' else '' end||";
                        sql += " case when nvl(length(ord_rnd_insp_date),0)>0 then 'RND,'  else '' end||";
                        sql += " case when nvl(length(ord_po_rel_date),0)>0 then 'POR,'  else '' end||";
                        sql += " case when nvl(length(ord_cargo_ready_date),0)>0 then 'CR,'  else '' end||";
                        sql += " case when nvl(length(ord_fcr_date),0)>0 then 'FCR,'  else '' end||";
                        sql += " case when nvl(length(ord_insp_date),0)>0 then 'INSP,'  else '' end||";
                        sql += " case when nvl(length(ord_stuf_date),0)>0 then 'STF,'  else '' end||";
                        sql += " case when nvl(length(ord_whd_date),0)>0 then 'WHD,'  else '' end||";
                        sql += " case when nvl(length(ord_dlv_pol_date),0)>0 then 'DLVPOL,' else '' end||  ";
                        sql += " case when nvl(length(ord_dlv_pod_date),0)>0 then 'DLVPOD'  else '' end  ";
                        sql += "  where ord_pkid in ('" + Record.ord_pkid + "')";
                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                    }

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
            }

            return RetData;
        }

        public Dictionary<string, object> Save(Joborderm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if ((ErrorMessage = AllValid(Record)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("joborderm", "EDIT", "ord_pkid", Record.ord_pkid);
                Rec.InsertDate("ord_booking_date", Record.ord_booking_date);
                Rec.InsertDate("ord_rnd_insp_date", Record.ord_rnd_insp_date);
                Rec.InsertDate("ord_po_rel_date", Record.ord_po_rel_date);
                Rec.InsertDate("ord_cargo_ready_date", Record.ord_cargo_ready_date);
                Rec.InsertDate("ord_fcr_date", Record.ord_fcr_date);
                Rec.InsertDate("ord_insp_date", Record.ord_insp_date);
                Rec.InsertDate("ord_stuf_date", Record.ord_stuf_date);
                Rec.InsertDate("ord_whd_date", Record.ord_whd_date);
                Rec.InsertDate("ord_dlv_pol_date", Record.ord_dlv_pol_date);
                Rec.InsertDate("ord_dlv_pod_date", Record.ord_dlv_pod_date);
                Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                Rec.InsertFunction("rec_edited_date", "SYSDATE");

                sql = Rec.UpdateRow();
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();

                sql = "update joborderm set ";
                sql += " ord_track_status =";
                sql += " case when nvl(length(ord_booking_date),0)>0 then 'BKD,' else '' end||";
                sql += " case when nvl(length(ord_rnd_insp_date),0)>0 then 'RND,'  else '' end||";
                sql += " case when nvl(length(ord_po_rel_date),0)>0 then 'POR,'  else '' end||";
                sql += " case when nvl(length(ord_cargo_ready_date),0)>0 then 'CR,'  else '' end||";
                sql += " case when nvl(length(ord_fcr_date),0)>0 then 'FCR,'  else '' end||";
                sql += " case when nvl(length(ord_insp_date),0)>0 then 'INSP,'  else '' end||";
                sql += " case when nvl(length(ord_stuf_date),0)>0 then 'STF,'  else '' end||";
                sql += " case when nvl(length(ord_whd_date),0)>0 then 'WHD,'  else '' end || ";
                sql += " case when nvl(length(ord_dlv_pol_date),0)>0 then 'DLVPOL,' else '' end||  ";
                sql += " case when nvl(length(ord_dlv_pod_date),0)>0 then 'DLVPOD'  else '' end  ";
                sql += "  where ord_pkid ='" + Record.ord_pkid + "'";

                Con_Oracle.BeginTransaction();
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

            return RetData;
        }


        public Dictionary<string, object> ChangeStatus(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string[] pkids = SearchData["pkids"].ToString().Split(',');
            string status = SearchData["status"].ToString();

            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string user_code = SearchData["user_code"].ToString();

            string color = "";

            int iCount = 0;

            string ordno = "";

            List<result> mList = new List<result>();
            result mRow;

            try
            {

                Con_Oracle = new DBConnection();
                

                foreach (string id in pkids)
                {
                    if (status == "APPROVED")
                    {
                        sql = "update joborderm set ord_status = '" + status + "' where ord_pkid ='" + id + "' and nvl(ord_status,'REPORTED') in ('REPORTED','ON HOLD') ";
                        color = "GREEN";
                    }
                    if (status == "ON HOLD") {
                        color = "PURPLE";
                        sql = "update joborderm set ord_status = '" + status + "' where ord_pkid ='" + id + "' and nvl(ord_status,'REPORTED') in ('REPORTED')";
                    }
                    if (status == "CANCELLED") {
                        color = "RED";
                        sql = "update joborderm set ord_status = '" + status + "' where ord_pkid ='" + id + "' and ord_status in ('ON HOLD')";
                    }

                    Con_Oracle.BeginTransaction();
                    iCount  = Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();

                    if (iCount > 0)
                    {
                        mRow = new result();
                        mRow.id = id;
                        mRow.status = status;
                        mRow.color = color;
                        mList.Add(mRow);


                        ordno = Con_Oracle.ExecuteScalar("select max(ord_po) as ord_po from joborderm where ord_pkid  = '" + id + "'" ).ToString();

                        Lib.AuditLog("PO", "PO", "PO STATUS", company_code ,branch_code, user_code, id, ordno ,status );

                    }
                }
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



    }

    public class result
    {
        public string id { get; set; }
        public string status { get; set; }
        public string color { get; set; }
    }

    }

