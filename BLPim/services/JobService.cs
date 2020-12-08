using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase.Connections;

namespace BLPim
{
    public class JobService : BL_Base
    {

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<pim_spot> mList = new List<pim_spot>();
            pim_spot mRow;

            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string comp_code = SearchData["comp_code"].ToString();
            string type = SearchData["type"].ToString();

            string user_id = SearchData["user_id"].ToString();
            Boolean user_admin = (Boolean)SearchData["user_admin"];

            string region_id = SearchData["region_id"].ToString();
            string vendor_id = SearchData["vendor_id"].ToString();
            string role_name = SearchData["role_name"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            string sql2 = "";

            try
            {
                sWhere = " where  a.rec_company_code ='" + comp_code + "' ";

                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  comp_name like '%" + searchstring.ToUpper() + "%' ";
                    sWhere += " )";
                }

                if (role_name == "SUPER ADMIN")
                    sql2 = "";
                else if (role_name == "ZONE ADMIN")
                    sWhere += " and store.comp_region_id = '" + region_id + "'";
                else if (role_name == "SALES EXECUTIVE")
                    sql2 = " inner join userd e on e.rec_type = 'S' and a.spot_store_id = e.user_branch_id and e.user_id ='" + user_id + "'";
                else if (role_name == "VENDOR")
                    sWhere += " and spot_vendor_id = '" + vendor_id + "'";
                else if (role_name == "RECCE USER")
                    sWhere += " and spot_vendor_id = '" + vendor_id + "'";


                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  ";
                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total ";
                    sql += " FROM pim_spotm a ";
                    sql += " inner join companym store on a.spot_store_id = store.comp_pkid ";

                    if (!user_admin)
                        sql += sql2;

                    sql += sWhere;
                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        page_rowcount = Lib.Conv2Integer(Dt_Temp.Rows[0]["total"].ToString());
                        page_count = Lib.Conv2Integer(Dt_Temp.Rows[0]["page_total"].ToString());
                    }
                    page_current = 1;
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
                sql += " select  spot_pkid, spot_slno, spot_date,  a.rec_created_by, a.rec_created_date, ";
                sql += " spot_store_id, store.comp_name as spot_store_name, ";
                sql += " spot_vendor_id,vendor.comp_name as spot_vendor_name, ";
                sql += " region.param_name as spot_region_name, ";
                sql += " spot_job_remarks,";
                sql += " am_by, am_date, am_status, am_remarks,";
                sql += " row_number() over(order by store.comp_name) rn ";
                sql += " from  pim_spotm a  ";
                sql += " inner join companym store on a.spot_store_id = store.comp_pkid ";
                sql += " left  join param region on store.comp_region_id = region.param_pkid ";
                sql += " inner join companym vendor on a.spot_vendor_id = vendor.comp_pkid ";
                sql += " left join approvalm on spot_pkid =  am_pkid ";

                if (!user_admin)
                    sql += sql2;

                sql += " " + sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by a.rec_created_date,spot_slno ";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();


                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new pim_spot();
                    mRow.spot_pkid = Dr["spot_pkid"].ToString();

                    mRow.spot_slno = Lib.Conv2Integer(Dr["spot_slno"].ToString());
                    mRow.spot_date  = Lib.DatetoStringDisplayformat(Dr["spot_date"]);

                    mRow.spot_store_name = Dr["spot_store_name"].ToString();
                    mRow.spot_vendor_name = Dr["spot_vendor_name"].ToString();
                    mRow.spot_region_name = Dr["spot_region_name"].ToString();

                    mRow.spot_job_remarks = Dr["spot_job_remarks"].ToString();

                    mRow.approved_by = Dr["am_by"].ToString();
                    mRow.approved_status = Dr["am_status"].ToString();
                    mRow.approved_remarks = Dr["am_remarks"].ToString();
                    mRow.approved_date = Lib.DatetoStringDisplayformat(Dr["am_date"]);

                    mRow.rec_created_by = Dr["rec_created_by"].ToString();
                    mRow.rec_created_date = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);


                    mList.Add(mRow);
                }

                Dt_List.Rows.Clear();
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

            return RetData;
        }

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            List<pim_spotd> mList = new List<pim_spotd>();

            pim_spot mRow = new pim_spot();
            pim_spotd mRowd = new pim_spotd();

            string id = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            
            string ServerImageUrl = Lib.GetSeverImageURL(comp_code);

            string sql2 = "";

            try
            {
                DataTable Dt_Rec = new DataTable();
                DataTable Dt_RecDet = new DataTable();

                sql = "select  spot_pkid, spot_slno, spot_date, ";
                sql += " spot_store_id, store.comp_name as spot_store_name, ";
                sql += " spot_vendor_id, vendor.comp_name as spot_vendor_name, ";
                sql += " spot_job_remarks, spot_store_contact_name, spot_store_contact_tel,am_status ";
                sql += " from pim_spotm a  ";
                sql += " left join approvalm on spot_pkid =  am_pkid";
                sql += " left join companym store on a.spot_store_id = store.comp_pkid ";
                sql += " left join companym vendor on a.spot_vendor_id = vendor.comp_pkid ";
                sql += " where spot_pkid = '" + id + "'";


                sql2 = "select  spotd_pkid, spotd_parent_id, spotd_name ,spotd_slno,spotd_uom, spotd_wd, spotd_ht, ";
                sql2 += " spotd_artwork_id, artwork.param_name as spotd_artwork_name, artwork.param_slno as spotd_artwork_slno,artwork.param_file_name as spotd_artwork_file_name,";
                sql2 += " spotd_product_id, product.param_name as spotd_product_name, ";
                sql2 += " spotd_close_view, spotd_long_view, spotd_final_view, spotd_status ";
                sql2 += " from pim_spotd a  ";
                sql2 += " left join param artwork on a.spotd_artwork_id = artwork.param_pkid ";
                sql2 += " left join param product on a.spotd_product_id = product.param_pkid ";
                sql2 += " where spotd_parent_id = '" + id + "'";
                sql2 += " order by spotd_slno";


                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Dt_RecDet = Con_Oracle.ExecuteQuery(sql2);
                Con_Oracle.CloseConnection();


                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new pim_spot();
                    mRow.spot_pkid = Dr["spot_pkid"].ToString();

                    if (Dr["spot_date"].Equals(DBNull.Value))
                        mRow.spot_date = "";
                    else
                        mRow.spot_date = Lib.DatetoString(Dr["spot_date"]);

                    mRow.spot_slno =  Lib.Conv2Integer(Dr["spot_slno"].ToString());

                    mRow.spot_store_id = Dr["spot_store_id"].ToString();
                    mRow.spot_store_name = Dr["spot_store_name"].ToString();

                    mRow.spot_vendor_id = Dr["spot_vendor_id"].ToString();
                    mRow.spot_vendor_name = Dr["spot_vendor_name"].ToString();

                    mRow.spot_job_remarks = Dr["spot_job_remarks"].ToString();

                    mRow.approved_status = Dr["am_status"].ToString();

                    mRow.spot_store_contact_name = Dr["spot_store_contact_name"].ToString();
                    mRow.spot_store_contact_tel = Dr["spot_store_contact_tel"].ToString();


                    break;
                }

                foreach (DataRow Dr in Dt_RecDet.Rows)
                {
                    mRowd = new pim_spotd();
                    mRowd.spotd_pkid = Dr["spotd_pkid"].ToString();
                    mRowd.spotd_parent_id = Dr["spotd_parent_id"].ToString();

                    mRowd.spotd_slno = Lib.Conv2Integer(Dr["spotd_slno"].ToString());
                    mRowd.spotd_name = Dr["spotd_name"].ToString();
                    mRowd.spotd_uom = Dr["spotd_uom"].ToString();
                    mRowd.spotd_wd = Lib.Convert2Decimal(Dr["spotd_wd"].ToString());
                    mRowd.spotd_ht = Lib.Convert2Decimal(Dr["spotd_ht"].ToString());

                    mRowd.spotd_artwork_id = Dr["spotd_artwork_id"].ToString();
                    mRowd.spotd_artwork_name = Dr["spotd_artwork_name"].ToString();
                    mRowd.spotd_artwork_file_name = Dr["spotd_artwork_file_name"].ToString();

                    mRowd.spotd_product_id = Dr["spotd_product_id"].ToString();
                    mRowd.spotd_product_name = Dr["spotd_product_name"].ToString();

                    mRowd.spotd_close_view = Dr["spotd_close_view"].ToString();
                    mRowd.spotd_long_view = Dr["spotd_long_view"].ToString();
                    mRowd.spotd_final_view = Dr["spotd_final_view"].ToString();
 
                    mRowd.spotd_close_view_file_uploaded = false;
                    if (Dr["spotd_close_view"].ToString().Length > 0)
                        mRowd.spotd_close_view_file_uploaded = true;

                    mRowd.spotd_long_view_file_uploaded = false;
                    if (Dr["spotd_long_view"].ToString().Length > 0)
                        mRowd.spotd_long_view_file_uploaded = true;

                    mRowd.spotd_final_view_file_uploaded = false;
                    if (Dr["spotd_final_view"].ToString().Length > 0)
                        mRowd.spotd_final_view_file_uploaded = true;

                    mRowd.spotd_status = Dr["spotd_status"].ToString();

                    mRowd.rec_mode = "EDIT";

                    mList.Add(mRowd);

                }






            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            RetData.Add("recorddet", mList);
            return RetData;
        }

        public string AllValid( pim_spot Record)
        {
            string str = "";
            try
            {
                /*

                sql = "select comp_pkid from (";
                sql += "select comp_pkid  from companym a where (a.comp_code = '{CODE}')  ";

                sql += ") a where comp_pkid <> '{PKID}'";


                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";
                */

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(pim_spot Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            Boolean retvalue = false;

            int iSlno = 0;

            try
            {
                Con_Oracle = new DBConnection();

                if (Record.spot_date.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Date Cannot Be Empty");

                if (Record.spot_store_id.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Store Cannot Be Empty");

                if (Record.spot_vendor_id.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Vendor Cannot Be Empty");


                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                if (Record.rec_mode == "ADD")
                {
                    sql = "select isnull(max(spot_slno), 1) + 1  as slno from pim_spotm where ";
                    sql += " rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    iSlno = Lib.Conv2Integer(Con_Oracle.ExecuteScalar(sql).ToString());
                }
                else
                    iSlno = Record.spot_slno;

                if (iSlno <= 0)
                {
                    throw new Exception("Invalid CF#");
                }


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("pim_spotm", Record.rec_mode, "spot_pkid", Record.spot_pkid);
                Rec.InsertDate("spot_date", Record.spot_date);
                Rec.InsertString("spot_store_id", Record.spot_store_id);

                Rec.InsertString("spot_vendor_id", Record.spot_vendor_id);

                Rec.InsertString("spot_job_remarks", Record.spot_job_remarks);
                

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("spot_slno", iSlno.ToString());
                    Rec.InsertString("spot_executive_name", Record._globalvariables.user_name);

                    Rec.InsertString("spot_store_contact_name", Record.spot_store_contact_name);
                    Rec.InsertString("spot_store_contact_tel", Record.spot_store_contact_tel);

                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    if (Con_Oracle.DB == "ORACLE")
                        Rec.InsertFunction("rec_created_date", "SYSDATE");
                    else
                        Rec.InsertFunction("rec_created_date", "GETDATE()");

                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    if (Con_Oracle.DB == "ORACLE")
                        Rec.InsertFunction("rec_edited_date", "SYSDATE");
                    else
                        Rec.InsertFunction("rec_edited_date", "GETDATE()");
                }

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
                retvalue = true;
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                retvalue = false;
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("retvalue",retvalue);
            RetData.Add("slno", iSlno);

            return RetData;
        }


        public Dictionary<string, object> SaveStatus(pim_spotd Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Boolean retvalue = false;
            try
            {
                Con_Oracle = new DBConnection();
                DBRecord Rec = new DBRecord();
                Rec.CreateRow("pim_spotd", "EDIT", "spotd_pkid", Record.spotd_pkid);
                Rec.InsertString("spotd_status", Record.spotd_status);
                sql = Rec.UpdateRow();
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
                retvalue = true;
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                retvalue = false;
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("retvalue", retvalue);

            return RetData;
        }




        public Dictionary<string, object> Delete(Dictionary<string, object> SearchData)
        {
            Boolean bRet = false;
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            string pkid = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();


            Lib.RemoveFileFolder(comp_code, pkid);

            DataTable dt_test = new DataTable();
            sql = "select spotd_pkid from pim_spotd where spotd_parent_id = '" + pkid + "'";
            dt_test = Con_Oracle.ExecuteQuery(sql);
            foreach ( DataRow Dr in dt_test.Rows)
            {
                Lib.RemoveFileFolder(comp_code, Dr["spotd_pkid"].ToString());
            }

            try
            {
                Con_Oracle.BeginTransaction();
                sql = "delete from pim_spotd where spotd_parent_id = '" + pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                sql = "delete from pim_spotm where spot_pkid = '" + pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
                bRet = true;
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                bRet = false;
                throw Ex;
            }
            RetData.Add("retvalue", bRet);
            return RetData;
        }






        // Recce User

        public Dictionary<string, object> GetRecord_recce_user(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            List<pim_spotd> mList = new List<pim_spotd>();

            pim_spot mRow = new pim_spot();

            string id = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();


            try
            {
                DataTable Dt_Rec = new DataTable();
                DataTable Dt_RecDet = new DataTable();

                sql = "select  spot_pkid, ";
                sql += " spot_recce_id, b.user_name ";
                sql += " from pim_spotm a  ";
                sql += " left join userm b on a.spot_recce_id =  b.user_pkid";
                sql += " where spot_pkid = '" + id + "'";


                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();


                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new pim_spot();
                    mRow.spot_pkid = Dr["spot_pkid"].ToString();

                    mRow.spot_recce_id = Dr["spot_recce_id"].ToString();
                    mRow.spot_recce_name = Dr["user_name"].ToString();

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



        public Dictionary<string, object> Save_recce_user(pim_spot Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            Boolean retvalue = false;

            try
            {
                Con_Oracle = new DBConnection();

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("pim_spotm", "EDIT", "spot_pkid", Record.spot_pkid);
                Rec.InsertString("spot_recce_id", Record.spot_recce_id);

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
                retvalue = true;
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                retvalue = false;
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("retvalue", retvalue);

            return RetData;
        }



    }
}
