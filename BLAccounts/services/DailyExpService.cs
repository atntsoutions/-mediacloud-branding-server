using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

using System.Drawing;
using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLAccounts
{
    public class DailyExpService : BL_Base
    {

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<DailyExpm> mList = new List<DailyExpm>();
            DailyExpm mRow;

            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string year_code = SearchData["year_code"].ToString();

            string from_date = "";
            if (SearchData.ContainsKey("from_date"))
                from_date = SearchData["from_date"].ToString();

            string to_date = "";
            if (SearchData.ContainsKey("to_date"))
                to_date = SearchData["to_date"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);

                sWhere = " where a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.rec_branch_code = '{BRCODE}'";
                sWhere += " and a.dem_year =  {FYEAR} ";
                if (searchstring != "")
                {
                    //sWhere += " and (";
                    //sWhere += "  upper(a.hbl_no) like '%" + searchstring.ToUpper() + "%'";
                    //sWhere += " or ";
                    //sWhere += "  upper(a.hbl_bl_no) like '%" + searchstring.ToUpper() + "%'";
                    //sWhere += " or ";
                    //sWhere += "  upper(a.hbl_folder_no) like '%" + searchstring.ToUpper() + "%'";
                    //sWhere += " or ";
                    //sWhere += "  upper(agent1.cust_name) like '%" + searchstring.ToUpper() + "%'";
                    //sWhere += " or ";
                    //sWhere += "  upper(carr.param_name) like '%" + searchstring.ToUpper() + "%'";
                    //sWhere += " )";
                }
                if (from_date != "NULL")
                    sWhere += "  and a.dem_date >= '{FDATE}' ";
                if (to_date != "NULL")
                    sWhere += "  and a.dem_date <= '{EDATE}' ";

               
                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);
                sWhere = sWhere.Replace("{FYEAR}", year_code);
                sWhere = sWhere.Replace("{FDATE}", from_date);
                sWhere = sWhere.Replace("{EDATE}", to_date);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM dailyexpm  a ";
                    sql += " left join customerm party on a.dem_party_id = party.cust_pkid ";
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
                sql += " select dem_pkid, dem_year, dem_cfno, dem_date,dem_inv_date,dem_exp_date, ";
                sql += " party.cust_name as dem_party_name,";
                sql += " a.rec_created_date,row_number() over(order by dem_cfno) rn ";
                sql += " from dailyexpm a ";
                sql += " left join customerm party on a.dem_party_id = party.cust_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by dem_cfno";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new DailyExpm();
                    mRow.dem_pkid = Dr["dem_pkid"].ToString();
                    mRow.dem_cfno = Lib.Conv2Integer(Dr["dem_cfno"].ToString());
                    mRow.dem_date = Lib.DatetoStringDisplayformat(Dr["dem_date"]);
                    mRow.dem_party_name = Dr["dem_party_name"].ToString();
                    mRow.dem_inv_date = Lib.DatetoStringDisplayformat(Dr["dem_inv_date"]);
                    mRow.dem_exp_date = Lib.DatetoStringDisplayformat(Dr["dem_exp_date"]);
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
            DailyExpm mRow = new DailyExpm();

            string id = SearchData["pkid"].ToString();
            string mode = SearchData["mode"].ToString();

            try
            {

                if (mode == "ADD")
                {
                    mRow.dem_pkid = id;
                    mRow.dem_cfno = 0;
                    mRow.dem_date = Lib.DatetoString(DateTime.Now);
                    mRow.dem_genjob_id = "";
                    mRow.dem_genjob_no = "";
                    mRow.dem_genjob_prefix = "";
                    mRow.dem_party_id = "";
                    mRow.dem_party_code = "";
                    mRow.dem_party_name = "";
                    mRow.dem_party_br_id = "";
                    mRow.dem_party_br_no = "";
                    mRow.dem_party_br_addr = "";
                    mRow.dem_inv_date = "";
                    mRow.dem_exp_date = "";
                    mRow.dem_edit_code = "";
                    mRow.lock_record = false;

                    mRow.dem_party_br_gst = "";
                    mRow.dem_driver_name = "";
                    mRow.dem_container = "";
                    mRow.dem_vehicle_no = "";
                    mRow.dem_from = "";
                    mRow.dem_to = "";
                }


                DataTable Dt_Rec = new DataTable();

                sql = " select dem_pkid,dem_year,dem_cfno,dem_date,dem_genjob_id,gjob.hbl_no as dem_genjob_no,gjob.hbl_prefix as dem_genjob_prefix,";
                sql += " dem_party_id , party.cust_code as  dem_party_code, party.cust_name as  dem_party_name,a.dem_party_br_id ,partyaddr.add_gstin as  dem_party_br_gst,";
                sql += " partyaddr.add_branch_slno as  dem_party_br_no,partyaddr.add_line1||'\n'||partyaddr.add_line2||'\n'||partyaddr.add_line3 as dem_party_br_addr, ";
                sql += " dem_inv_date ,dem_exp_date ,dem_edit_code, ";
                sql += " gjob.hbl_book_cntr as dem_container,gj.gj_driver_name as dem_driver_name,gj.gj_vehicle_no as dem_vehicle_no,gj.gj_from as dem_from,gj.gj_to1 as dem_to ";
                sql += " from dailyexpm a  ";
                sql += " inner join hblm gjob on a.dem_genjob_id = gjob.hbl_pkid ";
                sql += " inner join genjobm gj on gjob.hbl_pkid = gj.gj_parent_id";
                sql += " left join customerm party on a.dem_party_id = party.cust_pkid ";
                sql += " left join addressm partyaddr on a.dem_party_br_id = partyaddr.add_pkid ";

                sql += " where  a.dem_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow.dem_pkid = Dr["dem_pkid"].ToString();
                    mRow.dem_year = Lib.Conv2Integer(Dr["dem_year"].ToString());
                    mRow.dem_cfno = Lib.Conv2Integer(Dr["dem_cfno"].ToString());
                    mRow.dem_date = Lib.DatetoString(Dr["dem_date"]);
                    mRow.dem_genjob_id = Dr["dem_genjob_id"].ToString();
                    mRow.dem_genjob_no = Dr["dem_genjob_no"].ToString();
                    mRow.dem_genjob_prefix = Dr["dem_genjob_prefix"].ToString();
                    mRow.dem_party_id = Dr["dem_party_id"].ToString();
                    mRow.dem_party_code = Dr["dem_party_code"].ToString();
                    mRow.dem_party_name = Dr["dem_party_name"].ToString();
                    mRow.dem_party_br_id = Dr["dem_party_br_id"].ToString();
                    mRow.dem_party_br_no = Dr["dem_party_br_no"].ToString();
                    mRow.dem_party_br_addr = Dr["dem_party_br_addr"].ToString();
                    mRow.dem_inv_date = Lib.DatetoString(Dr["dem_inv_date"]);
                    mRow.dem_exp_date = Lib.DatetoString(Dr["dem_exp_date"]);
                    mRow.dem_party_br_gst = Dr["dem_party_br_gst"].ToString();
                    mRow.dem_driver_name = Dr["dem_driver_name"].ToString();
                    mRow.dem_container = Dr["dem_container"].ToString();
                    mRow.dem_vehicle_no = Dr["dem_vehicle_no"].ToString();
                    mRow.dem_from = Dr["dem_from"].ToString();
                    mRow.dem_to = Dr["dem_to"].ToString();

                    mRow.dem_edit_code = Dr["dem_edit_code"].ToString();
                    mRow.lock_record = true;
                    if (Dr["dem_edit_code"].ToString().IndexOf("{S}") >= 0)
                        mRow.lock_record = false;

                    break;
                }

                List<DailyExpd> mList = new List<DailyExpd>();
                DailyExpd dRow;

                if (mode == "ADD")
                {
                    sql = " select a.ded_pkid as ded_headerid, a.ded_slno as ded_slno,";
                    sql += "  a.ded_acid , ac.acc_code as ded_accode,ac.acc_name as ded_acname,";
                    sql += "  a.ded_nature,";
                    sql += "  a.ded_accrid,accr.acc_code as ded_accrcode,a.ded_type ,";
                    sql += "  a.ded_amt as ded_amt ,nvl(a.rec_deleted,'N') as rec_deleted ";
                    sql += "  from dailyexpd a";
                    sql += "  left join acctm ac on a.ded_acid = ac.acc_pkid";
                    sql += "  left join acctm accr on a.ded_accrid = accr.acc_pkid";
                    sql += "  where a.ded_parentid = '0' ";
                    sql += "  order by a.ded_slno";
                }
                else
                {
                    sql = " select a.ded_pkid as ded_headerid, a.ded_slno as ded_slno,";
                    sql += "  a.ded_acid , ac.acc_code as ded_accode,ac.acc_name as ded_acname,";
                    sql += "  a.ded_nature,";
                    sql += "  a.ded_accrid,accr.acc_code as ded_accrcode,a.ded_type ,";
                    sql += "  b.ded_amt as ded_amt,nvl(a.rec_deleted,'N') as rec_deleted  ";
                    sql += "  from dailyexpd a";
                    sql += "  left join dailyexpd b on  a.ded_pkid =  b.ded_headerid and b.ded_parentid = '{ID}'";
                    sql += "  left join acctm ac on nvl(b.ded_acid,a.ded_acid) = ac.acc_pkid";
                    sql += "  left join acctm accr on nvl(b.ded_accrid,a.ded_accrid) = accr.acc_pkid";
                    sql += "  where a.ded_parentid = '0' ";
                    sql += "  order by a.ded_slno";

                    sql = sql.Replace("{ID}", id);
                }


                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    dRow = new DailyExpd();
                    dRow.ded_slno = Lib.Conv2Integer(Dr["ded_slno"].ToString());
                    dRow.ded_acid = Dr["ded_acid"].ToString();
                    dRow.ded_accode = Dr["ded_accode"].ToString();
                    dRow.ded_acname = Dr["ded_acname"].ToString();
                    dRow.ded_nature = Dr["ded_nature"].ToString();
                    dRow.ded_accrid = Dr["ded_accrid"].ToString();
                    dRow.ded_accrcode = Dr["ded_accrcode"].ToString();
                    dRow.ded_headerid = Dr["ded_headerid"].ToString();
                    dRow.ded_type = Dr["ded_type"].ToString();
                    dRow.ded_amt = Lib.Convert2Decimal(Dr["ded_amt"].ToString());
                    dRow.ded_old_amt = Lib.Convert2Decimal(Dr["ded_amt"].ToString());
                    dRow.rec_deleted = Dr["rec_deleted"].ToString();

                    mList.Add(dRow);
                }
                mRow.detList = mList;
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

        public string AllValid(DailyExpm Record)
        {
            string str = "";
            try
            {
                //if (Record.mbl_no.Trim().Length > 0)
                //{
                //    sql = "select hbl_pkid from (";
                //    sql += "select hbl_pkid  from hblm a where a.hbl_bl_no = '{BLNO}'  ";
                //    sql += " and a.rec_company_code = '{COMPCODE}'";
                //    sql += " and a.rec_branch_code = '{BRCODE}'";
                //    sql += " and a.rec_category = '{CATEGORY}'";
                //    sql += " and a.hbl_type = 'MBL-AE'";
                //    sql += ") a where hbl_pkid <> '{PKID}'";

                //    sql = sql.Replace("{BLNO}", Record.mbl_no.Trim());
                //    sql = sql.Replace("{PKID}", Record.mbl_pkid);
                //    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                //    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                //    sql = sql.Replace("{CATEGORY}", Record.rec_category);

                //    if (Con_Oracle.IsRowExists(sql))
                //        Lib.AddError(ref str, " | This Master No Already Exists");
                //}
                //if (Record.mbl_folder_no.Trim().Length > 0 && Record.mbl_folder_no.ToUpper().Trim() != "DIRECT")
                //{
                //    sql = "select hbl_pkid from (";
                //    sql += "select hbl_pkid  from hblm a where a.hbl_folder_no = '{FOLDERNO}'  ";
                //    sql += " and a.rec_company_code = '{COMPCODE}'";
                //    sql += " and a.rec_branch_code = '{BRCODE}'";
                //    sql += " and a.rec_category = '{CATEGORY}'";
                //    sql += " and a.hbl_type = 'MBL-AE'";
                //    sql += ") a where hbl_pkid <> '{PKID}'";

                //    sql = sql.Replace("{FOLDERNO}", Record.mbl_folder_no.Trim());
                //    sql = sql.Replace("{PKID}", Record.mbl_pkid);
                //    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                //    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                //    sql = sql.Replace("{CATEGORY}", Record.rec_category);

                //    if (Con_Oracle.IsRowExists(sql))
                //        Lib.AddError(ref str, " | Folder No Already Exists");
                //}

                if (Record.dem_party_id.Trim().Length > 0 || Record.dem_party_br_id.Trim().Length > 0)
                {
                    sql = "select add_pkid from addressm where add_pkid = '{ADD_BRID}'";
                    sql += " and  add_parent_id = '{PARENT_ID}'";
                    sql = sql.Replace("{ADD_BRID}", Record.dem_party_br_id);
                    sql = sql.Replace("{PARENT_ID}", Record.dem_party_id);
                    if (!Con_Oracle.IsRowExists(sql))
                        Lib.AddError(ref str, " Invalid Party Address ");
                }
                
                if (Lib.IsFutureDate(Record.dem_date))
                {
                    Lib.AddError(ref str, " Date Cannot be a Future Date ");
                }
            }
            catch (Exception Ex)
            {
                Lib.AddError(ref str, Ex.Message.ToString());
            }
            return str;
        }


        public Dictionary<string, object> Save(DailyExpm Record)
        {
            string sql = "";
            string CfNo = "";
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


                if (Record.rec_mode == "ADD")
                {
                    sql = "select nvl(max(dem_cfno)+1,1001) as cfno from dailyexpm a ";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                    sql += " and a.dem_year =  {FYEAR} ";

                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                    sql = sql.Replace("{FYEAR}", Record._globalvariables.year_code);

                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        CfNo = Dt_Temp.Rows[0]["cfno"].ToString();
                        Record.dem_cfno = Lib.Conv2Integer(Dt_Temp.Rows[0]["cfno"].ToString());
                    }
                    else
                    {
                        ErrorMessage = "CF Number Not Found Try again";

                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception(ErrorMessage);
                    }
                }


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("Dailyexpm", Record.rec_mode, "dem_pkid", Record.dem_pkid);
                Rec.InsertDate("dem_date", Record.dem_date);
                Rec.InsertString("dem_genjob_id", Record.dem_genjob_id);
                Rec.InsertString("dem_party_id", Record.dem_party_id);
                Rec.InsertString("dem_party_br_id", Record.dem_party_br_id);
                Rec.InsertDate("dem_inv_date", Record.dem_inv_date);
                Rec.InsertDate("dem_exp_date", Record.dem_exp_date);


                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("dem_edit_code", "{S}");
                    Rec.InsertNumeric("dem_cfno", Lib.Conv2Integer(Record.dem_cfno.ToString()).ToString());
                    Rec.InsertNumeric("dem_year", Lib.Conv2Integer(Record._globalvariables.year_code).ToString());
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_created_date", "SYSDATE"); 
                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_edited_date", "SYSDATE");
                }

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);

                sql = "delete from dailyexpd where ded_parentid ='" + Record.dem_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                foreach (DailyExpd Row in Record.detList)
                {
                    if (Row.ded_amt != 0)
                    {
                        Row.ded_pkid = Guid.NewGuid().ToString().ToUpper();
                        Rec.CreateRow("Dailyexpd", "ADD", "ded_pkid", Row.ded_pkid);
                        Rec.InsertString("ded_parentid", Record.dem_pkid);
                        Rec.InsertString("ded_headerid", Row.ded_headerid);
                        Rec.InsertNumeric("ded_slno", Lib.Conv2Integer(Row.ded_slno.ToString()).ToString());
                        Rec.InsertString("ded_acid", Row.ded_acid);
                        Rec.InsertString("ded_nature", Row.ded_nature);
                        Rec.InsertString("ded_accrid", Row.ded_accrid);
                        Rec.InsertString("ded_type", Row.ded_type);
                        Rec.InsertNumeric("ded_amt", Lib.Conv2Decimal(Row.ded_amt.ToString()).ToString());
                        Rec.InsertString("rec_deleted", "N");
                        sql = Rec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                }

                //string ccType = "";
                //ccType = "";
                ////Costcenter Updation
                //sql = Lib.GetCostCenterSQL(Record.rec_mode, Record.dem_pkid, Record.dem_cfno.ToString(), Record.dem_cfno.ToString(), Record.dem_date, ccType,
                //    Record._globalvariables.year_code, "", Record._globalvariables.comp_code, Record._globalvariables.branch_code, "");
                //Con_Oracle.ExecuteNonQuery(sql);

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
            
            RetData.Add("cfno", CfNo);
            return RetData;
        }
         
 
    }
}
