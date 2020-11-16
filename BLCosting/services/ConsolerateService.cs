using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

using XL.XSheet;


namespace BLCosting
{
    public class ConsolerateService : BL_Base
    {


        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Consolerate> mList = new List<Consolerate>();
            Consolerate mRow;

            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();

            Boolean min_rate = (Boolean)SearchData["min_rate"];

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {

                sWhere = " where a.rec_company_code = '{COMPANY_CODE}' ";
                if (min_rate)
                    sWhere += " and a.cr_rate_type = 'MINRATE' ";
                else
                    sWhere += " and a.cr_rate_type = 'DESTIN' ";

                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += " upper(a.cr_agent_name)  like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.cr_branch_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.cr_rate_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.cr_rate_value) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                sWhere = sWhere.Replace("{COMPANY_CODE}", company_code);
                // sWhere = sWhere.Replace("{BRANCH_CODE}", branch_code);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM consolerate  a ";
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
                sql += " select cr_pkid,cr_agent_name,cr_branch_code,cr_branch_name,cr_cntr_type,cr_ex_rate_gbp, ";
                sql += " cr_org_inc_thc,cr_org_inc_bl,cr_org_exp_thc,cr_org_exp_emtyplce,cr_org_exp_misc,cr_org_exp_stuff,";
                sql += " cr_org_exp_trans,cr_org_exp_surrend,cr_org_exp_cfs,cr_org_exp_survey,";
                sql += " cr_rate_code,cr_rate_value,";
                sql += " row_number() over(order by cr_branch_code,cr_agent_name) rn ";
                sql += " from consolerate a ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by cr_branch_code,cr_agent_name";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Consolerate();
                    mRow.cr_pkid = Dr["cr_pkid"].ToString();
                    mRow.cr_agent_name = Dr["cr_agent_name"].ToString();
                    mRow.cr_branch_code = Dr["cr_branch_code"].ToString();
                    mRow.cr_branch_name = Dr["cr_branch_name"].ToString();
                    mRow.cr_cntr_type = Dr["cr_cntr_type"].ToString();
                    mRow.cr_ex_rate_gbp = Lib.Conv2Decimal(Dr["cr_ex_rate_gbp"].ToString());
                    mRow.cr_org_inc_thc = Lib.Conv2Decimal(Dr["cr_org_inc_thc"].ToString());
                    mRow.cr_org_inc_bl = Lib.Conv2Decimal(Dr["cr_org_inc_bl"].ToString());
                    mRow.cr_org_exp_thc = Lib.Conv2Decimal(Dr["cr_org_exp_thc"].ToString());
                    mRow.cr_org_exp_emtyplce = Lib.Conv2Decimal(Dr["cr_org_exp_emtyplce"].ToString());
                    mRow.cr_org_exp_misc = Lib.Conv2Decimal(Dr["cr_org_exp_misc"].ToString());
                    mRow.cr_org_exp_stuff = Lib.Conv2Decimal(Dr["cr_org_exp_stuff"].ToString());
                    mRow.cr_org_exp_trans = Lib.Conv2Decimal(Dr["cr_org_exp_trans"].ToString());
                    mRow.cr_org_exp_surrend = Lib.Conv2Decimal(Dr["cr_org_exp_surrend"].ToString());
                    mRow.cr_org_exp_cfs = Lib.Conv2Decimal(Dr["cr_org_exp_cfs"].ToString());
                    mRow.cr_org_exp_survey = Lib.Conv2Decimal(Dr["cr_org_exp_survey"].ToString());
                    mRow.cr_rate_code = Dr["cr_rate_code"].ToString();
                    mRow.cr_rate_value = Lib.Conv2Decimal(Dr["cr_rate_value"].ToString());

                    mList.Add(mRow);
                }

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
            Consolerate mRow = new Consolerate();

            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select cr_pkid,cr_agent_name,cr_branch_code,cr_branch_name,cr_cntr_type,cr_ex_rate_gbp, ";
                sql += " cr_org_inc_thc,cr_org_inc_bl,cr_org_exp_thc, cr_org_exp_emtyplce,cr_org_exp_misc,";
                sql += " cr_org_exp_stuff,cr_org_exp_trans,cr_org_exp_cfs,cr_org_exp_survey,cr_org_exp_cseal,cr_org_exp_surrend, ";
                sql += " cr_des_inc_thc,cr_des_inc_hndg_cbm,cr_des_inc_hndg_ton,cr_des_inc_bl,";
                sql += " cr_des_exp_terml,cr_des_exp_bl,cr_des_exp_shunt,cr_des_exp_unpack,cr_des_exp_lolo,";
                sql += " cr_des_exp_securty,cr_des_exp_isps,cr_des_exp_tpw,cr_des_exp_tdoc,cr_rate_code,cr_rate_value ";
                sql += " from consolerate a  ";
                sql += " where  a.cr_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Consolerate();
                    mRow.cr_pkid = Dr["cr_pkid"].ToString();
                    mRow.cr_agent_name = Dr["cr_agent_name"].ToString();
                    mRow.cr_branch_code = Dr["cr_branch_code"].ToString();
                    mRow.cr_branch_name = Dr["cr_branch_name"].ToString();
                    mRow.cr_cntr_type = Dr["cr_cntr_type"].ToString();
                    mRow.cr_ex_rate_gbp = Lib.Conv2Decimal(Dr["cr_ex_rate_gbp"].ToString());
                    mRow.cr_org_inc_thc = Lib.Conv2Decimal(Dr["cr_org_inc_thc"].ToString());
                    mRow.cr_org_inc_bl = Lib.Conv2Decimal(Dr["cr_org_inc_bl"].ToString());
                    mRow.cr_org_exp_thc = Lib.Conv2Decimal(Dr["cr_org_exp_thc"].ToString());
                    mRow.cr_org_exp_emtyplce = Lib.Conv2Decimal(Dr["cr_org_exp_emtyplce"].ToString());
                    mRow.cr_org_exp_misc = Lib.Conv2Decimal(Dr["cr_org_exp_misc"].ToString());
                    mRow.cr_org_exp_stuff = Lib.Conv2Decimal(Dr["cr_org_exp_stuff"].ToString());
                    mRow.cr_org_exp_trans = Lib.Conv2Decimal(Dr["cr_org_exp_trans"].ToString());
                    mRow.cr_org_exp_cfs = Lib.Conv2Decimal(Dr["cr_org_exp_cfs"].ToString());
                    mRow.cr_org_exp_survey = Lib.Conv2Decimal(Dr["cr_org_exp_survey"].ToString());
                    mRow.cr_org_exp_cseal = Lib.Conv2Decimal(Dr["cr_org_exp_cseal"].ToString());
                    mRow.cr_org_exp_surrend = Lib.Conv2Decimal(Dr["cr_org_exp_surrend"].ToString());
                    mRow.cr_des_inc_thc = Lib.Conv2Decimal(Dr["cr_des_inc_thc"].ToString());
                    mRow.cr_des_inc_hndg_cbm = Lib.Conv2Decimal(Dr["cr_des_inc_hndg_cbm"].ToString());
                    mRow.cr_des_inc_hndg_ton = Lib.Conv2Decimal(Dr["cr_des_inc_hndg_ton"].ToString());
                    mRow.cr_des_inc_bl = Lib.Conv2Decimal(Dr["cr_des_inc_bl"].ToString());
                    mRow.cr_des_exp_terml = Lib.Conv2Decimal(Dr["cr_des_exp_terml"].ToString());
                    mRow.cr_des_exp_bl = Lib.Conv2Decimal(Dr["cr_des_exp_bl"].ToString());
                    mRow.cr_des_exp_shunt = Lib.Conv2Decimal(Dr["cr_des_exp_shunt"].ToString());
                    mRow.cr_des_exp_unpack = Lib.Conv2Decimal(Dr["cr_des_exp_unpack"].ToString());
                    mRow.cr_des_exp_lolo = Lib.Conv2Decimal(Dr["cr_des_exp_lolo"].ToString());
                    mRow.cr_des_exp_securty = Lib.Conv2Decimal(Dr["cr_des_exp_securty"].ToString());
                    mRow.cr_des_exp_isps = Lib.Conv2Decimal(Dr["cr_des_exp_isps"].ToString());
                    mRow.cr_des_exp_tpw = Lib.Conv2Decimal(Dr["cr_des_exp_tpw"].ToString());
                    mRow.cr_des_exp_tdoc = Lib.Conv2Decimal(Dr["cr_des_exp_tdoc"].ToString());
                    mRow.cr_rate_code = Dr["cr_rate_code"].ToString();
                    mRow.cr_rate_value = Lib.Conv2Decimal(Dr["cr_rate_value"].ToString());

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

        public string AllValid(Consolerate Record)
        {
            string str = "";

            try
            {
                if (Record.min_rate)
                {
                    if (Record.cr_rate_code.ToString() == "")
                        Lib.AddError(ref str, " | Code Cannot be Blank");

                    if (Lib.Conv2Decimal(Record.cr_rate_value.ToString()) == 0)
                        Lib.AddError(ref str, " | Per Rate Cannot be Blank");
                }
                if (!Record.min_rate)
                {
                    if (Record.cr_cntr_type.ToString() == "")
                        Lib.AddError(ref str, " | Container Type Cannot be Blank");

                    sql = "";
                    sql = "select cr_pkid from (";
                    sql += " select cr_pkid from consolerate where cr_rate_type = 'DESTIN'";
                    sql += " and cr_branch_code = '{BRANCH_CODE}' ";
                    sql += " and cr_agent_name = '{AGENT}' ";
                    sql += " and cr_cntr_type = '{CNTRTYPE}'";
                    sql += ") a where cr_pkid <> '{PKID}'";

                    sql = sql.Replace("{BRANCH_CODE}", Record.cr_branch_code.ToString());
                    sql = sql.Replace("{AGENT}", Record.cr_agent_name.ToString());
                    sql = sql.Replace("{CNTRTYPE}", Record.cr_cntr_type.ToString());
                    sql = sql.Replace("{PKID}", Record.cr_pkid.ToString());

                    if (Con_Oracle.IsRowExists(sql))
                        Lib.AddError(ref str, " | Record Already Exists ");
                }

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(Consolerate Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";

            try
            {
                Con_Oracle = new DBConnection();

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("consolerate", Record.rec_mode, "cr_pkid", Record.cr_pkid);
                if (Record.min_rate)
                {
                    Rec.InsertString("cr_rate_code", Record.cr_rate_code);
                    Rec.InsertNumeric("cr_rate_value", Lib.Conv2Decimal(Record.cr_rate_value.ToString()).ToString());
                }

                if (!Record.min_rate)
                {
                    Rec.InsertString("cr_agent_name", Record.cr_agent_name);
                    Rec.InsertString("cr_branch_code", Record.cr_branch_code);
                    Rec.InsertString("cr_branch_name", Record.cr_branch_name);
                    Rec.InsertString("cr_cntr_type", Record.cr_cntr_type);
                    Rec.InsertNumeric("cr_ex_rate_gbp", Lib.Conv2Decimal(Record.cr_ex_rate_gbp.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_inc_thc", Lib.Conv2Decimal(Record.cr_org_inc_thc.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_inc_bl", Lib.Conv2Decimal(Record.cr_org_inc_bl.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_thc", Lib.Conv2Decimal(Record.cr_org_exp_thc.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_emtyplce", Lib.Conv2Decimal(Record.cr_org_exp_emtyplce.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_misc", Lib.Conv2Decimal(Record.cr_org_exp_misc.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_stuff", Lib.Conv2Decimal(Record.cr_org_exp_stuff.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_trans", Lib.Conv2Decimal(Record.cr_org_exp_trans.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_cfs", Lib.Conv2Decimal(Record.cr_org_exp_cfs.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_survey", Lib.Conv2Decimal(Record.cr_org_exp_survey.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_cseal", Lib.Conv2Decimal(Record.cr_org_exp_cseal.ToString()).ToString());
                    Rec.InsertNumeric("cr_org_exp_surrend", Lib.Conv2Decimal(Record.cr_org_exp_surrend.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_inc_thc", Lib.Conv2Decimal(Record.cr_des_inc_thc.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_inc_hndg_cbm", Lib.Conv2Decimal(Record.cr_des_inc_hndg_cbm.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_inc_hndg_ton", Lib.Conv2Decimal(Record.cr_des_inc_hndg_ton.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_inc_bl", Lib.Conv2Decimal(Record.cr_des_inc_bl.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_terml", Lib.Conv2Decimal(Record.cr_des_exp_terml.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_bl", Lib.Conv2Decimal(Record.cr_des_exp_bl.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_shunt", Lib.Conv2Decimal(Record.cr_des_exp_shunt.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_unpack", Lib.Conv2Decimal(Record.cr_des_exp_unpack.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_lolo", Lib.Conv2Decimal(Record.cr_des_exp_lolo.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_securty", Lib.Conv2Decimal(Record.cr_des_exp_securty.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_isps", Lib.Conv2Decimal(Record.cr_des_exp_isps.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_tpw", Lib.Conv2Decimal(Record.cr_des_exp_tpw.ToString()).ToString());
                    Rec.InsertNumeric("cr_des_exp_tdoc", Lib.Conv2Decimal(Record.cr_des_exp_tdoc.ToString()).ToString());

                }

                if (Record.rec_mode == "ADD")
                {
                    if (Record.min_rate)
                        Rec.InsertString("cr_rate_type", "MINRATE");
                    else
                        Rec.InsertString("cr_rate_type", "DESTIN");

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
            return RetData;
        }

    }
}

