using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;


using DataBase;
using DataBase_Oracle.Connections;

namespace BLAccounts
{
    public class TaxmService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Taxm> mList = new List<Taxm>();
            Taxm mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  a.rec_company_code = '" + SearchData["company_code"].ToString() + "' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(tax_desc) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(b.acc_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(b.acc_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  ";
                    sql += " FROM taxm a inner join acctm b on a.tax_acc_id=b.acc_pkid ";
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
                sql += " select a.*,param_code as tax_sac_code from ( ";
                sql += "  select  tax_pkid,tax_desc,tax_acc_id,acc_code, acc_name,acc_sac_id,";
                sql += "  tax_from_dt,tax_to_dt, tax_cgst_rate, tax_sgst_rate, tax_igst_rate,";
                sql += "  row_number() over(order by acc_name, tax_from_dt) rn ";
                sql += "  from taxm a  ";
                sql += "  inner join acctm b on a.tax_acc_id = b.acc_pkid ";
                sql += sWhere;
                sql += ") a ";
                sql += "  left  join  param c on a.acc_sac_id = c.param_pkid ";
                sql += " where rn between {startrow} and {endrow}";
                sql += " order by acc_name, tax_from_dt";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Taxm();
                    mRow.tax_pkid = Dr["tax_pkid"].ToString();
                    mRow.tax_desc = Dr["tax_desc"].ToString();
                    mRow.tax_acc_code = Dr["acc_code"].ToString();
                    mRow.tax_acc_name = Dr["acc_name"].ToString();
                    mRow.tax_sac_code = Dr["tax_sac_code"].ToString();

                    mRow.tax_from_dt =  Lib.DatetoStringDisplayformat(Dr["tax_from_dt"]);
                    if (Dr["tax_to_dt"].Equals(DBNull.Value))
                        mRow.tax_to_dt = "";
                    else
                        mRow.tax_to_dt = Lib.DatetoStringDisplayformat(Dr["tax_to_dt"]);

                    mRow.tax_cgst_rate = Lib.Conv2Decimal(Dr["tax_cgst_rate"].ToString());
                    mRow.tax_sgst_rate = Lib.Conv2Decimal(Dr["tax_sgst_rate"].ToString());
                    mRow.tax_igst_rate = Lib.Conv2Decimal(Dr["tax_igst_rate"].ToString());
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
            Taxm mRow = new Taxm();

            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select ";
                sql += " tax_pkid,tax_desc,tax_acc_id,acc_code, acc_name, tax_acc_id, tax_from_dt,tax_to_dt, ";
                sql += " tax_cgst_rate,tax_sgst_rate,tax_igst_rate ";
                sql += " from taxm a ";
                sql += " left join acctm b on a.tax_acc_id = b.acc_pkid ";
                sql += " where  a.tax_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Taxm();
                    mRow.tax_pkid = Dr["tax_pkid"].ToString();
                    mRow.tax_desc = Dr["tax_desc"].ToString();
                    mRow.tax_acc_id = Dr["tax_acc_id"].ToString();
                    mRow.tax_acc_code = Dr["acc_code"].ToString();
                    mRow.tax_acc_name = Dr["acc_name"].ToString();

                    mRow.tax_from_dt = Lib.DatetoString(Dr["tax_from_dt"]);

                    if (Dr["tax_to_dt"].Equals(DBNull.Value))
                        mRow.tax_to_dt = "";
                    else
                        mRow.tax_to_dt = Lib.DatetoString(Dr["tax_to_dt"]);

                    mRow.tax_cgst_rate = Lib.Conv2Decimal(Dr["tax_cgst_rate"].ToString());
                    mRow.tax_sgst_rate = Lib.Conv2Decimal(Dr["tax_sgst_rate"].ToString());
                    mRow.tax_igst_rate = Lib.Conv2Decimal(Dr["tax_igst_rate"].ToString());

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


        public string AllValid(Taxm  Record)
        {
            string str = "";
            try
            {

                sql += " select tax_from_dt  from taxm  where rec_company_code = '{COMPANY}' ";
                sql += " and tax_pkid <> '{PKID}' and tax_acc_id = '{ACCID}'";
                sql += " and tax_to_dt is null ";

                sql = sql.Replace("{COMPANY}", Record._globalvariables.comp_code);
                sql = sql.Replace("{PKID}", Record.tax_pkid);
                sql = sql.Replace("{ACCID}", Record.tax_acc_id);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Blank To-date exists";

                if (str == "")
                {
                    sql = "";
                    sql += " select tax_from_dt from taxm  where  rec_company_code = '{COMPANY}' ";
                    sql += " and tax_pkid <> '{PKID}' and tax_acc_id = '{ACCID}' ";
                    sql += " and '{FROMDT}' between tax_from_dt and nvl( tax_to_dt, sysdate) ";

                    sql = sql.Replace("{COMPANY}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{FROMDT}", Lib.StringToDate(Record.tax_from_dt));
                    sql = sql.Replace("{PKID}", Record.tax_pkid);
                    sql = sql.Replace("{ACCID}", Record.tax_acc_id);

                    if (Con_Oracle.IsRowExists(sql))
                        str = "Tax Setup Exist for same period";
                }
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Taxm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();


                //if (Record.acc_type_id.Trim().Length <= 0)
                //    Lib.AddError(ref ErrorMessage, "A/c Type Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("taxm", Record.rec_mode, "tax_pkid", Record.tax_pkid);
                Rec.InsertString("tax_desc", Record.tax_desc);
                Rec.InsertString("tax_acc_id", Record.tax_acc_id);
                Rec.InsertDate("tax_from_dt", Record.tax_from_dt);
                Rec.InsertDate("tax_to_dt", Record.tax_to_dt);

                Rec.InsertNumeric("tax_cgst_rate", Record.tax_cgst_rate.ToString());
                Rec.InsertNumeric("tax_sgst_rate", Record.tax_sgst_rate.ToString());
                Rec.InsertNumeric("tax_igst_rate", Record.tax_igst_rate.ToString());

                if (Record.rec_mode == "ADD")
                {
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
            return RetData;
        }


        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "actypem");
            RetData.Add("actypem", lovservice.Lov(parameter)["actypem"]);

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "acgroupm");
            RetData.Add("acgroupm", lovservice.Lov(parameter)["acgroupm"]);

            return RetData;
        }


    }
}
