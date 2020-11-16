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
    public class PayRequestService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<PayRequestm> mList = new List<PayRequestm>();
            PayRequestm mRow;

            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string year_code = SearchData["year_code"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {

                sWhere = " where a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.rec_branch_code = '{BRCODE}'";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  a.pay_no  like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(party.acc_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }
        
                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);
        
                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM payrequestm  a ";
                    sql += " left join acctm party on a.pay_acc_id = party.acc_pkid ";
                   
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
                sql += " select pay_pkid,pay_no,pay_docno,pay_date,a.rec_created_date,a.rec_created_by";
                sql += " ,party.acc_name as pay_acc_name,pay_amt,pay_jvh_id,pay_parent_id,pay_chq_name  ";
                sql += " ,row_number() over(order by a.pay_no) rn ";
                sql += " from payrequestm a ";
                sql += " left join acctm party on a.pay_acc_id = party.acc_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by pay_no";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new PayRequestm();
                    mRow.pay_pkid = Dr["pay_pkid"].ToString();
                    mRow.pay_no = Lib.Conv2Integer(Dr["pay_no"].ToString());
                    mRow.rec_created_date = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    mRow.rec_created_by = Dr["rec_created_by"].ToString();
                    mRow.pay_docno = Dr["pay_docno"].ToString();
                    mRow.pay_date = Lib.DatetoStringDisplayformat(Dr["pay_date"]);
                    mRow.pay_acc_name = Dr["pay_acc_name"].ToString();
                    mRow.pay_amt = Lib.Convert2Decimal(Dr["pay_amt"].ToString());
                    mRow.pay_jvh_id = Dr["pay_jvh_id"].ToString();
                    mRow.pay_parent_id = Dr["pay_parent_id"].ToString();
                    mRow.pay_chq_name= Dr["pay_chq_name"].ToString();
                    mRow.rowdisplayed = false;
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
            PayRequestm mRow = new PayRequestm();
            string id = "";
            if (SearchData.ContainsKey("pkid"))
                id = SearchData["pkid"].ToString();


            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select pay_pkid,pay_no,pay_date";
                sql += " ,pay_acc_id,party.acc_code as pay_acc_code,party.acc_name as pay_acc_name ";
                sql += " from payrequestm a ";
                sql += " left join acctm party on a.pay_acc_id = party.acc_pkid ";
                sql += " where  a.pay_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new PayRequestm();
                    mRow.pay_pkid = Dr["pay_pkid"].ToString();
                    mRow.pay_no = Lib.Conv2Integer(Dr["pay_no"].ToString());
                    mRow.pay_date = Lib.DatetoString(Dr["pay_date"]);
                    mRow.pay_acc_id = Dr["pay_acc_id"].ToString();
                    mRow.pay_acc_code = Dr["pay_acc_code"].ToString();
                    mRow.pay_acc_name = Dr["pay_acc_name"].ToString();
                    break;
                }

                Dt_Rec.Rows.Clear();
                Con_Oracle.CloseConnection();
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

        public string AllValid(PayRequestm Record)
        {
            string str = "";

            return str;
        }

        public Dictionary<string, object> Save(PayRequestm Record)
        {
            string sql = "";
            string DocNo = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            DataTable Dt_Temp;
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
                    sql = "select nvl(max(pay_no) + 1,1001) as payno from payrequestm a ";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.rec_branch_code = '{BRCODE}'";
     
                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);

                    Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        DocNo = Dt_Temp.Rows[0]["payno"].ToString();
                        Record.pay_no = Lib.Conv2Integer(Dt_Temp.Rows[0]["payno"].ToString());
                    }
                    else
                    {
                        ErrorMessage = "Payment Request Number Not Found Try again";

                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception(ErrorMessage);
                    }
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("payrequestm", Record.rec_mode, "pay_pkid", Record.pay_pkid);

                Rec.InsertDate("pay_date", Record.pay_date);
                Rec.InsertString("pay_acc_id", Record.pay_acc_id);
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("pay_no", Record.pay_no.ToString());
                    Rec.InsertString("rec_deleted", "N");
                    Rec.InsertNumeric("pay_year", Record._globalvariables.year_code);
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
            RetData.Add("payno", DocNo);
            return RetData;
        }



        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            //Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "SALES EXECUTIVE");
            //RetData.Add("smanlist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "CITY");
            //RetData.Add("citylist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "STATE");
            //RetData.Add("statelist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;
        }


    }
}
