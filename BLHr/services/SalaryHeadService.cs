using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLHr
{
    public class SalaryHeadService : BL_Base
    {
      
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<salaryheadm> mList = new List<salaryheadm>();
            salaryheadm mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where a.rec_company_code = '" + company_code + "'";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a. sal_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a. sal_desc) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " ) ";
                  
                }
               
                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM salaryheadm  a ";
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
                sql += " select sal_pkid,sal_code,sal_desc,sal_head,sal_head_order, ";
                sql += " b.acc_code as sal_acc_code, b.acc_name as sal_acc_name, ";
                sql += " row_number() over(order by sal_code) rn ";
                sql += " from salaryheadm a ";
                sql += " left join acctm b on a.sal_acc_id = b.acc_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by sal_code";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                
                foreach (DataRow Dr in Dt_List.Rows)
                {

                    mRow = new salaryheadm();
                    mRow.sal_pkid = Dr["sal_pkid"].ToString();
                    mRow.sal_code = Dr["sal_code"].ToString();
                    mRow.sal_desc = Dr["sal_desc"].ToString();
                    mRow.sal_head = Dr["sal_head"].ToString();
                    mRow.sal_acc_code = Dr["sal_acc_code"].ToString();
                    mRow.sal_acc_name = Dr["sal_acc_name"].ToString();
                    mRow.sal_head_order = Lib.Conv2Integer (Dr["sal_head_order"].ToString());

                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);

            return RetData;
        }


    


        public Dictionary<string, object>  GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            salaryheadm mRow =new salaryheadm();
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql += " select sal_pkid,sal_code,sal_desc,sal_head,sal_head_order ";
                sql += " ,sal_acc_id, b.acc_code as sal_acc_code, b.acc_name as sal_acc_name ";
                sql += " from salaryheadm a  ";
                sql += "  left join acctm b on a.sal_acc_id = b.acc_pkid ";
                sql += " where  a.sal_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new salaryheadm();
                    mRow.sal_pkid = Dr["sal_pkid"].ToString();
                    mRow.sal_code = Dr["sal_code"].ToString();
                    mRow.sal_desc = Dr["sal_desc"].ToString();
                    mRow.sal_head = Dr["sal_head"].ToString();
                    mRow.sal_head_order =Lib.Conv2Integer (Dr["sal_head_order"].ToString());
                    mRow.sal_acc_id = Dr["sal_acc_id"].ToString();
                    mRow.sal_acc_code = Dr["sal_acc_code"].ToString();
                    mRow.sal_acc_name = Dr["sal_acc_name"].ToString();

                    break;
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            return RetData;
        }


        public string AllValid(salaryheadm Record)
        { 
            string str = "";
            DateTime tdate = DateTime.Now;
            try
            {
                if (Record.sal_code.Trim().Length <= 0)
                    Lib.AddError(ref str, " | Code Cannot Be Empty");

                if (Record.sal_code.Trim().Length > 0)
                {

                    sql = "select sal_pkid from (";
                    sql += "select sal_pkid  from salaryheadm a where a.rec_company_code = '" + Record._globalvariables.comp_code + "' and a.sal_code = '{CODE}'  ";
                    sql += ") a where sal_pkid <> '{PKID}'";

                    sql = sql.Replace("{CODE}", Record.sal_code);
                    sql = sql.Replace("{PKID}", Record.sal_pkid);

                    if (Con_Oracle.IsRowExists(sql))
                        Lib.AddError(ref str, " | Code Exists");
                }

                if (Record.sal_desc.Trim().Length > 0)
                {

                    sql = "select sal_pkid from (";
                    sql += "select sal_pkid  from salaryheadm a where  a.rec_company_code = '" + Record._globalvariables.comp_code + "' and a.sal_desc = '{NAME}'  ";
                    sql += ") a where sal_pkid <> '{PKID}'";

                    sql = sql.Replace("{NAME}", Record.sal_desc);
                    sql = sql.Replace("{PKID}", Record.sal_pkid);


                    if (Con_Oracle.IsRowExists(sql))

                        Lib.AddError(ref str, " | Description Exists");
                }

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        
        public Dictionary<string, object> Save(salaryheadm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();


                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);



                DBRecord Rec = new DBRecord();
                Rec.CreateRow("salaryheadm", Record.rec_mode, "sal_pkid", Record.sal_pkid);
                Rec.InsertString("sal_code", Record.sal_code);
                Rec.InsertString("sal_desc", Record.sal_desc);
                Rec.InsertString("sal_head", Record.sal_head);
                Rec.InsertString("sal_acc_id", Record.sal_acc_id);
                Rec.InsertNumeric("sal_head_order", Record.sal_head_order.ToString());
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

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //parameter.Add("comp_code", comp_code);
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }

    }
}
