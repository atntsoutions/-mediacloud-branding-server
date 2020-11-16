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
    public class CostCenterService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<CostCenterm> mList = new List<CostCenterm>();
            CostCenterm mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];


            string comp_code = SearchData["comp_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string cc_type = SearchData["cc_type"].ToString();
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  a.rec_company_code = '" + comp_code + "'";
                if (cc_type == "EMPLOYEE")
                {
                    sWhere += " and a.cc_type = 'EMPLOYEE' ";
                }
                if (cc_type == "COST CENTER")
                {
                    sWhere += " and a.rec_branch_code = '" + branch_code + "'";
                    sWhere += " and a.cc_type = 'COST CENTER' ";
                }
               
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.cc_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.cc_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM costcenterm a ";
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
                sql += "  select  cc_pkid,  cc_code, cc_name,cc_type, ";
                sql += "  row_number() over(order by cc_type,cc_code) rn ";
                sql += "  from costcenterm a  ";
                sql += sWhere;
                sql += ") a ";
                sql += " where rn between {startrow} and {endrow}";
                sql += " order by cc_type,cc_code";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new CostCenterm();
                    mRow.cc_pkid = Dr["cc_pkid"].ToString();
                    mRow.cc_code = Dr["cc_code"].ToString();
                    mRow.cc_name = Dr["cc_name"].ToString();
                    mRow.cc_type = Dr["cc_type"].ToString();
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
            CostCenterm mRow = new CostCenterm();
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  cc_pkid,  cc_code, cc_name,cc_type  ";
                sql += " from costcenterm a ";
                sql += " where  a.cc_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new CostCenterm();
                    mRow.cc_pkid = Dr["cc_pkid"].ToString();
                    mRow.cc_code = Dr["cc_code"].ToString();
                    mRow.cc_name = Dr["cc_name"].ToString();

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

        public string AllValid(CostCenterm Record)
        {
            string str = "";
            try
            {

                sql = "select cc_pkid from (";
                sql += " select cc_pkid from costcenterm a where a.rec_company_code = '{COMPANY_CODE}'  ";

                if (Record.cc_type == "COST CENTER")
                    sql += " and a.rec_branch_code = '{BRANCH_CODE}' ";

                sql += " and a.cc_type = '{CCTYPE}' ";
                sql += " and (a.cc_code = '{CODE}' or a.cc_name = '{NAME}')  ";
                sql += ") a where cc_pkid <> '{PKID}'";

                sql = sql.Replace("{COMPANY_CODE}", Record._globalvariables.comp_code);
                sql = sql.Replace("{BRANCH_CODE}", Record._globalvariables.branch_code);
                sql = sql.Replace("{CODE}", Record.cc_code);
                sql = sql.Replace("{NAME}", Record.cc_name);
                sql = sql.Replace("{PKID}", Record.cc_pkid);
                sql = sql.Replace("{CCTYPE}", Record.cc_type);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";


            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(CostCenterm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if (Record.cc_code.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Code Cannot Be Empty");

                if (Record.cc_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("costcenterm", Record.rec_mode, "cc_pkid", Record.cc_pkid);
                Rec.InsertString("cc_code", Record.cc_code);
                Rec.InsertString("cc_name", Record.cc_name);
                if (Record.rec_mode == "ADD")
                {
                    if (Record.cc_type == "EMPLOYEE")
                        Rec.InsertString("cc_type", "EMPLOYEE");
                    if (Record.cc_type == "COST CENTER")
                    {
                        Rec.InsertString("cc_type", "COST CENTER");
                        Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                    }

                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);

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
            //Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            return RetData;
        }


    }
}
