using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase.Connections;

namespace BLAdmin
{
    public class ModuleService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Modulem> mList = new List<Modulem>();
            Modulem mRow;

            string type = SearchData["type"].ToString();
            string comp_code = SearchData["comp_code"].ToString().ToUpper();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  rec_company_code ='" + comp_code + "' " ;
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.module_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM modulem  a "  ;
                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM modulem a ";

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
                sql += "  select  module_pkid, module_name, module_order ,  row_number() over(order by module_order) rn ";
                sql += "  from modulem a " + sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by module_order";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Modulem();
                    mRow.module_pkid = Dr["module_pkid"].ToString();
                    mRow.module_name = Dr["module_name"].ToString();
                    mRow.module_order = Lib.Conv2Integer ( Dr["module_order"].ToString());
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
            Modulem mRow =new Modulem();
            
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  module_pkid, module_name, module_order ";
                sql += " from modulem a  ";
                sql += " where  a.module_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Modulem ();
                    mRow.module_pkid = Dr["module_pkid"].ToString();
                    mRow.module_name = Dr["module_name"].ToString();
                    mRow.module_order = Lib.Conv2Integer(Dr["module_order"].ToString());

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


        public string AllValid(Modulem Record)
        {
            string str = "";
            try
            {
                sql = "select module_pkid from (";
                sql += "select module_pkid  from modulem a where rec_company_code ='" + Record._globalvariables.comp_code + "' ";
                sql += " and (a.module_name = '{NAME}')  ";
                sql += ") a where module_pkid <> '{PKID}'";

                sql = sql.Replace("{NAME}", Record.module_name);
                sql = sql.Replace("{PKID}", Record.module_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Name Exists";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Modulem Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if (Record.module_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("modulem", Record.rec_mode, "module_pkid", Record.module_pkid);
                Rec.InsertString("module_name", Record.module_name, "Z");
                Rec.InsertNumeric("module_order", Record.module_order.ToString());

                if (Record.rec_mode == "ADD")
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);


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



    }
}
