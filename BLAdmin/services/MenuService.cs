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
    public class MenuService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Menum> mList = new List<Menum>();
            Menum  mRow;

            string comp_code = SearchData["comp_code"].ToString().ToUpper();
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
                sWhere = " where  a.rec_company_code ='" + comp_code + "' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(b.module_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or upper(a.menu_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or upper(a.menu_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or upper(a.menu_route1) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM menum  a  ";
                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM menum  a ";

                    sql += " inner join modulem b on menu_module_id = module_pkid ";
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
                sql += "  select  menu_pkid, a.menu_code, a.menu_name , module_name, menu_type, menu_route1,menu_route2, menu_order,  row_number() over(order by module_name, menu_name) rn ";
                sql += "  from menum a inner join modulem b on a.menu_module_id = b.module_pkid " + sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by module_name, a.menu_order,a.menu_name ";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Menum();
                    mRow.menu_pkid = Dr["menu_pkid"].ToString();
                    mRow.menu_code = Dr["menu_code"].ToString();
                    mRow.menu_name = Dr["menu_name"].ToString();
                    mRow.menu_route1 = Dr["menu_route1"].ToString();
                    mRow.menu_route2 = Dr["menu_route2"].ToString();
                    mRow.menu_type = Dr["menu_type"].ToString();
                    mRow.menu_order = Lib.Conv2Integer(Dr["menu_order"].ToString());
                    mRow.menu_module_name = Dr["module_name"].ToString();
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
            Menum mRow =new Menum();
            
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  a.menu_pkid, a.menu_code, a.menu_name, a.menu_route1,a.menu_route2, a.menu_type, menu_order, menu_module_id,menu_displayed ";
                sql += " from menum a  ";
                sql += " where  a.menu_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Menum();
                    mRow.menu_pkid = Dr["menu_pkid"].ToString();
                    mRow.menu_code = Dr["menu_code"].ToString();
                    mRow.menu_name = Dr["menu_name"].ToString();
                    mRow.menu_route1 = Dr["menu_route1"].ToString();
                    mRow.menu_route2 = Dr["menu_route2"].ToString();
                    mRow.menu_type = Dr["menu_type"].ToString();
                    mRow.menu_order = Lib.Conv2Integer(Dr["menu_order"].ToString());
                    mRow.menu_module_id = Dr["menu_module_id"].ToString();


                    mRow.menu_displayed = false;
                    if (Dr["menu_displayed"].ToString() == "Y")
                        mRow.menu_displayed = true;

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


        public string AllValid(Menum Record)
        {
            string str = "";
            try
            {
                sql = "select menu_pkid from (";
                sql += "select menu_pkid  from menum a where rec_company_code ='" + Record._globalvariables.comp_code + "' ";
                sql += " and (a.menu_code = '{CODE}' or a.menu_name = '{NAME}')  ";
                sql += ") a where menu_pkid <> '{PKID}'";

                sql = sql.Replace("{CODE}", Record.menu_code);
                sql = sql.Replace("{NAME}", Record.menu_name);
                sql = sql.Replace("{PKID}", Record.menu_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Menum Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if (Record.menu_code.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Code Cannot Be Empty");
                if (Record.menu_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Description Cannot Be Empty");
                if (Record.menu_module_id is null)
                    Lib.AddError(ref ErrorMessage, "Module Cannot Be Blank");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("menum", Record.rec_mode, "menu_pkid", Record.menu_pkid);
                Rec.InsertString("menu_code", Record.menu_code);
                Rec.InsertString("menu_name", Record.menu_name, "Z");
                Rec.InsertString("menu_route1", Record.menu_route1, "Z");
                Rec.InsertString("menu_route2", Record.menu_route2, "Z");
                Rec.InsertString("menu_type", Record.menu_type);
                Rec.InsertNumeric("menu_order", Record.menu_order.ToString());
                Rec.InsertString("menu_module_id", Record.menu_module_id);

                if ( Record.menu_displayed )
                    Rec.InsertString("menu_displayed", "Y");
                else 
                    Rec.InsertString("menu_displayed", "N");

                if ( Record.rec_mode == "ADD")
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

        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();
            List<Modulem> mList = new List<Modulem>();
            Modulem mRow;

            string comp_code = SearchData["comp_code"].ToString();

            try
            {
                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select module_pkid,module_name from modulem where rec_company_code='" + comp_code + "'";
                sql += " order by module_name";

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Modulem();
                    mRow.module_pkid = Dr["module_pkid"].ToString();
                    mRow.module_name = Dr["module_name"].ToString();
                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("modules", mList);

            return RetData;
        }



    }
}
