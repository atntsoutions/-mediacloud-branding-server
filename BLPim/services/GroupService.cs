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
    public class GroupService : BL_Base
    {

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<tablesm> mList = new List<tablesm>();
            tablesm mRow;

            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string comp_code = SearchData["comp_code"].ToString();
            string type = SearchData["type"].ToString();
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
                    sWhere += "  tab_name like '%" + searchstring.ToLower() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  ";
                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total ";

                    sql += " FROM tablesm a ";
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
                sql += " select  tab_pkid,tab_name,tab_table_name, a.rec_created_by, a.rec_created_date ";
                sql += " ,row_number() over(order by tab_name) rn ";
                sql += " from  tablesm a  ";
                sql += " " + sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by tab_name ";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();


                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new tablesm();
                    mRow.tab_pkid = Dr["tab_pkid"].ToString();
                    mRow.tab_name = Dr["tab_name"].ToString();
                    mRow.tab_table_name = Dr["tab_table_name"].ToString();
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

            List<pim_groupm> mList = new List<pim_groupm>();
            pim_groupm mRow = new pim_groupm();

            string parentid = SearchData["parentid"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string table_name = SearchData["grp_table_name"].ToString();
            Boolean showInactive = (Boolean) SearchData["showinactive"];

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  grp_pkid, grp_name, grp_parent_id, grp_level_slno,grp_level_id, grp_level_name, rec_hidden, grp_table_name ";
                sql += " from pim_groupm a  ";
                sql += " where a.rec_company_code ='" + company_code + "' and a.grp_table_name = '" + table_name + "'";
                if  ( parentid == "")
                    sql += " and  a.grp_parent_id is null ";
                else 
                    sql += " and  a.grp_parent_id ='" + parentid + "'";

                if ( showInactive == false )
                    sql += " and  a.rec_hidden ='N'";


                sql += " order by grp_level_slno ";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new pim_groupm();
                    mRow.grp_pkid = Dr["grp_pkid"].ToString();
                    mRow.grp_name = Dr["grp_name"].ToString();
                    mRow.grp_parent_id = Dr["grp_parent_id"].ToString();

                    mRow.grp_level_id = Dr["grp_level_id"].ToString();
                    mRow.grp_level_name = Dr["grp_level_name"].ToString();
                    mRow.grp_table_name = Dr["grp_table_name"].ToString();


                    mRow.rec_hidden =  Dr["rec_hidden"].ToString() == "Y" ? true : false;
                    mList.Add(mRow);
                }
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

        public string AllValid( pim_groupm Record)
        {
            string str = "";
            try
            {
                sql = "select comp_pkid from (";
                sql += "select comp_pkid  from companym a where (a.comp_code = '{CODE}')  ";

                sql += ") a where comp_pkid <> '{PKID}'";

                //sql = sql.Replace("{CODE}", Record.comp_code);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";


            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(pim_groupm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            Boolean retvalue = false;

            DataTable dt_parent = null;
                   

            try
            {
                Con_Oracle = new DBConnection();

                if (Record.grp_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "name Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                if (Record.rec_type == "ROOT")
                {
                    Record.grp_level = 1;
                    Record.grp_parent_id = "";
                }
                else 
                {
                    sql = "select * from pim_groupm where grp_pkid = '" + Record.grp_parent_id + "'";
                    dt_parent = Con_Oracle.ExecuteQuery(sql);
                    if (dt_parent.Rows.Count <= 0)
                    {
                        if (ErrorMessage != "")
                            throw new Exception("Invalid Parent");
                    }
                    Record.grp_level = Lib.Conv2Integer(dt_parent.Rows[0]["grp_level"].ToString()) + 1;
                    Record.grp_parent_id  = dt_parent.Rows[0]["grp_pkid"].ToString();
                }

                if (Record.grp_level <= 0)
                {
                    throw new Exception("Invalid group level");
                }

                sql = "select max(grp_name) as grp_name from pim_groupm where ";
                sql += " rec_company_code = '" + Record._globalvariables.comp_code + "'";
                sql += " and grp_name = '" + Record.grp_name.ToString().ToLower() + "'";
                sql += " and grp_table_name = '" + Record.grp_table_name.ToString().ToLower() + "'";

                if (Record.grp_parent_id == "")
                    sql += " and grp_parent_id is null ";
                else
                    sql += " and grp_parent_id = '" + Record.grp_parent_id + "'";

                if ( Con_Oracle.IsRowExists(sql) )
                {
                    throw new Exception("Dupliate name not  allowed ");
                }



                sql = "select nvl(max(grp_level_slno), 100000) + 1  as slno from pim_groupm where ";
                
                if ( Con_Oracle.DB == "SQL")
                    sql = "select isnull(max(grp_level_slno), 100000) + 1  as slno from pim_groupm where ";

                sql += " rec_company_code = '" + Record._globalvariables.comp_code + "'";
                sql += " and grp_table_name = '" + Record.grp_table_name + "'";
                if (Record.grp_parent_id == "" )
                    sql += " and grp_parent_id is null ";
                else 
                    sql += " and grp_parent_id = '" + Record.grp_parent_id + "'";
                
                int iSlno = Lib.Conv2Integer( Con_Oracle.ExecuteScalar(sql).ToString());

                if (iSlno <= 0)
                {
                    throw new Exception("Invalid SL#");
                }


                if (Record.rec_type == "ROOT")
                {
                    Record.grp_level_id = iSlno.ToString();
                    Record.grp_level_name = Record.grp_name;
                }
                else
                {
                    Record.grp_level_id = dt_parent.Rows[0]["grp_level_id"].ToString() + "-" + iSlno.ToString();
                    Record.grp_level_name = dt_parent.Rows[0]["grp_level_name"].ToString() + "\\" + Record.grp_name;
                }


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("pim_groupm", Record.rec_mode, "grp_pkid", Record.grp_pkid);
                Rec.InsertString("grp_name", Record.grp_name, "L");

                Rec.InsertString("grp_parent_id", Record.grp_parent_id);
                Rec.InsertNumeric("grp_level", Record.grp_level.ToString());

                Rec.InsertString("grp_level_slno",iSlno.ToString());
                Rec.InsertString("grp_level_id", Record.grp_level_id);
                Rec.InsertString("grp_level_name", Record.grp_level_name, "L");

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("grp_table_name", Record.grp_table_name);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_hidden", "N");
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
            RetData.Add("grp_level_id", Record.grp_level_id);
            RetData.Add("grp_level_name", Record.grp_level_name);

            return RetData;
        }


        public Dictionary<string, object> Delete(Dictionary<string, object> SearchData)
        {
            Boolean bRet = false;
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            string pkid = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string table_name = SearchData["table_name"].ToString();

            sql = "select max(grp_pkid) as id from pim_groupm where grp_parent_id = '" + pkid  + "'";
            if ( Con_Oracle.IsRowExists(sql))
                throw new Exception("Sub Group Exists");

            sql = "select  grp_level_id  from pim_groupm where grp_pkid = '" + pkid + "'";
            DataTable dt_test = new DataTable();
            dt_test = Con_Oracle.ExecuteQuery(sql);
            if ( dt_test.Rows.Count <= 0)
                throw new Exception("Invalid Group");

            sql = "";
            sql += " select count(*) as tot from pim_groupm a  ";
            sql += " inner join pim_docm b on a.grp_pkid = b.doc_grp_id ";
            sql += " where grp_table_name = '" + table_name +"' and grp_level_id like '" + dt_test.Rows[0]["grp_level_id"].ToString()  +"%' ";

            dt_test = Con_Oracle.ExecuteQuery(sql);
            if ( Lib.Conv2Integer(dt_test.Rows[0]["tot"].ToString()) > 0)
                throw new Exception("Group Data Exists");

            sql = "delete from pim_groupm where grp_pkid = '" + pkid + "'";
            
            try
            {
                Con_Oracle.BeginTransaction();
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
        



    }
}
