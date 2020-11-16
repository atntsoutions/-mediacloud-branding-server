using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase.Connections;

namespace BLMaster
{
    public class TableService : BL_Base
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
                    sWhere += "  tab_name like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM  tablesm a ";

                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM tablesm a ";

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
                sql += "  select  a.tab_pkid, tab_name, tab_table_name, tab_caption,  tab_id, tab_store, tab_group, tab_sku,tab_file,";
                sql += " a.rec_created_by, a.rec_created_date ";
                sql += " ,row_number() over(order by tab_name) rn ";
                sql += " from tablesm a ";
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
                    mRow.tab_caption = Dr["tab_caption"].ToString();

                    mRow.tab_id = Dr["tab_id"].ToString();
                    mRow.tab_store = Dr["tab_store"].ToString();
                    mRow.tab_group = Dr["tab_group"].ToString();
                    mRow.tab_sku = Dr["tab_sku"].ToString();
                    mRow.tab_file = Dr["tab_file"].ToString();


                    mRow.rec_created_by = Dr["rec_created_by"].ToString();
                    mRow.rec_created_date = Dr["rec_created_date"].ToString();
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

            tablesm mRow = new tablesm();

            List<tablesd> mList = new List<tablesd>();
            tablesd mRec = new tablesd();


            string pkid = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select tab_pkid, tab_name, tab_table_name, tab_caption, tab_id, tab_store, tab_group, tab_sku,tab_file, ";
                sql += " tab_sku_duplication, tab_store_duplication,tab_campaign_table ";
                sql += " from Tablesm a ";
                sql += " where tab_pkid = '" + pkid + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();


                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new tablesm();
                    mRow.tab_pkid = Dr["tab_pkid"].ToString();
                    mRow.tab_name = Dr["tab_name"].ToString();
                    mRow.tab_table_name = Dr["tab_table_name"].ToString();
                    mRow.tab_caption = Dr["tab_caption"].ToString();

                    mRow.tab_id = Dr["tab_id"].ToString();
                    mRow.tab_store = Dr["tab_store"].ToString();
                    mRow.tab_group = Dr["tab_group"].ToString();
                    mRow.tab_sku = Dr["tab_sku"].ToString();
                    mRow.tab_file = Dr["tab_file"].ToString();

                    mRow.tab_sku_duplication = Dr["tab_sku_duplication"].ToString() == "Y" ?  true : false ;
                    mRow.tab_store_duplication = Dr["tab_store_duplication"].ToString() == "Y" ? true : false;
                    mRow.tab_campaign_table = Dr["tab_campaign_table"].ToString() == "Y" ? true : false;


                    break;
                }

                Dt_Rec = new DataTable();

                sql = " select  b.* from ";
                sql += " tablesd b ";
                sql += " where b.tabd_parent_id = '" + pkid + "' ";
                sql += " order by tabd_col_order ";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRec = new tablesd();
                    mRec.tabd_pkid = Dr["tabd_pkid"].ToString();
                    mRec.tabd_parent_id = Dr["tabd_parent_id"].ToString();
                    mRec.tabd_tab_name = mRow.tab_name;
                    mRec.tabd_table_name = mRow.tab_table_name;
                    mRec.tabd_col_name = Dr["tabd_col_name"].ToString();
                    mRec.tabd_col_caption = Dr["tabd_col_caption"].ToString();
                    mRec.tabd_col_type = Dr["tabd_col_type"].ToString();
                    mRec.tabd_col_case = Dr["tabd_col_case"].ToString();
                    mRec.tabd_col_mandatory = Dr["tabd_col_mandatory"].ToString();

                    mRec.tabd_col_id = Dr["tabd_col_id"].ToString();
                    mRec.tabd_col_value = Dr["tabd_col_value"].ToString();
                    mRec.tabd_col_list = Dr["tabd_col_list"].ToString();

                    mRec.tabd_col_rows = Lib.Conv2Integer(Dr["tabd_col_rows"].ToString());
                    mRec.tabd_col_len = Lib.Conv2Integer(Dr["tabd_col_len"].ToString());
                    mRec.tabd_col_dec = Lib.Conv2Integer(Dr["tabd_col_dec"].ToString());
                    mRec.tabd_col_order = Lib.Conv2Integer(Dr["tabd_col_order"].ToString());

                    mRec.rec_deleted = ( Dr["rec_deleted"].ToString() == "Y") ? true : false;

                    mRec.rec_mode = "EDIT";
                    mList.Add(mRec);
                }


            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            RetData.Add("records", mList);
            return RetData;
        }

        public string AllValid( tablesm Record)
        {
            string str = "";
            try
            {
                sql = "select tab_pkid from (";
                sql += "  select tab_pkid from tablesm a ";
                sql += " where rec_company_Code = '" +  Record._globalvariables.comp_code + "'  and tab_name = '" + Record.tab_name + "' ";
                sql += ") a where tab_pkid <> '"+ Record.tab_pkid  +"'";

                //sql = sql.Replace("{CODE}", Record.comp_code);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Table Name Exists";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(tablesm Record )
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            Boolean retvalue = false;

            string sql1 = "";

            DBRecord Rec = null;

            int iOrder = 0;

            try
            {
                Con_Oracle = new DBConnection();

                if (Record.tab_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Table Name Cannot Be Empty");

                if (Record.tab_caption.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Caption Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);

                Rec = new DBRecord();
                Rec.CreateRow("tablesm", Record.rec_mode, "tab_pkid", Record.tab_pkid);
                Rec.InsertString("tab_caption", Record.tab_caption, "P");

                Rec.InsertString("tab_id", Record.tab_id, "P");
                Rec.InsertString("tab_store", Record.tab_store, "P");
                Rec.InsertString("tab_group", Record.tab_group, "P");
                Rec.InsertString("tab_sku", Record.tab_sku, "P");
                Rec.InsertString("tab_file", Record.tab_file, "P");

                Rec.InsertString("tab_sku_duplication", Record.tab_sku_duplication ? "Y" : "N");
                Rec.InsertString("tab_store_duplication", Record.tab_store_duplication ? "Y" : "N");
                Rec.InsertString("tab_campaign_table", Record.tab_campaign_table ? "Y" : "N");


                if (Record.rec_mode == "ADD")
                {
                    // Table Creation
                    Record.tab_table_name = "TBL_" + Record._globalvariables.comp_code + "_" + Record.tab_name.ToUpper();
                    Rec.InsertString("tab_name", Record.tab_name, "U");
                    Rec.InsertString("tab_table_name", Record.tab_table_name, "U");
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    if ( Con_Oracle.DB == "ORACLE")
                        Rec.InsertFunction("rec_created_date", "SYSDATE");
                    else
                        Rec.InsertFunction("rec_created_date", "getdate()");

                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    if (Con_Oracle.DB == "ORACLE")
                        Rec.InsertFunction("rec_edited_date", "SYSDATE");
                    else
                        Rec.InsertFunction("rec_edited_date", "getdate()");
                }

                if (Record.rec_mode == "ADD")
                {
                    sql1 = "";
                    if (Con_Oracle.DB == "ORACLE")
                    {
                        sql1 += " CREATE TABLE " + Record.tab_table_name;
                        sql1 += " (  ";
                        //sql1 += " DOC_PKID          NVARCHAR2(40), ";
                        sql1 += " DOC_PARENT_ID        NVARCHAR2(40), ";
                        //sql1 += " DOC_SLNO          NUMBER(15),    ";
                        //sql1 += " DOC_TABLE_NAME    NVARCHAR2(60), ";
                        //sql1 += " DOC_NAME          NVARCHAR2(100), ";
                        //sql1 += " DOC_FILE_NAME     NVARCHAR2(100), ";
                        sql1 += " REC_COMPANY_CODE  NVARCHAR2(10), ";
                        sql1 += " REC_CREATED_BY    NVARCHAR2(15), ";
                        sql1 += " REC_CREATED_DATE  DATE,          ";
                        sql1 += " REC_EDITED_BY     NVARCHAR2(15), ";
                        sql1 += " REC_EDITED_DATE   DATE ";
                        sql1 += " ) ";
                    }
                    else
                    {
                        sql1 += " CREATE TABLE " + Record.tab_table_name;
                        sql1 += " (  ";
                        //sql1 += " DOC_PKID          NVARCHAR(40), ";
                        sql1 += " DOC_PARENT_ID        NVARCHAR(40), ";
                        //sql1 += " DOC_SLNO          NUMERIC(15),    ";
                        //sql1 += " DOC_TABLE_NAME    NVARCHAR(60), ";
                        //sql1 += " DOC_NAME          NVARCHAR(100), ";
                        //sql1 += " DOC_FILE_NAME     NVARCHAR(100), ";
                        sql1 += " REC_COMPANY_CODE  NVARCHAR(10), ";
                        sql1 += " REC_CREATED_BY    NVARCHAR(15), ";
                        sql1 += " REC_CREATED_DATE  DATETIME,          ";
                        sql1 += " REC_EDITED_BY     NVARCHAR(15), ";
                        sql1 += " REC_EDITED_DATE   DATETIME ";
                        sql1 += " ) ";
                    }

                }
                sql = Rec.UpdateRow();


                Rec = new DBRecord();
                Rec.CreateRow("menum", Record.rec_mode, "menu_pkid", Record.tab_pkid);
                
                Rec.InsertString("menu_code", "~PIM~" + Record.tab_name.ToUpper());
                Rec.InsertString("menu_name", Record.tab_caption, "P");
                Rec.InsertString("menu_route1", "pim/pim", "P");
                string str = "urlid" + ":" + "PIM" + "menuid" + ":" + "PIM" + "," + "type" + ":" + "TBL_VTC_PRODUCT";
                str= "{ \"urlid\":\"{PIM}\",\"menuid\":\"{PIM}\",\"type\":\"{TBL}\"}";

                str = str.Replace("{PIM}", "~PIM~"+ Record.tab_name.ToUpper());
                str = str.Replace("{TBL}", Record.tab_table_name.ToUpper());

                Rec.InsertString("menu_route2",str, "P");
                Rec.InsertString("menu_type", "AEDP");
                Rec.InsertString("menu_displayed", "Y");


                if (Record.rec_mode == "ADD")
                {
                    string sql10 = "select nvl(max(menu_order),10) + 10 from menum where rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    if ( Con_Oracle.DB == "SQL" )
                        sql10 = "select isnull(max(menu_order),10) + 10 from menum where rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    iOrder = Lib.Conv2Integer( Con_Oracle.ExecuteScalar(sql10).ToString());

                    Rec.InsertString("menu_order", iOrder.ToString());
                    Rec.InsertString("menu_module_id", "3C784E47-4EC4-AF88-60F0-8555036656FA");
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                }
                string sql2 = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                if ( sql1 != "")
                    Con_Oracle.ExecuteNonQuery(sql1);
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.ExecuteNonQuery(sql2);

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
            RetData.Add("table_name", Record.tab_table_name);
            RetData.Add("retvalue",retvalue);

            return RetData;
        }


        public string AllValid2(tablesd Record)
        {
            string str = "";
            try
            {
                sql = "select tabd_pkid from (";
                sql += " select tabd_pkid from tablesd a ";
                sql += " where tabd_parent_id = '" + Record.tabd_parent_id + "' and tabd_col_name = '" + Record.tabd_col_name + "' ";
                sql += ") a where tabd_pkid <> '" + Record.tabd_pkid + "'";

                //sql = sql.Replace("{CODE}", Record.comp_code);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Table Name Exists";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }



        public Dictionary<string, object> SaveDetail(tablesd Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            Boolean retvalue = false;

            string sql1 = "";

            DBRecord Rec = null;

            string flag = "";

            int iOrder = 0;

            try
            {
                Con_Oracle = new DBConnection();

                if (Record.tabd_col_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Column Name Cannot Be Empty");

                if (Record.tabd_col_caption.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Column Caption Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                Record.tabd_col_name = Record.tabd_col_name.ToUpper().Replace(" ", "");

                if ((ErrorMessage = AllValid2(Record)) != "")
                    throw new Exception(ErrorMessage);

                if (Record.rec_mode == "ADD" && Record.tabd_col_order <=0)
                {
                    sql = "select nvl(max(tabd_col_order),10) + 10 from tablesd where rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    if ( Con_Oracle.DB == "SQL")
                        sql = "select isnull(max(tabd_col_order),10) + 10 from tablesd where rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    sql += " and tabd_parent_id ='" + Record.tabd_parent_id + "'";
                    iOrder = Lib.Conv2Integer(Con_Oracle.ExecuteScalar(sql).ToString());
                }

                    Rec = new DBRecord();
                Rec.CreateRow("tablesd", Record.rec_mode, "tabd_pkid", Record.tabd_pkid);
                Rec.InsertString("tabd_col_caption", Record.tabd_col_caption, "P");
                Rec.InsertString("tabd_col_type", Record.tabd_col_type, "U");
                Rec.InsertString("tabd_col_case", Record.tabd_col_case, "U");
                Rec.InsertString("tabd_col_mandatory", Record.tabd_col_mandatory, "U");

                Rec.InsertString("tabd_col_id", Record.tabd_col_id, "P");
                Rec.InsertString("tabd_col_value", Record.tabd_col_value, "P");
                Rec.InsertString("tabd_col_list", Record.tabd_col_list, "P");

                

                Rec.InsertNumeric("tabd_col_rows", Record.tabd_col_rows.ToString());
                Rec.InsertNumeric("tabd_col_len", Record.tabd_col_len.ToString());
                Rec.InsertNumeric("tabd_col_dec", Record.tabd_col_dec.ToString());


                Rec.InsertString("rec_deleted", (Record.rec_deleted) ? "Y" : "N");

                if (Record.rec_mode == "ADD")
                {
                    if ( Record.tabd_col_order > 0)
                        Rec.InsertNumeric("tabd_col_order", Record.tabd_col_order.ToString());
                    else
                        Rec.InsertNumeric("tabd_col_order", iOrder.ToString());

                    Rec.InsertString("tabd_parent_id", Record.tabd_parent_id);
                    Rec.InsertString("tabd_col_name", Record.tabd_col_name, "U");

                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);

                    if ( Con_Oracle.DB == "ORACLE")
                        Rec.InsertFunction("rec_created_date", "SYSDATE");
                    else
                        Rec.InsertFunction("rec_created_date", "getdate()");

                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertNumeric("tabd_col_order", Record.tabd_col_order.ToString());
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    if (Con_Oracle.DB == "ORACLE")
                        Rec.InsertFunction("rec_edited_date", "SYSDATE");
                    else
                        Rec.InsertFunction("rec_edited_date", "getdate()");
                }


                if (Record.rec_mode == "ADD")
                    flag = " ADD ";

                if (Con_Oracle.DB == "ORACLE")
                {
                    if (Record.rec_mode == "EDIT")
                        flag = " MODIFY ";
                }
                else
                {
                    if (Record.rec_mode == "EDIT")
                        flag = " ALTER COLUMN ";
                }


                sql1 = "";

                if (Record.tabd_col_type == "TEXT" || Record.tabd_col_type == "MEMO")
                {
                    if ( Con_Oracle.DB =="ORACLE")
                        sql1 += "COL_" + Record.tabd_col_name + " nvarchar2(" + Record.tabd_col_len + ")";
                    else 
                        sql1 += "COL_" + Record.tabd_col_name + " nvarchar(" + Record.tabd_col_len + ")";
                }

                if (Record.tabd_col_type == "NUMBER")
                {
                    if (Con_Oracle.DB == "ORACLE")
                        sql1 += "COL_" + Record.tabd_col_name + " number(" + Record.tabd_col_len + "," + Record.tabd_col_dec + ")";
                    else
                        sql1 += "COL_" + Record.tabd_col_name + " numeric(" + Record.tabd_col_len + "," + Record.tabd_col_dec + ")";
                }

                if (Record.tabd_col_type == "DATE")
                {
                    if (Con_Oracle.DB == "ORACLE")
                        sql1 += "COL_" + Record.tabd_col_name + " DATE";
                    else
                        sql1 += "COL_" + Record.tabd_col_name + " DATETIME";
                }

                if (Record.tabd_col_type == "LIST")
                {
                    if (Con_Oracle.DB == "ORACLE")
                        sql1 += "COL_" + Record.tabd_col_name + " nvarchar2(40)";
                    else
                        sql1 += "COL_" + Record.tabd_col_name + " nvarchar(40)";
                }

                if (Record.tabd_col_type == "FILE")
                {
                    if (Con_Oracle.DB == "ORACLE")
                        sql1 += "COL_" + Record.tabd_col_name + " nvarchar2(100)";
                    else
                        sql1 += "COL_" + Record.tabd_col_name + " nvarchar(100)";
                }

                if ( sql1 != "")
                    sql1 = " ALTER TABLE " + Record.tabd_table_name.ToUpper() + " " + flag + " " + sql1;

                sql = Rec.UpdateRow();
                Con_Oracle.BeginTransaction();
                if (sql1 != "")
                    Con_Oracle.ExecuteNonQuery(sql1);
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
            RetData.Add("col_name", Record.tabd_col_name);
            RetData.Add("iorder", iOrder);

            return RetData;
        }





        public Dictionary<string, object> Delete(Dictionary<string, object> SearchData)
        {
            Boolean bRet = false;
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            string pkid = SearchData["pkid"].ToString();

            sql = "delete from tablesm  where tab_pkid = '" + pkid + "'";
            
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



        public Dictionary<string, object> DeleteDetail(Dictionary<string, object> SearchData)
        {
            Boolean bRet = false;
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            string pkid = SearchData["pkid"].ToString();
            string col_name = SearchData["col_name"].ToString();
            string table_name = SearchData["table_name"].ToString();

            string sql1 = "ALTER TABLE " + table_name + " drop column col_" + col_name ;

            sql = "delete from tablesd where tabd_pkid = '" + pkid + "'";

            try
            {
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.ExecuteNonQuery(sql1);
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
            RetData.Add("flag", bRet);
            return RetData;
        }


        public Dictionary<string, object> DeleteTable (Dictionary<string, object> SearchData)
        {
            Boolean bRet = false;
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            string pkid = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string table_name = SearchData["table_name"].ToString();

            string str = ReadColumns(pkid);
            if ( str.Length >0 )
                throw new Exception("Table Column Exists");

            string sql1 = "DROP TABLE " + table_name ;

           sql = "delete from tablesm where tab_pkid = '" + pkid + "'";

           string sql2 = "delete from menum where menu_pkid = '" + pkid + "'";

           string sql3 = "delete from pim_groupm where grp_table_name = '" + table_name + "'";

            try
            {
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql1);
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.ExecuteNonQuery(sql2);
                Con_Oracle.ExecuteNonQuery(sql3);
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
            RetData.Add("flag", bRet);
            return RetData;
        }






        public string ReadColumns( string id )
        {
            string str = "";

            sql = " select tabd_col_name from tablesd where tabd_parent_id ='" + id + "' order by tabd_col_order ";
            Con_Oracle = new DBConnection();
            DataTable Dt_test = new DataTable();

            try
            {
                Dt_test = Con_Oracle.ExecuteQuery(sql);
                foreach ( DataRow Dr in Dt_test.Rows)
                {
                    str += (str != "") ? "," : str;
                    str += "COL_" + Dr["tabd_col_name"].ToString();
                }

                Con_Oracle.CloseConnection();

            }
            catch ( Exception ex)
            {

                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                str = "";
                throw ex;

            }

            return str;

        }

    }

 
}
