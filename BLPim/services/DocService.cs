using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase.Connections;
using System.Xml;

namespace BLPim
{
    public class DocService : BL_Base
    {
        DataTable dt_tree = new DataTable();
        XmlDocument xmlDoc = new XmlDocument();
        int iCounter = -1;


        public IDictionary<string, object> List(Dictionary<string, object> SearchData, string ServerImagURL)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<pim_docm> mList = new List<pim_docm>();
            pim_docm mRow;

            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string table_name = SearchData["grp_table_name"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string type = SearchData["type"].ToString();
            string cap_store = SearchData["cap_store"].ToString();
            string user_id = SearchData["user_id"].ToString();
            Boolean user_admin  = (Boolean)SearchData["user_admin"];
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  a.rec_company_code ='" + comp_code + "' ";
                sWhere += " and doc_table_name = '" + table_name + "'"; ;
                
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  doc_name like '%" + searchstring.ToLower() + "%'";
                    if (cap_store != "")
                    {
                        sWhere += " or comp_name like '%" + searchstring.ToLower() + "%'";
                    }
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  ";
                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total ";

                    sql += " FROM pim_docm a ";
                    sql += " left join pim_groupm c on a.doc_grp_id  = c.grp_pkid ";
                    sql += " left join companym   d on a.doc_store_id = d.comp_pkid ";

                    if (cap_store != "" && !user_admin)
                        sql += " inner join userd e on e.rec_type = 'S' and a.doc_store_id = e.user_branch_id and e.user_id ='" + user_id + "'";


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
                sql += " select  a.doc_pkid,doc_slno,doc_name,doc_table_name, doc_file_name, d.comp_name as  store_name, grp_name, grp_level_name, doc_thumbnail, a.rec_created_by, a.rec_created_date ";
                sql += " ,row_number() over(order by doc_slno) rn ";
                sql += " from  pim_docm a  ";
                sql += " left join pim_groupm c on a.doc_grp_id = c.grp_pkid ";
                sql += " left join companym   d on a.doc_store_id = d.comp_pkid ";
                if (cap_store != "" && !user_admin)
                    sql += " inner join userd e on e.rec_type = 'S' and a.doc_store_id = e.user_branch_id and e.user_id ='" + user_id + "'";
                sql += " " + sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by doc_slno ";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new pim_docm();
                    mRow.doc_pkid = Dr["doc_pkid"].ToString();
                    mRow.doc_slno = Lib.Conv2Integer(Dr["doc_slno"].ToString());
                    mRow.doc_name = Dr["doc_name"].ToString();
                    mRow.doc_store_name = Dr["store_name"].ToString();
                    mRow.doc_file_name = Dr["doc_file_name"].ToString();
                    mRow.doc_grp_level_name = Dr["grp_level_name"].ToString();
                    mRow.doc_table_name = Dr["doc_table_name"].ToString();
                    mRow.rec_created_by = Dr["rec_created_by"].ToString();
                    mRow.rec_created_date = Dr["rec_created_date"].ToString();
                    mRow.doc_server_folder = Lib.getPath(ServerImagURL, comp_code, table_name, mRow.doc_slno.ToString(), false);
                    mRow.doc_thumbnail =  Dr["doc_thumbnail"].ToString();
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
        
        public Dictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            List<tablesd> mList = new List<tablesd>();
            tablesd mRec = new tablesd();

            tablesm Rec = new tablesm();

            string comp_code = SearchData["comp_code"].ToString();
            string table_name = SearchData["table_name"].ToString();

            try
            {

                Con_Oracle = new DBConnection();

                DataTable Dt_First = new DataTable();
                sql = " select tab_name, tab_table_name, tab_id, tab_store, tab_group, tab_sku,tab_file, tab_store_duplication from ";
                sql += " tablesm a ";
                sql += " where a.rec_company_code = '" + comp_code + "' and tab_table_name = '" + table_name + "'";
                Dt_First = Con_Oracle.ExecuteQuery(sql);

                Rec.tab_id = "";
                Rec.tab_store = "";
                Rec.tab_group = "";
                Rec.tab_sku = "";
                Rec.tab_file = "";
                

                if ( Dt_First.Rows.Count > 0)
                {
                    Rec.tab_id = Dt_First.Rows[0]["tab_id"].ToString();
                    Rec.tab_store = Dt_First.Rows[0]["tab_store"].ToString();
                    Rec.tab_group = Dt_First.Rows[0]["tab_group"].ToString();
                    Rec.tab_sku = Dt_First.Rows[0]["tab_sku"].ToString();
                    Rec.tab_file = Dt_First.Rows[0]["tab_file"].ToString();
                    Rec.tab_store_duplication = (Dt_First.Rows[0]["tab_store_duplication"].ToString() == "Y") ? true : false ;
                
                }


                DataTable Dt_Rec = new DataTable();
                sql = " select tab_name, tab_table_name, b.* from ";
                sql += " tablesm a inner join tablesd b on a.tab_pkid = tabd_parent_id ";
                sql += " where a.rec_company_code = '" + comp_code + "' and tab_table_name = '" + table_name + "'  and b.rec_deleted = 'N'";
                sql += " order by tabd_col_order ";

                
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRec = new tablesd();
                    mRec.tabd_table_name = Dr["tab_table_name"].ToString();
                    mRec.tabd_col_name = Dr["tabd_col_name"].ToString();
                    mRec.tabd_col_caption = Dr["tabd_col_caption"].ToString();
                    mRec.tabd_col_type = Dr["tabd_col_type"].ToString();
                    mRec.tabd_col_mandatory = Dr["tabd_col_mandatory"].ToString();
                    mRec.tabd_col_id = Dr["tabd_col_id"].ToString();
                    mRec.tabd_col_value = Dr["tabd_col_value"].ToString();
                    mRec.tabd_col_list = Dr["tabd_col_list"].ToString();
                    mRec.tabd_col_len = Lib.Conv2Integer(Dr["tabd_col_len"].ToString());
                    mRec.tabd_col_dec = Lib.Conv2Integer(Dr["tabd_col_dec"].ToString());
                    mRec.tabd_col_order = Lib.Conv2Integer(Dr["tabd_col_order"].ToString());

                    mRec.tabd_col_file_uploaded = false;

                    if ( Dr["tabd_col_type"].ToString() == "LIST")
                    {
                        mRec.tabd_col_value = "";
                        mRec.tabd_col_id = "";
                    }


                    mList.Add(mRec);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("list", mList);
            RetData.Add("tablesm",Rec );
            return RetData;
        }

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData, string ServerImageUrl)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            DataRow DROW = null;

            pim_docm mRow = new pim_docm();

            List<tablesd> mList = new List<tablesd>();
            tablesd mRec = new tablesd();

            LovService Lov = new LovService();



            string pkid = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string table_name = SearchData["table_name"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                string str = ReadColumns(table_name, comp_code);

                sql = " select grp_name, doc_table_name,doc_pkid,doc_grp_id,grp_level_name, doc_store_id, d.comp_name as store_name,  ";
                sql +=  " doc_slno, doc_name, doc_file_name, doc_thumbnail ";
                if (str.Length > 0)
                    sql += "," + str;
                sql += " from pim_docm a ";
                sql += " left join " + table_name + " b on a.doc_pkid = b.doc_parent_id ";
                sql += " left join pim_groupm c on a.doc_grp_id = c.grp_pkid ";
                sql += " left join companym   d on a.doc_store_id = d.comp_pkid ";
                sql += " where doc_pkid = '" + pkid + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                if ( Dt_Rec.Rows.Count > 0)
                    DROW = Dt_Rec.Rows[0];

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new pim_docm();
                    mRow.doc_pkid = Dr["doc_pkid"].ToString();
                    mRow.doc_slno =  Lib.Conv2Integer(Dr["doc_slno"].ToString());

                    mRow.doc_store_id = Dr["doc_store_id"].ToString();
                    mRow.doc_store_name = Dr["store_name"].ToString();

                    mRow.doc_grp_id = Dr["doc_grp_id"].ToString();
                    mRow.doc_grp_level_name = Dr["grp_level_name"].ToString();
                    mRow.doc_name = Dr["doc_name"].ToString();
                    mRow.doc_file_name = Dr["doc_file_name"].ToString();
                    mRow.doc_table_name = Dr["doc_table_name"].ToString();
                    mRow.doc_thumbnail = Dr["doc_thumbnail"].ToString();
                    mRow.doc_server_folder = Lib.getPath(ServerImageUrl, comp_code, table_name, mRow.doc_slno.ToString(), false);
                    mRow.doc_file_uploaded = false;

                    break;
                }

                Dt_Rec = new DataTable();

                sql = " select tab_name, tab_table_name, b.* from ";
                sql += " tablesm a inner join tablesd b on a.tab_pkid = tabd_parent_id ";
                sql += " where a.rec_company_code = '" + comp_code + "' and tab_table_name = '" + table_name + "' and b.rec_deleted ='N'";
                sql += " order by tabd_col_order ";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRec = new tablesd();
                    mRec.tabd_table_name = Dr["tab_table_name"].ToString();
                    mRec.tabd_col_name = Dr["tabd_col_name"].ToString();
                    mRec.tabd_col_caption = Dr["tabd_col_caption"].ToString();
                    mRec.tabd_col_type = Dr["tabd_col_type"].ToString();

                    mRec.tabd_col_case = Dr["tabd_col_case"].ToString();
                    mRec.tabd_col_mandatory = Dr["tabd_col_mandatory"].ToString();

                    mRec.tabd_col_rows = Lib.Conv2Integer(Dr["tabd_col_rows"].ToString());

                    if (Dr["tabd_col_type"].ToString() == "DATE")
                        mRec.tabd_col_value = Lib.DatetoString(DROW["COL_" + Dr["tabd_col_name"].ToString()]);
                    else if (Dr["tabd_col_type"].ToString() == "LIST")
                    {
                        mRec.tabd_col_list = Dr["tabd_col_list"].ToString();
                        mRec.tabd_col_id = DROW["COL_" + Dr["tabd_col_name"].ToString()].ToString();
                        mRec.tabd_col_value = "";
                        if (Dr["tabd_col_id"].ToString().Length > 0)
                            mRec.tabd_col_value = Lov.getParamValue(comp_code, DROW["COL_" + Dr["tabd_col_name"].ToString()].ToString(), "PARAM_NAME");
                    }
                    else
                         mRec.tabd_col_value = DROW["COL_" + Dr["tabd_col_name"].ToString()].ToString();


                    mRec.tabd_col_len = Lib.Conv2Integer(Dr["tabd_col_len"].ToString());
                    mRec.tabd_col_dec = Lib.Conv2Integer(Dr["tabd_col_dec"].ToString());
                    mRec.tabd_col_order = Lib.Conv2Integer(Dr["tabd_col_order"].ToString());

                    mRec.tabd_col_file_uploaded = false;


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
            RetData.Add("list", mList);
            return RetData;
        }

        public string AllValid( pim_docm Record, DataRow HeaderRow)
        {
            string str = "";
            try
            {

                if (HeaderRow["tab_sku"].ToString() != "" && HeaderRow["tab_sku_duplication"].ToString() == "N")
                {
                    sql = "select doc_pkid from (";
                    sql += " select doc_pkid from pim_docm a ";
                    sql += "  where doc_name = '" + Record.doc_name + "' ";
                    sql += "  and doc_table_name  = '" + Record.doc_table_name + "' ";
                    sql += " ) a where doc_pkid <> '" + Record.doc_pkid + "'";
                    if (Con_Oracle.IsRowExists(sql))
                        str = HeaderRow["tab_sku"].ToString() +  " Exists";
                }
                else if (HeaderRow["tab_store"].ToString() != "" &&  HeaderRow["tab_store_duplication"].ToString() == "N")
                {
                    sql = "select doc_pkid from (";
                    sql += " select doc_pkid from pim_docm a ";
                    sql += "  where doc_store_id = '" + Record.doc_store_id + "' ";
                    sql += "  and doc_table_name  = '" + Record.doc_table_name + "' ";
                    sql += " ) a where doc_pkid <> '" + Record.doc_pkid + "'";
                    if (Con_Oracle.IsRowExists(sql))
                        str = HeaderRow["tab_store"].ToString() + " Exists";
                }
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(pim_docm Record, tablesd [] Records, string ServerImageUrl)
        {
            DataTable Dt_Tablesm = new DataTable();
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            Boolean retvalue = false;

            DBRecord Rec = null;

            Boolean is_campaign_table = false;

            int iSlno = 0;


            try
            {
                Con_Oracle = new DBConnection();

                sql = "select * from tablesm where rec_company_code = '"+ Record._globalvariables.comp_code  + "' and tab_table_name = '"+ Record.doc_table_name + "'";
                Dt_Tablesm = Con_Oracle.ExecuteQuery(sql);

                if (Dt_Tablesm.Rows.Count > 0)
                {
                    is_campaign_table = (Dt_Tablesm.Rows[0]["tab_campaign_table"].ToString() == "Y") ? true : false;
                }

                if ( Dt_Tablesm.Rows.Count <=0)
                    Lib.AddError(ref ErrorMessage, "Entity Not Found");

                
                //if (Record.doc_name.Trim().Length <= 0)
                //  Lib.AddError(ref ErrorMessage, "name Cannot Be Empty");


                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record, Dt_Tablesm.Rows[0])) != "")
                    throw new Exception(ErrorMessage);


                if (Record.rec_mode == "ADD")
                {
                    sql = "select nvl(max(doc_slno), 1000) + 1  as slno from pim_docm ";
                    if (Con_Oracle.DB == "SQL")
                        sql = "select isnull(max(doc_slno), 1000) + 1  as slno from pim_docm ";
                    sql += " where rec_company_code = '" + Record._globalvariables.comp_code + "' and doc_table_name ='" + Record.doc_table_name +  "'";
                    //sql += " and doc_table_name = '" + Record.doc_table_name + "'";

                    iSlno = Lib.Conv2Integer(Con_Oracle.ExecuteScalar(sql).ToString());

                    Record.doc_slno = iSlno;

                    if (iSlno <= 0)
                    {
                        throw new Exception("Invalid SL#");
                    }
                }
                else
                {
                    iSlno = Record.doc_slno;
                }

                sql = "";
                string sql1 = "";


                Rec = new DBRecord();
                Rec.CreateRow("pim_docm", Record.rec_mode, "doc_pkid", Record.doc_pkid);
                Rec.InsertString("doc_name", Record.doc_name, "P");
                Rec.InsertString("doc_file_name", Record.doc_file_name, "P");
                Rec.InsertString("doc_store_id", Record.doc_store_id);
                Rec.InsertString("doc_grp_id", Record.doc_grp_id);

                if ( Record.doc_file_name.Trim().Trim().Length <= 0 )
                    Rec.InsertString("doc_thumbnail", "");
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("doc_slno", Record.doc_slno.ToString());
                    Rec.InsertString("doc_table_name", Record.doc_table_name);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
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


                Rec = new DBRecord();
                Rec.CreateRow(Record.doc_table_name, Record.rec_mode, "doc_parent_id", Record.doc_pkid);
                foreach ( tablesd mRow in Records)
                {
                    if ( mRow.tabd_col_type == "DATE")
                        Rec.InsertDate("COL_" + mRow.tabd_col_name, mRow.tabd_col_value);
                    else if (mRow.tabd_col_type == "LIST")
                        Rec.InsertString("COL_" + mRow.tabd_col_name, mRow.tabd_col_id, "P");
                    else 
                        Rec.InsertString( "COL_" + mRow.tabd_col_name ,mRow.tabd_col_value, "P");
                }

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    if ( Con_Oracle.DB == "ORACLE")
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
                
                sql1 = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.ExecuteNonQuery(sql1);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                try
                {
                    if (is_campaign_table == false)
                    {
                        google_uploader g = new google_uploader();
                        g.bSingle = true;
                        g.comp_code = Record._globalvariables.comp_code;
                        g.user_id = Record._globalvariables.user_pkid;
                        string str = g.Process(Record.doc_table_name, "name");
                        if (str != "")
                            g.UploadData(Record.doc_pkid);
                    }
                }
                catch ( Exception )
                {

                }

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
            RetData.Add("slno", iSlno);

            string server = Lib.getPath ( ServerImageUrl, Record._globalvariables.comp_code, Record.doc_table_name,Record.doc_slno.ToString(),false);
            RetData.Add("server", server);
            RetData.Add("thumbnail", Record.doc_thumbnail);

            return RetData;
        }

        public Boolean UpdateDocFileName(pim_docm Record, string fldName, string value = "")
        {
            Boolean bRet = false;

            Con_Oracle = new DBConnection();
            
            if (fldName.ToUpper() == "DOC_FILE_NAME" || fldName.ToUpper() == "DOC_THUMBNAIL")
                sql = "update pim_docm set " + fldName + " = '"+ value  + "' where doc_pkid = '" + Record.doc_pkid + "'";
            else 
                sql = "update " + Record.doc_table_name + " set " + fldName + " = '" + value +"' where doc_parent_id = '" + Record.doc_pkid + "'";

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
                    Con_Oracle.CloseConnection();
                bRet = false;
                throw Ex;
            }
            return bRet;
        }

        public DataTable getDataTableRecord(string pkid)
        {
            Con_Oracle = new DBConnection();

            sql = "select doc_pkid,doc_table_name, doc_file_name, doc_thumbnail from pim_docm where doc_pkid = '" + pkid + "'";
            DataTable Dt = new DataTable();

            try
            {
                Con_Oracle.BeginTransaction();
                Dt= Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                Dt = null;
                throw Ex;
            }
            return Dt;
        }

        public string ReadColumns(string table_name, string comp_code)
        {
            string str = "";

            sql = " select tabd_col_name from tablesm a inner join tablesd b on a.tab_pkid = b.tabd_parent_id ";
            sql += " where a.rec_company_code = '" + comp_code + "' and tab_table_name ='" + table_name + "' and b.rec_deleted='N' ";
            sql += " order by tabd_col_order ";
            Con_Oracle = new DBConnection();
            DataTable Dt_test = new DataTable();

            try
            {
                Dt_test = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_test.Rows)
                {
                    str += (str != "") ? "," : str;
                    str += "COL_" + Dr["tabd_col_name"].ToString();
                }

                Con_Oracle.CloseConnection();

            }
            catch (Exception ex)
            {

                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                str = "";
                throw ex;

            }

            return str;

        }
        
        public IDictionary<string, object> Delete(Dictionary<string, object> SearchData, string serverPath)
        {
            Boolean bRet = false;
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            string pkid = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string table_name = SearchData["table_name"].ToString();

            sql = "";
            sql += "select doc_pkid, doc_slno, doc_table_name, doc_file_name, doc_thumbnail ";
            sql += " from pim_docm ";
            sql += " where doc_pkid = '" +  pkid + "'";

            DataTable Dt_test = new DataTable();
            Dt_test = Con_Oracle.ExecuteQuery(sql);

            DataRow Dr = null;
            if (Dt_test.Rows.Count <= 0)
                throw new Exception("Cannot Locate Entity Record");

            Dr = Dt_test.Rows[0];

            string Folder = Lib.getPath(serverPath, comp_code, Dr["doc_table_name"].ToString(), Dr["doc_slno"].ToString(),false);
            try
            {
                /*
                if (Doc_File_Name != "")
                    Lib.RemoveFile(sFileName1, true);
                if (Doc_File_Name != "")
                    Lib.RemoveFile(sFileName2,true);
                */
                Lib.RemoveFolder(Folder,true);

                Con_Oracle.BeginTransaction();

                sql = "delete from " + table_name + " where doc_parent_id = '" + pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                sql = "delete from pim_docm where doc_pkid = '" + pkid + "'";
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

        public void CreateTree(string comp_code, string grp_table_name, string grp_level_id)
        {

            sql = "";
            sql += " select grp_pkid, grp_parent_id, grp_level, grp_name, grp_level_slno, grp_level_id, grp_level_name   from pim_groupm a ";
            sql += " where  a.rec_company_code = '" + comp_code + "'";
            sql += " and grp_table_name = '" + grp_table_name + "' and grp_level_id like '" + grp_level_id + "%' ";
            sql += " order by grp_level_id ";

            // create function
            sql = "";
            sql += " CREATE function SelectChild(@key as nvarchar(40)) ";
            sql += " returns xml  ";
            sql += " begin ";
            sql += "   return (  ";
            sql += "     select  ";
            sql += "     grp_slno as '@id',  ";
            sql += "     grp_name as '@name',  ";
            sql += "     dbo.SelectChild(grp_pkid)  ";
            sql += "     from pim_groupm  ";
            sql += "     where grp_parent_id = @key  ";
            sql += "     for xml path('item'), type  ";
            sql += "   )  ";
            sql += " end  ";

            // xml query
            sql = "";
            sql += " SELECT ";
            sql += " grp_slno as '@id', ";
            sql += " grp_name as '@name', ";
            sql += " dbo.SelectChild(grp_pkid) ";
            sql += " from(select * from pim_groupm a where rec_company_code='" + comp_code + "' and grp_table_name = '" + grp_table_name + "' and grp_level_id like '" + grp_level_id + "%') as a ";
            sql += " where grp_parent_id is null ";
            sql += " FOR XML PATH('item'), root('directories') ";

            Con_Oracle = new DBConnection();
            XmlDocument xDoc = Con_Oracle.ExecuteXmlReader(sql);
            int RowCount = dt_tree.Rows.Count - 1;

            //xmlDoc.Save(@"c:\test.xml");
        }

        public XmlNode ProcessTree(XmlNode parent_node, int RowCount)
        {
            XmlNode node;
            iCounter++;

            string id = dt_tree.Rows[iCounter]["grp_level_id"].ToString();
            string name = dt_tree.Rows[iCounter]["grp_name"].ToString();
            int cur_level = int.Parse(dt_tree.Rows[iCounter]["grp_level"].ToString());
            int next_level = 0;

            try
            {
                if (iCounter < RowCount)
                    next_level = int.Parse(dt_tree.Rows[iCounter + 1]["grp_level"].ToString());

                XmlNode cur_node = addNode(id, name, cur_level);
                parent_node.AppendChild(cur_node);

                if (iCounter == RowCount)
                    return parent_node;

                if (next_level > cur_level)
                    node = cur_node;
                else if (next_level == cur_level)
                    node = parent_node;
                else
                    node = parent_node;

                return ProcessTree(node, RowCount);

            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public XmlNode addNode(string id, string name, int cur_Level)
        {
            XmlNode node = xmlDoc.CreateElement("item");

            XmlAttribute attr = xmlDoc.CreateAttribute("id");
            attr.Value = id;
            node.Attributes.Append(attr);

            attr = xmlDoc.CreateAttribute("name");
            attr.Value = name;
            node.Attributes.Append(attr);

            return node;
        }

        public IDictionary<string, object> Download(Dictionary<string, object> SearchData, string ServerImagePath, string ServerReportPath)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            string sname = "";

            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string grp_level_id = SearchData["grp_level_id"].ToString();
            string grp_table_name = SearchData["grp_table_name"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string type = SearchData["type"].ToString();

            string folderid = System.Guid.NewGuid().ToString().ToUpper();


            string frontEndDataFormat = "dd-mm-yyyy";

            int iCtr = 0;

            string sqlSelect = "";
            string sqlFrom = "";
            string sqlWhere = "";

            string fileName_xml1 = "";
            string fileName_xml2 = "";
            string fileName_csv = "";

            string pathName = "";

            try
            {
                DataTable Dt_Tables = new DataTable();
                DataTable Dt_List = new DataTable();

                frontEndDataFormat = Lib.getSettings(comp_code, "FRONTEND-DATE-FORMAT", "name");
                string Folder = Path.Combine(ServerReportPath, System.DateTime.Today.ToString("yyyy-MM-dd"), folderid);
                Lib.CreateFolder(Folder);

                /*

                sql = "";
                sql += " SELECT ";
                sql += " grp_slno as '@id', ";
                sql += " grp_name as '@name', ";
                sql += " dbo.SelectChild(grp_pkid) ";
                sql += " from( ";
                sql += "      select * from pim_groupm a where rec_company_code='" + comp_code + "' ";
                sql += "      and grp_table_name = '" + grp_table_name + "' ";
                sql += "      and grp_level_id like '" + grp_level_id + "%'";
                sql += " ) as a ";
                sql += " where grp_parent_id is null ";
                sql += " FOR XML PATH('item'), root('directories') ";

                Con_Oracle = new DBConnection();

                XmlDocument xDoc = Con_Oracle.ExecuteXmlReader(sql);
                fileName_xml1 = Folder + "\\" + grp_table_name + "1.xml";
                xDoc.Save(fileName_xml1);


                sql = "";
                sql += " select ";
                sql += " grp_slno as '@grpid', ";
                sql += " doc_slno as '@id',  ";
                sql += " doc_name as '@name',  ";
                sql += " '' as '@remoteURL', ";
                sql += " doc_file_name as '@highResPdfURL', ";
                sql += " doc_thumbnail as '@thumb', ";
                sql += " 'true' as '@keepExternal', ";
                sql += " 'true' as '@accessibleFromClient', ";
                sql += " '' as 'fileInfo/@width',  ";
                sql += " '' as 'fileInfo/@height' , ";
                sql += " '' as 'fileInfo/@resolution', ";
                sql += " '' as 'fileInfo/@fileSize' ";
                sql += " from pim_groupm a ";
                sql += " inner join pim_docm b on a.grp_pkid = b.doc_grp_id ";
                sql += " where a.rec_company_code = '" + comp_code + "' ";
                sql += " and doc_table_name = '" + grp_table_name + "' ";
                sql += " and grp_level_id like '" + grp_level_id + "%'";
                sql += " for xml path('item'), root('assets'), type ";

                xDoc = Con_Oracle.ExecuteXmlReader(sql);
                fileName_xml2 = Folder + "\\" + grp_table_name + "2.xml";
                xDoc.Save(fileName_xml2);
                */



                sql = "";
                sql += "select b.*from tablesm a ";
                sql += "inner join tablesd b on a.tab_pkid = b.tabd_parent_id ";
                sql += " where a.rec_company_code = '" + comp_code + "' ";
                sql += " and tab_table_name = '" + grp_table_name + "' ";
                sql += " order by tabd_col_order ";
                Dt_Tables = Con_Oracle.ExecuteQuery(sql);

                sqlSelect = "";
                sqlSelect += "  select doc_pkid, doc_slno,doc_name, doc_file_name, doc_table_name, grp_name,grp_level_name ";

                sqlFrom = "";
                sqlFrom += " from pim_docm a inner join " + grp_table_name + " b on a.doc_pkid = b.doc_parent_id ";
                sqlFrom += " left join pim_groupm c on a.doc_grp_id = c.grp_pkid ";

                sqlWhere = "";
                sqlWhere += " where  a.rec_company_code ='" + comp_code + "' ";
                sqlWhere += " and a.doc_table_name = '" + grp_table_name + "'";

                if (searchstring != "")
                {
                    sqlWhere += " and (";
                    sqlWhere += "  a.doc_name like '%" + searchstring.ToLower() + "%'";
                    sqlWhere += " )";
                }

                foreach (DataRow Dr in Dt_Tables.Rows)
                {

                    if (Dr["tabd_col_type"].ToString() == "TEXT" || Dr["tabd_col_type"].ToString() == "NUMBER" || Dr["tabd_col_type"].ToString() == "FILE")
                        sqlSelect += ",b.COL_" + Dr["tabd_col_name"].ToString();
                    if (Dr["tabd_col_type"].ToString() == "DATE")
                        sqlSelect += ",to_char(b.COL_" + Dr["tabd_col_name"].ToString() + ",'" + frontEndDataFormat + "') as COL_" + Dr["tabd_col_name"].ToString();
                    if (Dr["tabd_col_type"].ToString() == "LIST")
                    {
                        iCtr++;
                        sqlSelect += "," + Dr["tabd_col_list"].ToString() + iCtr.ToString() + ".param_name as COL_" + Dr["tabd_col_name"].ToString();
                        sqlFrom += " left join param " + Dr["tabd_col_list"].ToString() + iCtr.ToString() + " on b.COL_" + Dr["tabd_col_name"].ToString() + " = " + Dr["tabd_col_list"].ToString() + iCtr.ToString() + ".param_pkid";
                    }
                }

                sqlSelect += ", a.rec_created_by, a.rec_created_date ";

                sql = "";
                sql = sqlSelect + sqlFrom + sqlWhere;
                sql += " order by c.grp_level_name, a.doc_slno ";

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                fileName_csv = Folder + "\\" + grp_table_name + ".csv";

                string data = "";
                string heading = "ID,GROUP,SKU,FOLDER,FILE";
                foreach (DataRow Dr in Dt_Tables.Rows)
                    heading += "," + Dr["tabd_col_name"].ToString();
                heading += ",CREATED-BY, CREATED-DATE";

                using (StreamWriter sw = File.CreateText(fileName_csv))
                {
                    sw.WriteLine(heading);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        data = "";
                        data += Dr["doc_slno"].ToString();
                        data += "," + Dr["grp_level_name"].ToString();
                        data += "," + Dr["doc_name"].ToString();


                        if (Dr["doc_file_name"].ToString().Length > 0)
                        {
                            pathName = Lib.getPath(ServerImagePath, comp_code, Dr["doc_Table_name"].ToString(), Dr["doc_slno"].ToString(), false);

                            data += "," + pathName.ToString().ToLower();
                            data += "," + Dr["doc_file_name"].ToString();
                        }
                        else
                            data += ",,";

                        foreach (DataRow Dr2 in Dt_Tables.Rows)
                        {
                            if (Dr2["tabd_col_type"].ToString() == "FILE")
                            {
                                sname = Dr["COL_" + Dr2["tabd_col_name"].ToString()].ToString();
                                if (sname != "")
                                    data += "," + sname;
                                else
                                    data += ",";
                            }
                            else if (Dr2["tabd_col_type"].ToString() == "TEXT")
                                data += "," + Dr["COL_" + Dr2["tabd_col_name"].ToString()].ToString();
                            else
                                data += "," + Dr["COL_" + Dr2["tabd_col_name"].ToString()].ToString();
                        }

                        data += "," + Dr["rec_created_by"].ToString();
                        data += "," + Dr["rec_created_date"].ToString();

                        sw.WriteLine(data);
                    }

                }
                Dt_List.Rows.Clear();

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("status", "OK");
            RetData.Add("filename_csv", fileName_csv);
            RetData.Add("filetype_csv", "csv");
            RetData.Add("filedisplayname_csv", grp_table_name + ".csv");

            //RetData.Add("filename_xml1", fileName_xml1);
            //RetData.Add("filetype_xml1", "xml");
            //RetData.Add("filedisplayname_xml1", grp_table_name + "1.xml");

            //RetData.Add("filename_xml2", fileName_xml2);
            //RetData.Add("filetype_xml2", "xml");
            //RetData.Add("filedisplayname_xml2", grp_table_name + "2.xml");

            return RetData;
        }


    }
}
