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
    public class UserRightService : BL_Base
    {

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            List<User> mList = new List<User>();
            User mRow;

            string comp_id = SearchData["comp_id"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();
            string rights_type = SearchData["rights_type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];

            long startrow = 0;
            long endrow = 0;

            try
            {

                sWhere = " where  a.rec_company_code ='" + comp_code + "'";
                if ( rights_type == "USER" )
                    sWhere   += " and user_code <> 'ADMIN'  ";
                if (rights_type == "ROLES")
                    sWhere += " and param_type = 'ROLES'  ";

                if (searchstring != "")
                {
                    if (rights_type == "USER")
                    {
                        sWhere += " and ( ";
                        sWhere += " user_name like '%" + searchstring + "%'";
                        sWhere += " or comp_name like '%" + searchstring + "%'";
                        sWhere += " ) ";
                    }

                    if (rights_type == "ROLES")
                    {
                        
                        sWhere += " and ( ";
                        sWhere += " param_name like '%" + searchstring + "%'";
                        sWhere += " ) ";
                    }

                }
                if (type == "NEW")
                {

                    if (rights_type == "USER")
                    {
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM userm  a ";
                        sql += " left join companym c on a.user_company_id = c.comp_pkid ";
                    }
                    if (rights_type == "ROLES")
                    {
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM param  a ";
                    }
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
                if (rights_type == "USER")
                {
                    sql += " select * from ( ";
                    sql += "  select  a.user_code,a.user_name,a.user_pkid, a.user_email, c.comp_pkid, comp_name, 0 as user_rights_total,";
                    sql += " row_number() over(order by a.user_name, c.comp_name) rn ";
                    sql += "from userm a ";
                    sql += "left  join companym c on a.user_company_id = c.comp_pkid ";
                    sql += sWhere;
                    sql += ") a where rn between {startrow} and {endrow}";
                    sql += " order by a.user_name, a.comp_name ";
                }
                if (rights_type == "ROLES")
                {
                    sql += " select * from ( ";
                    sql += "  select  a.param_code as user_code,a.param_name as user_name,a.param_pkid as user_pkid, null as user_email,";
                    sql += "'{COMPID}' as comp_pkid, null as comp_name, 0 as user_rights_total,";
                    sql += " row_number() over(order by a.param_name) rn ";
                    sql += "from param a ";
                    sql += sWhere;
                    sql += ") a where rn between {startrow} and {endrow}";
                    sql += " order by a.user_name ";
                }

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());
                sql = sql.Replace("{COMPID}", comp_id);


                Dt_List = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new User();
                    mRow.user_pkid = Dr["user_pkid"].ToString();
                    mRow.user_code = Dr["user_code"].ToString();
                    mRow.user_name = Dr["user_name"].ToString();
                    mRow.user_email = Dr["user_email"].ToString();
                    mRow.user_company_id = Dr["comp_pkid"].ToString();
                    mRow.user_branch_name = Dr["comp_name"].ToString();
                    if (Lib.Conv2Integer(Dr["user_rights_total"].ToString()) > 0)
                        mRow.user_rights_total = Lib.Conv2Integer(Dr["user_rights_total"].ToString());
                    mList.Add(mRow);
                }

                Con_Oracle.CloseConnection();
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



        public IDictionary<string, object> RightsList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            List<Modulem> modList = new List<Modulem>();
            Modulem modRow;

            List<UserRights> mList = new List<UserRights>();
            UserRights mRow;


            string type = SearchData["type"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string compid = SearchData["compid"].ToString();
            //string branchid = SearchData["branchid"].ToString();
            string userid = SearchData["userid"].ToString();

            try
            {
                DataTable Dt_Modules = new DataTable();
                sql = "select module_name  from modulem where rec_company_code= '" + comp_code  +"' order by module_order ";
                Dt_Modules = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Modules.Rows)
                {
                    modRow = new Modulem();
                    modRow.module_name = Dr["module_name"].ToString();
                    modList.Add(modRow);
                }

                DataTable Dt_List = new DataTable();
                sql = "";

                sql += " select rights_pkid, module_name, menu_pkid, menu_name, menu_type, rights_company,  ";
                sql += " rights_admin, rights_add, rights_edit,rights_delete, rights_print,  rights_email,rights_docs, ";
                sql += " rights_docs_upload,rights_view, rights_restricted,rights_approval ";
                sql += " from menum a ";
                sql += " inner join modulem b on menu_module_id = module_pkid ";
                sql += " left join userrights c on menu_pkid = rights_menu_id and ";
                sql += " rights_company_id = '{COMPID}' and rights_user_id = '{USERID}' ";
                sql += " where a.rec_company_code = '" + comp_code + "'";
                sql += " order by module_order, menu_order ";


                sql = sql.Replace("{COMPID}", compid);
                sql = sql.Replace("{USERID}", userid);

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new UserRights();
                    mRow.rights_user_id = userid;

                    mRow.rights_company_id = compid;
                    
                    //mRow.rights_branch_id = branchid;

                    if (Dr["rights_pkid"].Equals(DBNull.Value))
                        mRow.rights_id = System.Guid.NewGuid().ToString().ToUpper();
                    else if (Dr["rights_pkid"].ToString().Length <= 0)
                        mRow.rights_id = System.Guid.NewGuid().ToString().ToUpper();
                    else
                        mRow.rights_id = Dr["rights_pkid"].ToString();

                    mRow.module_name = Dr["module_name"].ToString();
                    mRow.menu_id = Dr["menu_pkid"].ToString();
                    mRow.menu_name = Dr["menu_name"].ToString();
                    mRow.menu_type = Dr["menu_type"].ToString();

                    mRow.rights_company = (Dr["rights_company"].ToString() == "Y") ? true : false;
                    mRow.rights_admin = (Dr["rights_admin"].ToString() == "Y") ? true : false;
                    mRow.rights_add = (Dr["rights_add"].ToString() == "Y") ? true : false;
                    mRow.rights_edit = (Dr["rights_edit"].ToString() == "Y") ? true : false;
                    mRow.rights_delete = (Dr["rights_delete"].ToString() == "Y") ? true : false;
                    mRow.rights_print = (Dr["rights_print"].ToString() == "Y") ? true : false;
                    mRow.rights_email = (Dr["rights_email"].ToString() == "Y") ? true : false;
                    mRow.rights_docs = (Dr["rights_docs"].ToString() == "Y") ? true : false;
                    mRow.rights_docs_upload = (Dr["rights_docs_upload"].ToString() == "Y") ? true : false;
                    mRow.rights_view = (Dr["rights_view"].ToString() == "Y") ? true : false;
                    mRow.rights_restricted = (Dr["rights_restricted"].ToString() == "Y") ? true : false;

                    mRow.rights_approval = Dr["rights_approval"].ToString();
                    mList.Add(mRow);
                }
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("list", mList);
            RetData.Add("modules", modList);

            return RetData;

        }





        public IDictionary<string, object> Save(UserRights_VM VM)
        {
            string USERID = "";
            string COMPID = "";
            string BRANCHID = "";
            string MENUID = "";

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            DataTable Dt_Rec = new DataTable();
            DBRecord mRec;

            Con_Oracle = new DBConnection();

            GlobalVariables mGbl = VM.globalVariables;

            try
            {
                Con_Oracle.BeginTransaction();
                foreach (var Rec in VM.userRights)
                {

                    USERID = Rec.rights_user_id;
                    COMPID = Rec.rights_company_id;
                    BRANCHID = Rec.rights_branch_id;
                    MENUID = Rec.menu_id;

                    sql = "";
                    sql += " delete from userrights where ";
                    sql += " rights_user_id = '{USERID}'  and ";
                    sql += " rights_company_id = '{COMPID}' and ";
                    sql += " rights_menu_id = '{MENUID}' ";

                    sql = sql.Replace("{USERID}", USERID);
                    sql = sql.Replace("{COMPID}", COMPID);
                    sql = sql.Replace("{MENUID}", MENUID);
                    Con_Oracle.ExecuteNonQuery(sql);


                    if (Rec.rights_company || Rec.rights_admin || Rec.rights_add || Rec.rights_edit ||
                        Rec.rights_delete || Rec.rights_print || Rec.rights_email || Rec.rights_docs ||
                        Rec.rights_view || Rec.rights_approval.Length > 0)
                    {
                        if (Rec.rights_add || Rec.rights_edit)
                            Rec.rights_view = true;

                        mRec = new DBRecord();
                        mRec.CreateRow("userrights", "ADD", "rights_pkid", Rec.rights_id.ToString());
                        mRec.InsertString("rights_company_id", Rec.rights_company_id.ToString());
                        //mRec.InsertString("rights_branch_id", Rec.rights_branch_id.ToString());
                        mRec.InsertString("rights_user_id", Rec.rights_user_id.ToString());
                        mRec.InsertString("rights_menu_id", Rec.menu_id.ToString());
                        mRec.InsertString("rights_company", (Rec.rights_company) ? "Y" : "N");
                        mRec.InsertString("rights_admin", (Rec.rights_admin) ? "Y" : "N");
                        mRec.InsertString("rights_add", (Rec.rights_add) ? "Y" : "N");
                        mRec.InsertString("rights_edit", (Rec.rights_edit) ? "Y" : "N");
                        mRec.InsertString("rights_delete", (Rec.rights_delete) ? "Y" : "N");
                        mRec.InsertString("rights_print", (Rec.rights_print) ? "Y" : "N");
                        mRec.InsertString("rights_email", (Rec.rights_email) ? "Y" : "N");
                        mRec.InsertString("rights_docs", (Rec.rights_docs) ? "Y" : "N");
                        mRec.InsertString("rights_docs_upload", (Rec.rights_docs_upload) ? "Y" : "N");
                        mRec.InsertString("rights_view", (Rec.rights_view) ? "Y" : "N");
                        mRec.InsertString("rights_restricted", (Rec.rights_restricted) ? "Y" : "N");
                        mRec.InsertString("rights_approval", Rec.rights_approval);

                        Con_Oracle.ExecuteNonQuery(mRec.UpdateRow());
                    }
                }


                /*
                sql = "";
                sql = " update userd set user_rights_total = (select count(*) from userrights where rights_user_id = '{USERID}' and rights_branch_id = '{BRANCHID}')";
                sql += " where user_id = '{USERID}' and user_branch_id = '{BRANCHID}'";
                sql = sql.Replace("{USERID}", USERID);
                sql = sql.Replace("{BRANCHID}", BRANCHID);
                Con_Oracle.ExecuteNonQuery(sql);
                */

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


        public IDictionary<string, object> CopyRights(UserRights_VM VM)
        {
            string USERID = "";
            string BRANCHID = "";
            string MENUID = "";
            string ErrorMessage = "";
            string sMode = "ADD";

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            DataTable Dt_Menu = new DataTable();

            DataTable Dt_Rec = new DataTable();

            DBRecord mRec;

            Con_Oracle = new DBConnection();

            GlobalVariables mGbl = VM.globalVariables;
            string COPYTO_USERID = VM.copyto_user_id;
            string COPYTO_BRANCHID = VM.copyto_branch_id;
            try
            {
                if (COPYTO_USERID.Length <= 0)
                    ErrorMessage += "| Copy User ID Not Found";

                DataTable Dt_Usrbranch = new DataTable();
                sql = "select user_branch_id from userd where user_id = '" + COPYTO_USERID + "'";
                if (COPYTO_BRANCHID.Trim() != "")
                    sql += " and user_branch_id = '" + COPYTO_BRANCHID + "'";

                Dt_Usrbranch = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Usrbranch.Rows.Count <= 0)
                    ErrorMessage += " | Branch Rights Not Found";

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                foreach (DataRow dr in Dt_Usrbranch.Rows)
                {
                    COPYTO_BRANCHID = dr["user_branch_id"].ToString();

                    Con_Oracle.BeginTransaction();
                    foreach (var Rec in VM.userRights)
                    {
                        Rec.rights_id = Guid.NewGuid().ToString().ToUpper();
                        sMode = "ADD";

                        USERID = COPYTO_USERID;
                        BRANCHID = COPYTO_BRANCHID;
                        MENUID = Rec.menu_id;

                        sql = "";
                        sql += " select rights_pkid from userrights where ";
                        sql += " rights_user_id = '{USERID}'  and ";
                        sql += " rights_branch_id = '{BRANCHID}' and ";
                        sql += " rights_menu_id = '{MENUID}' ";

                        sql = sql.Replace("{USERID}", USERID);
                        sql = sql.Replace("{BRANCHID}", BRANCHID);
                        sql = sql.Replace("{MENUID}", MENUID);

                        Dt_Menu = new DataTable();
                        Dt_Menu = Con_Oracle.ExecuteQuery(sql);
                        if (Dt_Menu.Rows.Count > 0)
                        {
                            Rec.rights_id = Dt_Menu.Rows[0].ToString();
                            sMode = "EDIT";
                        }

                        if (Rec.rights_company || Rec.rights_admin || Rec.rights_add || Rec.rights_edit ||
                            Rec.rights_delete || Rec.rights_print || Rec.rights_email || Rec.rights_docs ||
                            Rec.rights_view)
                        {
                            if (Rec.rights_add || Rec.rights_edit)
                                Rec.rights_view = true;

                            Rec.rights_branch_id = BRANCHID;
                            Rec.rights_user_id = USERID;

                            mRec = new DBRecord();
                            mRec.CreateRow("userrights", sMode, "rights_pkid", Rec.rights_id.ToString());
                            mRec.InsertString("rights_branch_id", Rec.rights_branch_id.ToString());
                            mRec.InsertString("rights_user_id", Rec.rights_user_id.ToString());
                            mRec.InsertString("rights_menu_id", Rec.menu_id.ToString());
                            mRec.InsertString("rights_company", (Rec.rights_company) ? "Y" : "N");
                            mRec.InsertString("rights_admin", (Rec.rights_admin) ? "Y" : "N");
                            mRec.InsertString("rights_add", (Rec.rights_add) ? "Y" : "N");
                            mRec.InsertString("rights_edit", (Rec.rights_edit) ? "Y" : "N");
                            mRec.InsertString("rights_delete", (Rec.rights_delete) ? "Y" : "N");
                            mRec.InsertString("rights_print", (Rec.rights_print) ? "Y" : "N");
                            mRec.InsertString("rights_email", (Rec.rights_email) ? "Y" : "N");
                            mRec.InsertString("rights_docs", (Rec.rights_docs) ? "Y" : "N");
                            mRec.InsertString("rights_docs_upload", (Rec.rights_docs_upload) ? "Y" : "N");
                            mRec.InsertString("rights_view", (Rec.rights_view) ? "Y" : "N");
                            // mRec.InsertString("rights_approval", Rec.rights_approval);

                            Con_Oracle.ExecuteNonQuery(mRec.UpdateRow());
                        }
                    }

                    sql = "";
                    sql = " update userd set user_rights_total = (select count(*) from userrights where rights_user_id = '{USERID}' and rights_branch_id = '{BRANCHID}')";
                    sql += " where user_id = '{USERID}' and user_branch_id = '{BRANCHID}'";
                    sql = sql.Replace("{USERID}", USERID);
                    sql = sql.Replace("{BRANCHID}", BRANCHID);
                    Con_Oracle.ExecuteNonQuery(sql);

                    Con_Oracle.CommitTransaction();
                }
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
