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
    public class UserService : BL_Base
    {

        public IDictionary<string, object> LoadCompany(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Companym> mList = new List<Companym>();
            Companym mRow;
            object sversion;
            try
            {

                sql = "select version from versionm";
                DataTable Dt_version = new DataTable();
                sversion = Con_Oracle.ExecuteScalar(sql);

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select comp_pkid,comp_code, comp_name ";
                sql += " from companym b ";
                sql += " where comp_type ='C' ";
                sql += " order by comp_order,comp_name ";



                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Companym();
                    mRow.comp_pkid = Dr["comp_pkid"].ToString();
                    mRow.comp_code = Dr["comp_code"].ToString();
                    mRow.comp_name = Dr["comp_name"].ToString();
                    mList.Add(mRow);
                }

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                Con_Oracle.CreateErrorLog(Ex.Message.ToString());
                throw Ex;
            }
            RetData.Add("list", mList);
            RetData.Add("version", sversion.ToString());
            return RetData;
        }

        public User ValidateUser(string username, string password, string company_code, string ipaddress)
        {
            string brid = "";
            string tokenid = System.Guid.NewGuid().ToString().ToUpper();
            User user = null;
            try
            {
                string spwd = Lib.Encrypt(password.ToUpper());
                string sql = "select user_pkid,user_code, user_name,user_email,user_company_id, user_branch_id,user_local_server, user_branch_user, ";
                sql += " b.param_pkid as sman_id, b.param_name as sman_name,";
                sql += " user_vendor_id, user_region_id, user_role_id,user_role_rights_id, role.param_name as user_role_name ";
                sql += " from userm a ";
                sql += " left join param b on user_sman_id = b.param_pkid ";
                sql += " left join param role on user_role_id = role.param_pkid ";

                sql += " where a.rec_company_code ='" + company_code + "'";
                sql += " and user_islocked='N'  ";
                sql += " and user_code='{user_code}' ";
                sql += " and user_password='{user_password}' ";
                sql = sql.Replace("{user_code}", username.ToUpper());
                sql = sql.Replace("{user_password}", spwd);

                Con_Oracle = new DBConnection();
                DataTable Dt_Record = new DataTable();
                Dt_Record = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Record.Rows)
                {
                    user = new User();
                    user.user_pkid = Dr["user_pkid"].ToString();
                    user.user_code = Dr["user_code"].ToString();
                    user.user_name = Dr["user_name"].ToString();
                    user.user_email = Dr["user_email"].ToString();
                    user.user_company_id = Dr["user_company_id"].ToString();
                    user.user_company_code = company_code;
                    user.user_branch_id = Dr["user_branch_id"].ToString();
                    brid = Dr["user_branch_id"].ToString();
                    user.user_sman_id = Dr["sman_id"].ToString();
                    user.user_sman_name = Dr["sman_name"].ToString();
                    user.user_local_server = Dr["user_local_server"].ToString();

                    user.user_vendor_id = Dr["user_vendor_id"].ToString();
                    user.user_region_id = Dr["user_region_id"].ToString();
                    user.user_role_id = Dr["user_role_id"].ToString();
                    user.user_role_name = Dr["user_role_name"].ToString().ToUpper();
                    user.user_role_rights_id = Dr["user_role_rights_id"].ToString();

                    user.user_ipaddress = ipaddress;
                    user.user_token_id = tokenid;

                    user.user_branch_user = false;
                    if (Dr["user_branch_user"].ToString() == "Y")
                        user.user_branch_user = true;
                }

                if (brid == "" && user == null)
                {
                    sql = " select user_branch_id from userm a  ";
                    sql += " where a.rec_company_code ='" + company_code + "'";
                    sql += " and user_islocked='N'  ";
                    sql += " and user_code='" + username.ToUpper() + "'  and user_branch_id is not null ";
                    DataTable Dt_br1 = new DataTable();
                    Dt_br1 = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_br1.Rows)
                    {
                        brid = Dr["user_branch_id"].ToString();
                    }
                }

                if (brid != "")
                {
                    sql = " select comp_code from companym where comp_pkid = '" + brid + "'";
                    DataTable Dt_br = new DataTable();
                    Dt_br = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_br.Rows)
                    {
                        brid = Dr["comp_code"].ToString();
                    }
                }

                Con_Oracle.CloseConnection();

                string sAction = "SUCCESS";
                if (user == null)
                    sAction = "FAILED";

                string sRemark = ipaddress;
                string IpHostName = GetIPAddressName(company_code, ipaddress);
                decimal IpHostExist = 0;
                if (IpHostName.Trim() != "")
                {
                    IpHostExist = 1;
                    sRemark += "-" + IpHostName;
                }   
                WriteLog("USER-LOGIN", sAction, company_code, brid, username, tokenid, password, sRemark, IpHostExist);
            }
            catch (Exception Ex)
            {
                Con_Oracle.CloseConnection();
                throw Ex;
            }
            return user;
        }

        private string GetIPAddressName(string Comp_code, string iPAddress)
        {
            string iName = "";
          //  sql = "select param_name from param where rec_company_code ='" + Comp_code + "' and param_type='IP ADDRESS' and param_code = '" + iPAddress + "'";
            sql = "select param_name from param where param_type='IP ADDRESS' and param_code = '" + iPAddress + "'";
            try
            {
                Con_Oracle = new DBConnection();
                DataTable Dt_ip = new DataTable();
                Dt_ip = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_ip.Rows)
                {
                    iName = Dr["param_name"].ToString();
                }
            }
            catch (Exception)
            {
            }
            return iName;
        }

        private void WriteLog(string stype, string saction, string comp_code, string branch_code, string user_code, string sKey, string user_pwd, string remarks, decimal IsIPNameBlank = 0)
        {
            try
            {
                Lib.AuditLog("LOGIN", stype, saction, comp_code, branch_code, user_code, sKey, user_pwd, remarks,IsIPNameBlank);
            }
            catch (Exception)
            {
            }
        }
        
        public IDictionary<string, object> LoadMenu(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> Modules = new Dictionary<string, object>();
            Dictionary<string, object> data = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Menum> mList = new List<Menum>();
            Menum mRow;

            List<Modulem> mListModule = new List<Modulem>();
            Modulem moduleRow;

            string usercode = SearchData["usercode"].ToString();
            string userid = SearchData["userid"].ToString();

            string compid = SearchData["compid"].ToString();
            string compcode = SearchData["compcode"].ToString();

            string yearid = SearchData["yearid"].ToString();
            string ipaddress = SearchData["ipaddress"].ToString();
            string tokenid = SearchData["tokenid"].ToString();
            string branchcode = "";

            try
            {

                string serverImageURL  = Lib.GetSeverImageURLWithCompany(compcode);
                string serverImagePath = Lib.GetImagePathWithCompany(compcode);
                string serverReportPath = Lib.GetReportPath(compcode);


                DataTable Dt_Comp = new DataTable();

                sql = "";
                sql += " select comp_pkid, comp_code, comp_name, '' as branch_pkid, '' as branch_code, '' as branch_name, ''  as branch_type ";
                sql += " from companym ";
                sql += " where comp_pkid = '{COMPID}' ";

                sql = sql.Replace("{COMPID}", compid);
                Dt_Comp = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Comp.Rows)
                {
                    branchcode = Dr["branch_code"].ToString();
                    data.Add("comp_pkid", Dr["comp_pkid"].ToString());
                    data.Add("comp_code", Dr["comp_code"].ToString());
                    data.Add("comp_name", Dr["comp_name"].ToString());

                    data.Add("branch_pkid", Dr["branch_pkid"].ToString());
                    data.Add("branch_code", Dr["branch_code"].ToString());
                    data.Add("branch_name", Dr["branch_name"].ToString());
                    data.Add("branch_type", Dr["branch_type"].ToString());

                    data.Add("report_folder", @"c:\Reports");

                    data.Add("server_image_url",  serverImageURL);
                    data.Add("server_image_path", serverImagePath);
                    data.Add("server_report_path", serverReportPath);

                }



                DataTable Dt_Menu = new DataTable();
                sql = "";

                if (usercode.StartsWith("ADMIN"))
                {

                    sql += " select module_name, module_order,menu_pkid, menu_code, menu_name, menu_route1, menu_route2, menu_type, menu_displayed, ";
                    sql += " 'Y' as rights_company,'Y' as rights_admin, 'N' as rights_restricted, 'Y' as rights_add, 'Y' as rights_edit,'Y' as rights_delete, 'Y' as rights_print,  'Y' as rights_email, 'Y' as rights_docs, 'Y' as rights_docs_upload, ";
                    sql += " 'Y' as rights_view,' ' as rights_approval ";
                    sql += " from menum a inner ";
                    sql += " join modulem b on a.menu_module_id = b.module_pkid ";
                    sql += " where a.rec_company_code = '" + compcode + "'";
                    sql += "order by module_order, menu_order ";
                }
                else
                {
                    sql += " select module_name, module_order,menu_pkid, menu_code, menu_name, menu_route1, menu_route2, menu_type, menu_displayed, ";
                    sql += " rights_company,rights_admin,rights_restricted,rights_add,rights_edit,rights_delete, rights_print,  rights_email,rights_docs,rights_docs_upload, ";
                    sql += " rights_view,rights_approval ";
                    sql += " from menum a inner ";
                    sql += " join modulem b on a.menu_module_id = b.module_pkid ";
                    sql += " inner join userrights c on a.menu_pkid = c.rights_menu_id ";
                    sql += " where rights_user_id = (select USER_ROLE_RIGHTS_ID from userm  where USER_PKID = '{USERID}') and rights_company_id = '{COMPID}' ";
                    sql += "order by module_order, menu_order ";
                }

                sql = sql.Replace("{USERID}", userid);
                sql = sql.Replace("{COMPID}", compid);

                Dt_Menu = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Menu.Rows)
                {
                    mRow = new Menum();
                    mRow.menu_pkid = Dr["menu_pkid"].ToString();
                    mRow.menu_code = Dr["menu_code"].ToString();
                    mRow.menu_name = Dr["menu_name"].ToString();
                    mRow.menu_route1 = Dr["menu_route1"].ToString();
                    mRow.menu_route2 = Dr["menu_route2"].ToString();
                    mRow.menu_type = Dr["menu_type"].ToString();
                    mRow.menu_module_name = Dr["module_name"].ToString();

                    if (Dr["menu_displayed"].ToString() == "Y")
                        mRow.menu_displayed = true;
                    else
                        mRow.menu_displayed = false;


                    mRow.menu_sep = false;
                    if (mRow.menu_type.ToString().Contains("-"))
                        mRow.menu_sep = true;

                    mRow.rights_company = (Dr["rights_company"].ToString() == "Y") ? true : false;
                    mRow.rights_admin = (Dr["rights_admin"].ToString() == "Y") ? true : false;
                    mRow.rights_restricted = (Dr["rights_restricted"].ToString() == "Y") ? true : false;
                    mRow.rights_add = (Dr["rights_add"].ToString() == "Y") ? true : false;
                    mRow.rights_edit = (Dr["rights_edit"].ToString() == "Y") ? true : false;
                    mRow.rights_delete = (Dr["rights_delete"].ToString() == "Y") ? true : false;
                    mRow.rights_print = (Dr["rights_print"].ToString() == "Y") ? true : false;
                    mRow.rights_email = (Dr["rights_email"].ToString() == "Y") ? true : false;
                    mRow.rights_docs = (Dr["rights_docs"].ToString() == "Y") ? true : false;
                    mRow.rights_docs_upload = (Dr["rights_docs_upload"].ToString() == "Y") ? true : false;
                    mRow.rights_view = (Dr["rights_view"].ToString() == "Y") ? true : false;
                    mRow.rights_approval = Dr["rights_approval"].ToString();

                    if (!Modules.ContainsKey(mRow.menu_module_name))
                    {
                        Modules.Add(mRow.menu_module_name, mRow.menu_module_name);
                        moduleRow = new Modulem();
                        moduleRow.module_name = mRow.menu_module_name;
                        mListModule.Add(moduleRow);
                    }
                    mList.Add(mRow);
                }




                Con_Oracle.CloseConnection();

                string Errormsg = "";

                if (Dt_Comp.Rows.Count <= 0)
                    Errormsg += " | Company/Branch Data Not Found";

                if (Dt_Menu.Rows.Count <= 0)
                    Errormsg += " | No Menu Options Provided";

                string sAction = "SUCCESS";
                if (Errormsg != "")
                    sAction = "FAILED";

                string sYrName = "";
                //if (data.ContainsKey("year_name"))
                //    sYrName = data["year_name"].ToString();

                string sRemark = ipaddress;
                string IpHostName = GetIPAddressName(compcode, ipaddress);
                decimal IpHostExist = 0;
                if (IpHostName.Trim() != "")
                {
                    IpHostExist = 1;
                    sRemark += "-" + IpHostName;
                }

                WriteLog("BRANCH-LOGIN", sAction, compcode, branchcode, usercode, tokenid, sYrName, sRemark, IpHostExist);
                if (Errormsg != "")
                    throw new Exception(Errormsg);

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("data", data);
            RetData.Add("list", mList);
            RetData.Add("modules", mListModule);
            return RetData;
        }
        
        public IDictionary<string, object> LoadBranch(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Companym> mList = new List<Companym>();
            Companym mRow;


            string userid = SearchData["userid"].ToString();
            string usercode = SearchData["usercode"].ToString();
            string compid = SearchData["compid"].ToString();
            string compcode = SearchData["compcode"].ToString();

            try
            {
                DataTable Dt_List = new DataTable();
                sql = "";
                if (usercode.StartsWith("ADMIN"))
                {
                    sql += " select comp_pkid, comp_name ";
                    sql += " from companym b ";
                    sql += " where comp_type ='B' and comp_parent_id = '" + compid + "' ";
                    sql += " order by comp_name ";
                }
                else
                {
                    sql += " select comp_pkid, comp_name ";
                    sql += " from userd a inner join companym b on a.user_branch_id = b.comp_pkid ";
                    sql += " where a.user_id = '" + userid + "' ";
                    sql += " order by comp_name ";
                }

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Companym();
                    mRow.comp_pkid = Dr["comp_pkid"].ToString();
                    mRow.comp_name = Dr["comp_name"].ToString();
                    mList.Add(mRow);
                }



            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("branchlist", mList);
            return RetData;
        }
        
        private Param GetCurrency(string Curr_ID)
        {
            Param mRow = new Param();
            mRow.param_rate = 0;
            mRow.param_id1 = ""; 

            DataTable Dt_Param = new DataTable();
            if (Curr_ID != "")
            {
                sql = "select param_rate,param_id1 from param where param_pkid ='" + Curr_ID + "'";
                Dt_Param = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Param.Rows.Count > 0)
                {
                    mRow.param_rate = Lib.Convert2Decimal(Dt_Param.Rows[0]["param_rate"].ToString());
                    mRow.param_id1 = Dt_Param.Rows[0]["param_id1"].ToString();
                }
            }
            Dt_Param.Rows.Clear();
            return mRow;
        }

        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();
            List<Companym> mList = new List<Companym>();
            Companym mRow;

            //string type = SearchData["type"].ToString();

            try
            {
                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select comp_pkid,comp_name from companym where comp_type ='B' order by comp_name";

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Companym();
                    mRow.comp_pkid = Dr["comp_pkid"].ToString();
                    mRow.comp_name = Dr["comp_name"].ToString();
                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("companylist", mList);

            return RetData;
        }

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<User> mList = new List<User>();
            User mRow;

            string comp_code = SearchData["comp_code"].ToString();
            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();

            Boolean user_admin = (Boolean)SearchData["user_admin"];
            string region_id = SearchData["region_id"].ToString();
            string vendor_id = SearchData["vendor_id"].ToString();
            string role_name = SearchData["role_name"].ToString();


            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            Boolean rights_admin = (Boolean)SearchData["rights_admin"];
            string user_pkid = SearchData["user_pkid"].ToString();
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  a.rec_company_code ='" + comp_code + "'  and a.user_code <> 'ADMIN'  ";

                //if ( rights_admin ==false)
                //{
                //    sWhere += " and a.user_pkid = '" + user_pkid + "'";
                //}


                if (role_name == "SUPER ADMIN")
                    sWhere += " and role.param_name in('SUPER ADMIN', 'ZONE ADMIN','SALES EXECUTIVE','VENDOR','RECCE USER') ";

                if (role_name == "ZONE ADMIN")
                {
                    sWhere += " and a.user_region_id = '" + region_id + "'";
                    sWhere += " and role.param_name in('SALES EXECUTIVE','VENDOR','RECCE USER')";
                }
                if (role_name == "SALES EXECUTIVE")
                {
                    sWhere += " and a.user_region_id = '" + region_id + "'";
                    sWhere += " and role.param_name in('VENDOR','RECCE USER')";
                }
                if (role_name == "VENDOR")
                    sWhere += " and role.param_name in('RECCE USER') and a.user_vendor_id ='" + vendor_id + "' ";


                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  a.user_code like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or a.user_name like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or a.user_email like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or b.comp_name like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM userm a ";
                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM userm  a ";



                    sql += " left join companym b on a.user_branch_id = b.comp_pkid ";
                    sql += " left join param role on a.user_role_id = role.param_pkid ";

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
                sql += " select  a.user_pkid, a.user_code, a.user_name , a.user_email, b.comp_name as branch_name,a.user_branch_user,";
                sql += " a.user_local_server, parent.user_name as user_parent_name, ";
                sql += " sman.param_name as user_sman_name, role.param_name as user_role_name,";
                sql += " region.param_name as user_region_name, vendor.comp_name as user_vendor_name, ";
                sql += " row_number() over(order by a.user_name) rn ";
                sql += " from userm a ";
                sql += " left join companym b on a.user_branch_id = b.comp_pkid ";
                sql += " left join param sman on a.user_sman_id = sman.param_pkid ";
                sql += " left join userm parent on a.user_parent_id = parent.user_pkid ";
                sql += " left join param role on a.user_role_id = role.param_pkid ";

                sql += " left join companym vendor on a.user_vendor_id = vendor.comp_pkid ";
                sql += " left join param region  on a.user_region_id = region.param_pkid ";

                sql += " " + sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by a.user_name ";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new User();
                    mRow.user_pkid = Dr["user_pkid"].ToString();
                    mRow.user_code = Dr["user_code"].ToString();

                    mRow.user_name = Dr["user_name"].ToString();

                    mRow.user_parent_name = Dr["user_parent_name"].ToString();

                    mRow.user_email = Dr["user_email"].ToString();
                    mRow.user_sman_name = Dr["user_sman_name"].ToString();
                    mRow.user_branch_name = Dr["branch_name"].ToString();

                    if (Dr["user_role_name"].Equals(DBNull.Value) )
                        mRow.user_role_name = "user";
                    else 
                        mRow.user_role_name = Dr["user_role_name"].ToString();

                    mRow.user_vendor_name = Dr["user_vendor_name"].ToString();
                    mRow.user_region_name = Dr["user_region_name"].ToString();

                    mRow.user_local_server = Dr["user_local_server"].ToString();
                    mRow.user_branch_user = Dr["user_branch_user"].ToString() == "Y" ? true : false;
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

        public Dictionary<string, object> NewUserDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            User mRow = new User();

            string id = SearchData["pkid"].ToString();
            string role_name = SearchData["role_name"].ToString();
            string comp_id = SearchData["comp_id"].ToString();
            string comp_code = SearchData["comp_code"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  ";
                sql += " role.param_pkid as user_role_id, role.param_name as user_role_name,";
                sql += " a.user_region_id, region.param_name as user_region_name,";
                sql += " a.user_vendor_id, vendor.comp_name as user_vendor_name";
                sql += " from userm a  ";
                sql += " left join param region on a.user_region_id = region.param_pkid ";
                sql += " left join companym vendor on a.user_vendor_id = vendor.comp_pkid ";
                sql += " left join param role  on role.param_type = 'ROLES' and role.param_name = 'RECCE USER' and role.rec_company_code = '" + comp_code + "'";
                sql += " where  a.user_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new User();

                    mRow.user_role_id = Dr["user_role_id"].ToString();
                    mRow.user_role_name = Dr["user_role_name"].ToString();

                    mRow.user_region_id = Dr["user_region_id"].ToString();
                    mRow.user_region_name = Dr["user_region_name"].ToString();

                    mRow.user_vendor_id = Dr["user_vendor_id"].ToString();
                    mRow.user_vendor_name = Dr["user_vendor_name"].ToString();

                    break;
                }

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


        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            User mRow = new User();

            List<Userd> mList = new List<Userd>();
            Userd mRowDet;


            string id = SearchData["pkid"].ToString();
            string comp_id = SearchData["comp_id"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  a.user_pkid, a.user_code, a.user_name, a.user_email, a.user_password, a.user_branch_id,a.user_branch_user,a.user_local_server,";
                sql += " a.user_parent_id, parent.user_name as user_parent_name,  ";
                sql += " a.user_sman_id, sman.param_code as user_sman_code,sman.param_name as user_sman_name,a.user_email_pwd, ";
                sql += " a.user_role_id, role.param_name as user_role_name, ";
                sql += " a.user_region_id, region.param_name as user_region_name,";
                sql += " a.user_vendor_id, vendor.comp_name as user_vendor_name";

                sql += " from userm a  ";
                sql += " left join param sman on a.user_sman_id = sman.param_pkid ";
                sql += " left join userm parent on a.user_parent_id = parent.user_pkid ";
                sql += " left join param role on a.user_role_id = role.param_pkid ";

                sql += " left join param region on a.user_region_id = region.param_pkid ";
                sql += " left join companym vendor on a.user_vendor_id = vendor.comp_pkid ";


                sql += " where  a.user_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new User();
                    mRow.user_pkid = Dr["user_pkid"].ToString();
                    mRow.user_code = Dr["user_code"].ToString();
                    mRow.user_name = Dr["user_name"].ToString();
                    mRow.user_parent_id = Dr["user_parent_id"].ToString();
                    mRow.user_parent_name = Dr["user_parent_name"].ToString();
                    mRow.user_email = Dr["user_email"].ToString();
                    mRow.user_password = Lib.Decrypt(Dr["user_password"].ToString());
                    mRow.user_branch_id = Dr["user_branch_id"].ToString();
                    mRow.user_email_pwd = Dr["user_email_pwd"].ToString();


                    mRow.user_role_id = Dr["user_role_id"].ToString();
                    mRow.user_role_name = Dr["user_role_name"].ToString();

                    mRow.user_region_id = Dr["user_region_id"].ToString();
                    mRow.user_region_name = Dr["user_region_name"].ToString();

                    mRow.user_vendor_id = Dr["user_vendor_id"].ToString();
                    mRow.user_vendor_name = Dr["user_vendor_name"].ToString();

                    mRow.user_local_server = Dr["user_local_server"].ToString();

                    mRow.user_branch_user = false;
                    if (Dr["user_branch_user"].ToString() == "Y")
                        mRow.user_branch_user = true;
                    mRow.user_sman_id = Dr["user_sman_id"].ToString();
                    mRow.user_sman_code = Dr["user_sman_code"].ToString();
                    mRow.user_sman_name = Dr["user_sman_name"].ToString();
                    break;
                }


                DataTable Dt_Det = new DataTable();
                sql = "";
                sql += " select user_id, user_branch_id, branch.comp_pkid, branch.comp_name ";
                sql += " from companym branch ";
                sql += " left join userd b on branch.comp_pkid = b.user_branch_id  and b.user_id = '{USERID}' ";
                sql += " where comp_parent_id = '{COMP_ID}' and branch.comp_type = 'B' ";
                sql += " order by branch.comp_name ";

                sql = sql.Replace("{COMP_ID}", comp_id);
                sql = sql.Replace("{USERID}", id);

                Dt_Det = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Det.Rows)
                {
                    mRowDet = new Userd();
                    mRowDet.user_branch_id = Dr["comp_pkid"].ToString();
                    mRowDet.user_branch_name = Dr["comp_name"].ToString();
                    if (Dr["user_id"].Equals(DBNull.Value))
                        mRowDet.user_selected = false;
                    else
                        mRowDet.user_selected = true;

                    mRowDet.user_default_branch_id = false;
                    if (mRowDet.user_branch_id == mRow.user_branch_id)
                        mRowDet.user_default_branch_id = true;

                    mList.Add(mRowDet);
                }
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            RetData.Add("recorddet", mList);
            return RetData;
        }

        public string AllValid(User Record)
        {
            string str = "";
            try
            {
                sql = "select user_pkid from (";
                sql += "select user_pkid  from userm a where a.rec_company_code = '{COMP_CODE}' ";
                sql += " and (a.user_code = '{CODE}' or a.user_name = '{NAME}')  ";
                sql += ") a where user_pkid <> '{PKID}'";

                sql = sql.Replace("{COMP_CODE}", Record._globalvariables.comp_code);
                sql = sql.Replace("{CODE}", Record.user_code);
                sql = sql.Replace("{NAME}", Record.user_name);
                sql = sql.Replace("{PKID}", Record.user_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(User Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            string default_branch_id = "";
            try
            {
                Con_Oracle = new DBConnection();

                if (Record.user_code.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Code Cannot Be Empty");
                if (Record.user_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");
                if (Record.user_password.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Password Cannot Be Empty");
                if (Record.user_role_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Role Cannot Be Empty");

                if (Record.user_role_name.Trim().Length > 0)
                {

                    if (Record.user_role_name.Trim().ToUpper() == "ZONE ADMIN")
                    {
                        if (Record.user_region_id.Trim().Length <= 0)
                            Lib.AddError(ref ErrorMessage, "Region Need To Be Entered");
                    }
                    if (Record.user_role_name.Trim().ToUpper() == "SALES EXECUTIVE")
                    {
                        if (Record.user_region_id.Trim().Length <= 0)
                            Lib.AddError(ref ErrorMessage, "Region Need To Be Entered");
                    }

                    if (Record.user_role_name.Trim().ToUpper() == "VENDOR")
                    {
                        if (Record.user_vendor_id.Trim().Length <= 0)
                            Lib.AddError(ref ErrorMessage, "Vendor Need To Be Entered");
                        if (Record.user_region_id.Trim().Length <= 0)
                            Lib.AddError(ref ErrorMessage, "Region Need To Be Entered");
                    }

                    if (Record.user_role_name.Trim().ToUpper() == "RECCE USER" ) {
                        if (Record.user_vendor_id.Trim().Length <= 0)
                            Lib.AddError(ref ErrorMessage, "Vendor Need To Be Entered");
                        if (Record.user_region_id.Trim().Length <= 0)
                            Lib.AddError(ref ErrorMessage, "Region Need To Be Entered");
                    }

                }


                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("userm", Record.rec_mode, "user_pkid", Record.user_pkid);
                Rec.InsertString("user_code", Record.user_code);

                Rec.InsertString("user_name", Record.user_name);

                Rec.InsertString("user_parent_id", Record.user_parent_id);

                Rec.InsertString("user_email", Record.user_email, "L");

                Rec.InsertString("user_password", Lib.Encrypt(Record.user_password), "P");

                Rec.InsertString("user_branch_user", (Record.user_branch_user) ? "Y" : "N");

                Rec.InsertString("user_email_pwd", Record.user_email_pwd, "P");

                Rec.InsertString("user_sman_id", Record.user_sman_id);

                Rec.InsertString("user_role_id", Record.user_role_id);

                Rec.InsertString("user_vendor_id", Record.user_vendor_id);
                Rec.InsertString("user_region_id", Record.user_region_id);


                if ( Record.user_role_id.ToString().Length >0  )
                    Rec.InsertString("user_role_rights_id", Record.user_role_id);
                else
                    Rec.InsertString("user_role_rights_id", Record.user_pkid);

                Rec.InsertString("user_local_server", Record.user_local_server, "L");

                Rec.InsertString("user_issupervisor", "N");
                Rec.InsertString("user_islocked", "N");
                Rec.InsertString("rec_deleted", "N");
                Rec.InsertString("rec_updated", "N");
                Rec.InsertString("rec_printed", "N");
                Rec.InsertString("rec_locked", "N");
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("user_company_id", Record._globalvariables.comp_pkid);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                }
                else
                {
                    Rec.InsertString("user_company_id", Record._globalvariables.comp_pkid);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                }
                sql = Rec.UpdateRow();


                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);

                /*
                sql = "delete from userd where user_id = '{USRID}'";
                sql = sql.Replace("{USRID}", Record.user_pkid.ToString());
                Con_Oracle.ExecuteNonQuery(sql);
                foreach (Userd _Rec in Record.recorddet)
                {
                    if (_Rec.user_selected)
                    {
                        sql = " insert into userd (user_id, user_branch_id) values ('{USRID}','{BRID}')";
                        sql = sql.Replace("{USRID}", Record.user_pkid.ToString());
                        sql = sql.Replace("{BRID}", _Rec.user_branch_id.ToString());
                        Con_Oracle.ExecuteNonQuery(sql);

                        if (_Rec.user_default_branch_id)
                            default_branch_id = _Rec.user_branch_id.ToString();
                    }
                }
                sql = " update userm set user_branch_id  = '" + default_branch_id + "'  where user_pkid = '" + Record.user_pkid.ToString() + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                */

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


        public Boolean SaveDocuments(Dictionary<string, object> SearchData)
        {
            string sql2 = "";
            Boolean bRet = false;
            Con_Oracle = new DBConnection();

            try
            {
                if (SearchData["TYPE"].ToString() == "CUSTOMER")
                {
                    sql = "select  param_code from param where param_code= 'KYC' and param_pkid = '" + SearchData["CATG_ID"].ToString() + "'";
                    if (Con_Oracle.IsRowExists(sql))
                        sql2 = "update customerm set cust_kyc_status = 'Y' where cust_pkid ='" + SearchData["PARENT_ID"].ToString() + "'";
                }


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("documentm", "ADD", "doc_pkid", SearchData["PKID"].ToString());
                Rec.InsertString("doc_parent_id", SearchData["PARENT_ID"].ToString());
                Rec.InsertString("doc_type", SearchData["TYPE"].ToString());
                Rec.InsertString("doc_path", SearchData["PATH"].ToString());
                Rec.InsertString("doc_filesize", SearchData["SIZE"].ToString());
                Rec.InsertString("doc_catg_id", SearchData["CATG_ID"].ToString());
                Rec.InsertString("doc_file_name", SearchData["FILENAME"].ToString());
                Rec.InsertString("rec_created_by", SearchData["CREATED_BY"].ToString());
                Rec.InsertString("rec_company_code", SearchData["COMP_CODE"].ToString());
                Rec.InsertString("rec_branch_code", SearchData["BRANCH_CODE"].ToString());
                Rec.InsertString("rec_deleted", "N");

                if (Con_Oracle.DB == "ORACLE")
                    Rec.InsertFunction("rec_created_date", "SYSDATE");
                else
                    Rec.InsertFunction("rec_created_date", "GETDATE()");

                Rec.InsertString("doc_group_id", SearchData["GROUP_ID"].ToString());

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                if (sql2 != "")
                    Con_Oracle.ExecuteNonQuery(sql2);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                UpdateDocCount(SearchData["PARENT_ID"].ToString(), SearchData["TYPE"].ToString(), SearchData["GROUP_ID"].ToString());
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
            return bRet;

        }

        public IDictionary<string, object> DocumentList(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<documentm> mList = new List<documentm>();
            documentm mRow;

            string comp_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string parent_id = SearchData["parent_id"].ToString();
            string group_id = SearchData["group_id"].ToString();

            string root_folder = SearchData["root_folder"].ToString();
            string sub_folder = SearchData["sub_folder"].ToString();

            string subtype = "";
            if (SearchData.ContainsKey("subtype"))
                subtype = SearchData["subtype"].ToString();

            try
            {

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select doc_pkid,doc_catg_id, param_name as doc_catg_name, doc_path, doc_file_name,doc_filesize,";
                sql += " a.rec_created_by, a.rec_created_date,nvl(a.rec_deleted, 'N') as rec_deleted,";
                sql += " a.rec_deleted_by,doc_group_id ";
                sql += " from documentm a left join param b on doc_catg_id = param_pkid ";
                sql += " where ( doc_parent_id = '" + parent_id + "'";

                if (group_id.Trim() != "")
                    sql += " and doc_group_id = '" + group_id + "'";
                else
                    sql += " or doc_group_id = '" + parent_id + "'";

                sql += " )";

                if (subtype != "DELETED")
                    sql += " and nvl(a.rec_deleted, 'N') = 'N'";
                sql += " order by a.rec_created_date, doc_file_name";

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new documentm();
                    mRow.doc_pkid = Dr["doc_pkid"].ToString();
                    mRow.doc_catg_id = Dr["doc_catg_id"].ToString();
                    mRow.doc_catg_name = Dr["doc_catg_name"].ToString();
                    mRow.doc_file_name = Dr["doc_file_name"].ToString();
                    mRow.rec_created_by = Dr["rec_created_by"].ToString();
                    mRow.doc_group_id = Dr["doc_group_id"].ToString();

                    mRow.doc_file_size = Dr["doc_filesize"].ToString();

                    mRow.rec_deleted_by = "N";
                    if (Dr["rec_deleted"].ToString() == "Y")
                        mRow.rec_deleted_by = "Y - " + Dr["rec_deleted_by"].ToString();




                    //mRow.doc_full_name = @"D://documents/ALLDOCS/" + Dr["doc_path"].ToString() + "/" + Dr["doc_file_name"].ToString();

                    mRow.doc_full_name = root_folder + Dr["doc_path"].ToString() + "/" + Dr["doc_file_name"].ToString();

                    mRow.rec_created_date = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    mRow.row_displayed = false;
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

        public IDictionary<string, object> ExtraList(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<documentm> mList = new List<documentm>();
            documentm mRow;

            string comp_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();

            string root_folder = SearchData["root_folder"].ToString();
            string sub_folder = SearchData["sub_folder"].ToString();

            string copy_type = SearchData["copy_type"].ToString();
            string copy_no = SearchData["copy_no"].ToString();

            string year_code = SearchData["year_code"].ToString();

            string ids = "";

            try
            {
                DataTable Dt_Rec = new DataTable();
                sql = "";
                if (copy_type == "MBL-SE")
                {
                    sql = " select hbl_pkid as pkid from hblm where rec_company_code = '" + comp_code + "' and rec_branch_code = '" + branch_code + "' and hbl_Type = 'MBL-SE' ";
                    sql += " and hbl_no = '" + copy_no + "' and hbl_year =  " + year_code ;
                }
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    if (ids != "")
                        ids += ",";
                    ids += "'" + Dr["pkid"].ToString() + "'";
                }

                 DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select doc_pkid,doc_catg_id, param_name as doc_catg_name, doc_path, doc_file_name, doc_filesize,";
                sql += " a.rec_created_by, a.rec_created_date,nvl(a.rec_deleted, 'N') as rec_deleted,";
                sql += " a.rec_deleted_by,doc_group_id ";
                sql += " from documentm a inner join param b on doc_catg_id = param_pkid and param_code in ('INVOICE')";
                sql += " where ";
                sql += " (doc_parent_id in (" + ids + ") or doc_group_id in (" + ids + ") )";
                sql += " and nvl(a.rec_deleted, 'N') = 'N'";
                sql += " order by a.rec_created_date, doc_file_name";

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new documentm();
                    mRow.doc_pkid = Dr["doc_pkid"].ToString();
                    mRow.doc_catg_id = Dr["doc_catg_id"].ToString();
                    mRow.doc_catg_name = Dr["doc_catg_name"].ToString();
                    mRow.doc_file_name = Dr["doc_file_name"].ToString();
                    mRow.rec_created_by = Dr["rec_created_by"].ToString();
                    mRow.doc_group_id = Dr["doc_group_id"].ToString();
                    mRow.doc_file_size = Dr["doc_filesize"].ToString();
                    mRow.rec_deleted_by = "N";
                    if (Dr["rec_deleted"].ToString() == "Y")
                        mRow.rec_deleted_by = "Y - " + Dr["rec_deleted_by"].ToString();


                    //mRow.doc_full_name = @"D://documents/ALLDOCS/" + Dr["doc_path"].ToString() + "/" + Dr["doc_file_name"].ToString();

                    mRow.doc_full_name = root_folder + Dr["doc_path"].ToString() + "/" + Dr["doc_file_name"].ToString();

                    mRow.rec_created_date = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    mRow.row_displayed = false;
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

            return RetData;
        }

        public IDictionary<string, object> CopyFiles(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            
            
            string comp_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();

            string root_folder = SearchData["root_folder"].ToString();
            string sub_folder = SearchData["sub_folder"].ToString();

            string pkids = SearchData["pkids"].ToString();

            string type = SearchData["type"].ToString();


            string year_code = SearchData["year_code"].ToString();

            string parentid = SearchData["parentid"].ToString();

            string created_by = SearchData["created_by"].ToString() ;


            try
            {

                sql = " insert into documentm (";
                sql += " DOC_PKID, DOC_PARENT_ID, REC_CREATED_BY, REC_CREATED_DATE, ";
                sql += " DOC_TYPE, DOC_CATG_ID, DOC_PATH, DOC_FILE_NAME, DOC_REFNO, DOC_DESC, DOC_FILESIZE, DOC_READCOUNT, DOC_USER_ID, DOC_PROTECTED, DOC_DOCREFNO, DOC_CTR, REC_ORIGIN, ";
                sql += " REC_COMPANY_CODE, REC_BRANCH_CODE, REC_DELETED";
                sql += " ) ";
                sql += " select  upper(sys_guid()), '" + parentid + "', '" + created_by + "', sysdate, ";
                sql += " '" + type + "', DOC_CATG_ID, DOC_PATH, DOC_FILE_NAME, DOC_REFNO, DOC_DESC, DOC_FILESIZE, DOC_READCOUNT, DOC_USER_ID, DOC_PROTECTED, DOC_DOCREFNO, DOC_CTR, REC_ORIGIN, ";
                sql += " REC_COMPANY_CODE, REC_BRANCH_CODE, REC_DELETED";
                sql += " from documentm where doc_pkid in (" + pkids   + ")";

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteQuery(sql);
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

            RetData.Add("list", "");

            return RetData;
        }

        public IDictionary<string, object> LoadDocumentCategory(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "param");
            parameter.Add("param_type", "DOCTYPE");
            parameter.Add("comp_code", comp_code);
            RetData.Add("dtlist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }

        public IDictionary<string, object> DeleteDocument(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            string id = SearchData["pkid"].ToString();
            string user_code = SearchData["user_code"].ToString();

            string type = "";
            if (SearchData.ContainsKey("type"))
                type = SearchData["type"].ToString();
            string parentid = "";
            if (SearchData.ContainsKey("parentid"))
                parentid = SearchData["parentid"].ToString();

            string groupid = "";
            sql = "select doc_parent_id,doc_type, doc_group_id from documentm where doc_pkid  = '" + id + "'";
            DataTable Dt_Doc = new DataTable();
            Dt_Doc = Con_Oracle.ExecuteQuery(sql);
            if (Dt_Doc.Rows.Count > 0)
            {
                parentid = Dt_Doc.Rows[0]["doc_parent_id"].ToString();
                type = Dt_Doc.Rows[0]["doc_type"].ToString();
                groupid = Dt_Doc.Rows[0]["doc_group_id"].ToString();
            }

            sql = "update  documentm set rec_deleted = 'Y', rec_deleted_by ='" + user_code + "', ";
            sql += " rec_deleted_date = sysdate  where doc_pkid  = '" + id + "'";
            try
            {
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
                Dt_Doc.Rows.Clear();
                if (parentid != "")
                    UpdateDocCount(parentid, type, groupid);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            return RetData;
        }

        private void UpdateDocCount(string parent_id, string parent_type, string group_id = "")
        {
            Boolean bTrans = false;
            int Doc_Count = 0;
            string sql2 = "";
            try
            {
                Con_Oracle = new DBConnection();

                sql = " select count(doc_parent_id) as Docs from documentm where ";
                if (parent_type == "IMPORT" || parent_type == "MBL-SE" || parent_type == "MBL-AE")
                    sql += " ( doc_parent_id = '" + parent_id + "' or  doc_group_id = '" + parent_id + "' )";
                else
                    sql += "  doc_parent_id = '" + parent_id + "'";

                sql += "  and nvl(rec_deleted, 'N') = 'N'";
                DataTable Dt_Docs = new DataTable();
                Dt_Docs = Con_Oracle.ExecuteQuery(sql);

                if (Dt_Docs.Rows.Count > 0)
                    Doc_Count = Lib.Conv2Integer(Dt_Docs.Rows[0]["Docs"].ToString());


                sql = "";
                if (parent_type == "JOB")
                    sql = "update jobm set job_docs = {COUNT} where job_pkid ='{PKID}' ";
                else if (parent_type == "IMPORT" || parent_type == "MBL-SE" || parent_type == "MBL-AE")
                    sql = "update hblm set hbl_docs = {COUNT} where hbl_pkid ='{PKID}' ";
                else if (parent_type == "CUSTOMER")
                    sql = "update customerm set cust_docs = {COUNT} where cust_pkid ='{PKID}' ";
                else if (parent_type == "ACC-LEDGER")
                    sql = "update ledgerh set jvh_docs = {COUNT} where jvh_pkid ='{PKID}' ";
                else if (parent_type == "TDS-CERTIFICATE")
                    sql = "update tdscertm set tds_doc_count = {COUNT} where tds_pkid ='{PKID}' ";

                sql = sql.Replace("{COUNT}", Doc_Count.ToString());
                sql = sql.Replace("{PKID}", parent_id);


                if (group_id != "" && parent_type == "ACC-LEDGER")
                {
                    sql2 = " select count(doc_parent_id) as Docs from documentm ";
                    sql2 += "  where (doc_parent_id = '" + group_id + "' or doc_group_id = '" + group_id + "' )";
                    sql2 += "  and nvl(rec_deleted, 'N') = 'N'";
                    Dt_Docs = new DataTable();
                    Dt_Docs = Con_Oracle.ExecuteQuery(sql2);

                    if (Dt_Docs.Rows.Count > 0)
                        Doc_Count = Lib.Conv2Integer(Dt_Docs.Rows[0]["Docs"].ToString());

                    sql2 = "update hblm set hbl_docs = {COUNT} where hbl_pkid ='{PKID}' ";

                    sql2 = sql2.Replace("{COUNT}", Doc_Count.ToString());
                    sql2 = sql2.Replace("{PKID}", group_id);
                }


                if (sql != "")
                {
                    Con_Oracle.BeginTransaction();
                    bTrans = true;
                    Con_Oracle.ExecuteNonQuery(sql);
                    if (sql2 != "")
                        Con_Oracle.ExecuteNonQuery(sql2);
                    Con_Oracle.CommitTransaction();
                }
                Con_Oracle.CloseConnection();
                Dt_Docs.Rows.Clear();
            }
            catch (Exception)
            {
                if (Con_Oracle != null)
                {
                    if (bTrans)
                        Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
            }
        }

        private void UpdateLogin(string id)
        {
            try
            {
                Con_Oracle = new DBConnection();
                Con_Oracle.BeginTransaction();
                sql = "update auditlog set audit_Action = 'SUCCESS' where audit_pkey = '" + id  + "' and audit_type ='USER-LOGIN'";
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception)
            {
                Con_Oracle.RollbackTransaction();
                Con_Oracle.CloseConnection();
            }

        }

        private void UpdateLocking(Dictionary<string, object> GlobalData)
        {
            Boolean bTrans = false;
            Boolean IsPrevYear = false;
            int mm = 0;
            int dd = 0;
            int yy = 0;

            DateTime mDate2;
            DateTime mDate3;
            DateTime mDate5;
            DateTime mDate10;

            string mDate_IN = "";
            string mDate_PN = "";
            string mDate_BR = "";
            string mDate_BP = "";
            string mDate_CR = "";
            string mDate_CP = "";
            string mDate_DN = "";
            string mDate_CN = "";
            string mDate_DI = "";
            string mDate_CI = "";
            string mDate_JV = "";
            string mDate_HO = "";

            try
            {
                Con_Oracle = new DBConnection();

                string company_code = GlobalData["comp_code"].ToString();
                string branch_code = GlobalData["branch_code"].ToString();
                string year_code = GlobalData["year_code"].ToString();
                string year_start_date = GlobalData["year_start_date"].ToString(); //"yyyy-MM-dd"

                //string SERVERDATE = DateTime.Now.ToString(Lib.BACK_END_DATE_FORMAT);
                DateTime DATABASEDATE = DateTime.Now;
                int SERV_DATE_DD = DateTime.Now.Day;
                int SERV_DATE_MM = DateTime.Now.Month;
                int SERV_DATE_YY = DateTime.Now.Year;


                sql = "select to_char(sysdate,'dd-MON-yyyy') as sdate,to_char(sysdate,'dd') as dd";
                sql += ",to_char(sysdate,'mm') as mm,to_char(sysdate,'yyyy') as yy from dual";
                DataTable DT_SERVERDATE = new DataTable();
                DT_SERVERDATE = Con_Oracle.ExecuteQuery(sql);

                if (DT_SERVERDATE.Rows.Count > 0)
                {
                    SERV_DATE_YY = Lib.Conv2Integer(DT_SERVERDATE.Rows[0]["yy"].ToString());
                    SERV_DATE_MM = Lib.Conv2Integer(DT_SERVERDATE.Rows[0]["mm"].ToString());
                    SERV_DATE_DD = Lib.Conv2Integer(DT_SERVERDATE.Rows[0]["dd"].ToString());
                    //SERVERDATE = DT_SERVERDATE.Rows[0]["sdate"].ToString();
                    DATABASEDATE = new DateTime(SERV_DATE_YY, SERV_DATE_MM, SERV_DATE_DD);
                }
                DT_SERVERDATE.Rows.Clear();

                sql = "";
                sql += " select max(year_code) as year_code from yearm ";
                sql += " where rec_company_code = '" + company_code + "'";

                DataTable Dt_Yr = new DataTable();
                Dt_Yr = Con_Oracle.ExecuteQuery(sql);

                if (Dt_Yr.Rows.Count > 0)
                {
                    if (Dt_Yr.Rows[0]["year_code"].ToString() != year_code)
                        IsPrevYear = true;
                }
                Dt_Yr.Rows.Clear();

                sql = "select lock_pkid from lockingm ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and lock_year = " + year_code;
                sql += " and lock_created_date = '" + DATABASEDATE.ToString(Lib.BACK_END_DATE_FORMAT) + "' ";

                DataTable Dt_Locking = new DataTable();
                Dt_Locking = Con_Oracle.ExecuteQuery(sql);

                if (IsPrevYear || Dt_Locking.Rows.Count > 0)
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    return;
                }
                Dt_Locking.Rows.Clear();


                int LockDay = 2;
                DateTime LoginMonStartDate = new DateTime(SERV_DATE_YY, SERV_DATE_MM, 1);
                if (LoginMonStartDate.DayOfWeek == DayOfWeek.Sunday)
                    LockDay++;

                if (SERV_DATE_DD >= LockDay)
                {
                    dd = 1; mm = SERV_DATE_MM; yy = SERV_DATE_YY;
                }
                else
                {
                    dd = 1; mm = SERV_DATE_MM - 1; yy = SERV_DATE_YY;
                }
                if (mm == 0)
                {
                    mm = 12;
                    yy = yy - 1;
                }
                mDate2 = new DateTime(yy, mm, dd).AddDays(-1);

                if (SERV_DATE_DD >= 5)
                {
                    dd = 1; mm = SERV_DATE_MM; yy = SERV_DATE_YY;
                }
                else
                {
                    dd = 1; mm = SERV_DATE_MM - 1; yy = SERV_DATE_YY;
                }
                if (mm == 0)
                {
                    mm = 12;
                    yy = yy - 1;
                }
                mDate5 = new DateTime(yy, mm, dd).AddDays(-1);

                /*
                if (SERV_DATE_DD >= 3)
                {
                    dd = 1; mm = SERV_DATE_MM; yy = SERV_DATE_YY;
                }
                else
                {
                    dd = 1; mm = SERV_DATE_MM - 1; yy = SERV_DATE_YY;
                }
                if (mm == 0)
                {
                    mm = 12;
                    yy = yy - 1;
                }
                mDate3 = new DateTime(yy, mm, dd).AddDays(-1);

                if (SERV_DATE_DD >= 10)
                {
                    dd = 1; mm = SERV_DATE_MM; yy = SERV_DATE_YY;
                }
                else
                {
                    dd = 1; mm = SERV_DATE_MM - 1; yy = SERV_DATE_YY;
                }
                if (mm == 0)
                {
                    mm = 12;
                    yy = yy - 1;
                }
                mDate10 = new DateTime(yy, mm, dd).AddDays(-1);

                */

                //Previous Month Locking

                //mDate_IN = mDate2.ToString("dd/MM/yyyy");
                //mDate_PN = mDate2.ToString("dd/MM/yyyy");

                //mDate_JV = mDate10.ToString("dd/MM/yyyy");
                //mDate_HO = mDate10.ToString("dd/MM/yyyy");//Costing Jv

                //After 5 Days nLocking
                /*
                sql = " select jvh_type ,to_char(mdate - 5,'DD/MM/YYYY') as mdate from ( ";
                sql += "  select distinct jvh_type ,max(jvh_date) over(partition by jvh_type) as mdate  from ledgerh ";
                sql += "  where rec_company_code = '" + company_code + "'";
                sql += "  and rec_branch_code = '" + branch_code + "'";
                sql += "  and jvh_year= " + year_code;
                sql += "  and jvh_type in ('BR','BP','CR','CP','DN','CN','DI','CI') ";
                sql += "  ) a ";
                DataTable Dt_Trans = new DataTable();
                Dt_Trans = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Trans.Rows)
                {
                    if (Dr["jvh_type"].ToString() == "BR" && !Dr["mdate"].Equals(DBNull.Value))
                        mDate_BR = Dr["mdate"].ToString();
                    if (Dr["jvh_type"].ToString() == "BP" && !Dr["mdate"].Equals(DBNull.Value))
                        mDate_BP = Dr["mdate"].ToString();
                    if (Dr["jvh_type"].ToString() == "CR" && !Dr["mdate"].Equals(DBNull.Value))
                        mDate_CR = Dr["mdate"].ToString();
                    if (Dr["jvh_type"].ToString() == "CP" && !Dr["mdate"].Equals(DBNull.Value))
                        mDate_CP = Dr["mdate"].ToString();
                    if (Dr["jvh_type"].ToString() == "DN" && !Dr["mdate"].Equals(DBNull.Value))
                        mDate_DN = Dr["mdate"].ToString();
                    if (Dr["jvh_type"].ToString() == "CN" && !Dr["mdate"].Equals(DBNull.Value))
                        mDate_CN = Dr["mdate"].ToString();
                    if (Dr["jvh_type"].ToString() == "DI" && !Dr["mdate"].Equals(DBNull.Value))
                        mDate_DI = Dr["mdate"].ToString();
                    if (Dr["jvh_type"].ToString() == "CI" && !Dr["mdate"].Equals(DBNull.Value))
                        mDate_CI = Dr["mdate"].ToString();
                }
                Dt_Trans.Rows.Clear();

                DateTime YearStartDate = DateTime.Parse(year_start_date); //yyyy-MM-dd
                if (mDate_BR == "")
                    mDate_BR = YearStartDate.ToString("dd/MM/yyyy");
                if (mDate_BP == "")
                    mDate_BP = YearStartDate.ToString("dd/MM/yyyy");
                if (mDate_CR == "")
                    mDate_CR = YearStartDate.ToString("dd/MM/yyyy");
                if (mDate_CP == "")
                    mDate_CP = YearStartDate.ToString("dd/MM/yyyy");
                if (mDate_DN == "")
                    mDate_DN = YearStartDate.ToString("dd/MM/yyyy");
                if (mDate_CN == "")
                    mDate_CN = YearStartDate.ToString("dd/MM/yyyy");
                if (mDate_DI == "")
                    mDate_DI = YearStartDate.ToString("dd/MM/yyyy");
                if (mDate_CI == "")
                    mDate_CI = YearStartDate.ToString("dd/MM/yyyy");


                mDate_IN = mDate3.ToString("dd/MM/yyyy");
                mDate_PN = mDate3.ToString("dd/MM/yyyy");
                mDate_JV = mDate3.ToString("dd/MM/yyyy");
                mDate_HO = mDate3.ToString("dd/MM/yyyy");
                mDate_BR = mDate3.ToString("dd/MM/yyyy");
                mDate_BP = mDate3.ToString("dd/MM/yyyy");
                mDate_CR = mDate3.ToString("dd/MM/yyyy");
                mDate_CP = mDate3.ToString("dd/MM/yyyy");
                mDate_DN = mDate3.ToString("dd/MM/yyyy");
                mDate_CN = mDate3.ToString("dd/MM/yyyy");
                mDate_DI = mDate3.ToString("dd/MM/yyyy");
                mDate_CI = mDate3.ToString("dd/MM/yyyy");
                */


                mDate_IN = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_PN = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_JV = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_HO = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_BR = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_BP = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_CR = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_CP = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_DN = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_CN = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_DI = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);
                mDate_CI = mDate2.ToString(Lib.BACK_END_DATE_FORMAT);




                string pkid = ""; string LockMode = "";
                sql = "select lock_pkid from lockingm ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and lock_year = " + year_code;
                DataTable dt_lock = new DataTable();
                dt_lock = Con_Oracle.ExecuteQuery(sql);
                if (dt_lock.Rows.Count <= 0)
                {
                    pkid = Guid.NewGuid().ToString().ToUpper();
                    LockMode = "ADD";
                }
                else
                {
                    pkid = dt_lock.Rows[0]["lock_pkid"].ToString();
                    LockMode = "EDIT";
                }
                dt_lock.Rows.Clear();

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("lockingm", LockMode, "lock_pkid", pkid);
                Rec.InsertString("rec_company_code", company_code);
                Rec.InsertString("rec_branch_code", branch_code);
                Rec.InsertNumeric("lock_year", year_code);
                Rec.InsertDate("lock_in", mDate_IN);
                Rec.InsertDate("lock_pn", mDate_PN);
                Rec.InsertDate("lock_dn", mDate_DN);
                Rec.InsertDate("lock_cn", mDate_CN);
                Rec.InsertDate("lock_di", mDate_DI);
                Rec.InsertDate("lock_ci", mDate_CI);
                Rec.InsertDate("lock_cr", mDate_CR);
                Rec.InsertDate("lock_cp", mDate_CP);
                Rec.InsertDate("lock_br", mDate_BR);
                Rec.InsertDate("lock_bp", mDate_BP);
                Rec.InsertDate("lock_jv", mDate_JV);
                Rec.InsertDate("lock_ho", mDate_HO);
                Rec.InsertDate("lock_created_date", DATABASEDATE);

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                bTrans = true;
                Con_Oracle.ExecuteNonQuery(sql);

                //Invoices 
                string mdate = DateTime.Now.AddDays(-1).ToString(Lib.BACK_END_DATE_FORMAT);
                sql = "update ledgerh  set jvh_edit_code = null ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and jvh_year = " + year_code;
                sql += " and jvh_type in ('IN','PN') ";
                sql += " and jvh_date <= '" + mdate + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                //Transactions
                sql = "update ledgerh  set jvh_edit_code = null ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and jvh_year = " + year_code;
                sql += " and (jvh_date <= '" + mDate2.ToString(Lib.BACK_END_DATE_FORMAT) + "' or nvl(rec_aprvd,'N')='Y' )";// For BP, Chq Approval could be locked
                sql += " and jvh_type not in ('OI') ";//,'OP'
                Con_Oracle.ExecuteNonQuery(sql);

                //Costing
                sql = "update costingm  set cost_edit_code = null ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and cost_year = " + year_code;
                sql += " and cost_date <= '" + mDate2.ToString(Lib.BACK_END_DATE_FORMAT) + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                //Operations
                sql = "update hblm  set hbl_edit_code = null ";
                sql += " where exists (select 1 from ledgerh where hbl_pkid = jvh_cc_id ";
                sql += " and rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and jvh_year = " + year_code;
                sql += " and jvh_type in ('PN', 'IN') ";
                sql += " and jvh_date <= '" + mDate2.ToString(Lib.BACK_END_DATE_FORMAT) + "'";
                sql += " )";
                Con_Oracle.ExecuteNonQuery(sql);

                //Folder lock Master lock
                sql = "update hblm  set hbl_edit_code = null ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and hbl_year = " + year_code;
                sql += " and hbl_type in ('MBL-AE', 'MBL-AI','MBL-SE', 'MBL-SI') ";
                sql += " and hbl_folder_sent_date is not null ";
                Con_Oracle.ExecuteNonQuery(sql);

                //Folder lock House Lock (jobincome)
                sql = " update hblm h set h.hbl_edit_code = null ";
                sql += "  where exists (select 1 from hblm m where h.hbl_mbl_id = m.hbl_pkid ";
                sql += "  and m.rec_company_code = '" + company_code + "'";
                sql += "  and m.rec_branch_code = '" + branch_code + "'";
                sql += "  and m.hbl_year = " + year_code;
                sql += "  and m.hbl_type in ('MBL-AE', 'MBL-AI','MBL-SE', 'MBL-SI') ";
                sql += "  and m.hbl_folder_sent_date is not null";
                sql += "  )";
                Con_Oracle.ExecuteNonQuery(sql);

                //Folder lock Mbl Invoice
                sql = " update ledgerh set jvh_edit_code=null ";
                sql += "  where exists (select 1 from hblm where jvh_cc_id=hbl_pkid";
                sql += "  and rec_company_code = '" + company_code + "'";
                sql += "  and rec_branch_code = '" + branch_code + "'";
                sql += "  and hbl_year = " + year_code;
                sql += "  and hbl_type in ('MBL-AE', 'MBL-AI','MBL-SE', 'MBL-SI') ";
                sql += "  and hbl_folder_sent_date is not null";
                sql += "  )";
                Con_Oracle.ExecuteNonQuery(sql);

                // lock salarymaster, payroll 
                sql = " update salarym set sal_edit_code = null ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and (sal_month = sal_year or sal_date <= '" + mDate5.ToString(Lib.BACK_END_DATE_FORMAT) + "')";
                Con_Oracle.ExecuteNonQuery(sql);

                // lock leavemaster, leavedetails
                sql = " update leavem set lev_edit_code = null ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and (lev_month = 0 or lev_date <= '" + mDate5.ToString(Lib.BACK_END_DATE_FORMAT) + "')";
                Con_Oracle.ExecuteNonQuery(sql);

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    if (bTrans)
                        Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
        }
    }
}
