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
    public class CompanyService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Companym> mList = new List<Companym>();
            Companym mRow;

            string type = SearchData["type"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string rowtype = SearchData["rowtype"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();


            string user_pkid = SearchData["user_pkid"].ToString();

            string region_id = SearchData["region_id"].ToString();
            string vendor_id = SearchData["vendor_id"].ToString();
            string role_name = SearchData["role_name"].ToString();


            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            Boolean rights_admin = ( Boolean)SearchData["rights_admin"];
            long startrow = 0;
            long endrow = 0;

            string sql2 = "";

            try
            {
                sWhere = " where  a.comp_type ='" + rowtype + "' "; ;

                if (rowtype == "B" || rowtype == "S" || rowtype == "V")
                    sWhere += "  and a.rec_company_code = '" + comp_code + "'";
                if (rights_admin == false)
                {
                    if (rowtype == "C" )
                        sWhere += " and a.comp_code = '" + comp_code + "'";
                    if (rowtype == "S")
                    {
                        if (role_name == "ZONE ADMIN" || role_name == "SALES EXECUTIVE" || role_name == "VENDOR" || role_name == "RECCE USER")
                        {
                            sql2 = " inner join userd c on a.comp_pkid = c.user_branch_id and c.user_id = '" + user_pkid + "'";
                        }
                    }

                    if (rowtype == "V")
                    {
                        if (role_name == "ZONE ADMIN" || role_name == "SALES EXECUTIVE")
                            sWhere += " and a.comp_region_id = '" + region_id + "'";
                        if (role_name == "VENDOR" || role_name == "RECCE USER")
                            sWhere += " and a.comp_pkid = '" + vendor_id + "'";
                    }

                }
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  a.comp_code like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or a.comp_name like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM companym a ";
                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM companym a ";
                    sql += sql2;
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
                sql += "  select  a.comp_pkid, a.comp_code, a.comp_name, a.comp_short_name, a.comp_type, b.comp_name as comp_parent_name,a.comp_address1, a.comp_district, a.comp_state ";
                sql += "  ,a.comp_web,a.comp_ptc,a.comp_email,a.comp_tel,a.comp_mobile,a.comp_fax,a.comp_order, a.comp_approver_email, a.comp_receiver_email, region.param_name as comp_region_name ";
                sql += "  ,row_number() over(order by a.comp_name) rn ";
                sql += "  from companym a left join companym b on a.comp_parent_id = b.comp_pkid ";
                sql += "  left join param region on a.comp_region_id = region.param_pkid ";
                sql += sql2;
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by a.comp_name ";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Companym();
                    mRow.comp_pkid = Dr["comp_pkid"].ToString();
                    mRow.comp_code = Dr["comp_code"].ToString();
                    mRow.comp_name = Dr["comp_name"].ToString();
                    mRow.comp_short_name = Dr["comp_short_name"].ToString();
                    mRow.comp_address1 = Dr["comp_address1"].ToString();
                    mRow.comp_type = Dr["comp_type"].ToString();
                    mRow.comp_district = Dr["comp_district"].ToString();
                    mRow.comp_state = Dr["comp_state"].ToString();
                    mRow.comp_parent_name = Dr["comp_parent_name"].ToString();
                    mRow.comp_region_name = Dr["comp_region_name"].ToString();
                    mRow.comp_web = Dr["comp_web"].ToString();
                    mRow.comp_ptc = Dr["comp_ptc"].ToString();
                    mRow.comp_email = Dr["comp_email"].ToString();
                    mRow.comp_approver_email = Dr["comp_approver_email"].ToString();
                    mRow.comp_receiver_email = Dr["comp_receiver_email"].ToString();
                    mRow.comp_tel = Dr["comp_tel"].ToString();
                    mRow.comp_mobile = Dr["comp_mobile"].ToString();
                    mRow.comp_fax = Dr["comp_fax"].ToString();
                    mRow.comp_order = Lib.Conv2Integer(Dr["comp_order"].ToString());
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
                     

        public Dictionary<string, object>  GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Companym mRow =new Companym();
            
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  a.comp_pkid, a.comp_code, a.comp_name, comp_short_name, comp_type, comp_parent_id,a.comp_district, a.comp_state ";
                sql += " ,comp_address1,comp_address2,comp_address3";
                sql += " ,comp_tel,comp_fax,comp_web,comp_email, comp_approver_email, comp_receiver_email";
                sql += " ,comp_ptc,comp_mobile, comp_logo_name, comp_image_name";
                sql += " ,comp_order, comp_region_id, region.param_name as comp_region_name, comp_location ";
                sql += " from companym a  ";
                sql += " left join param region on a.comp_region_id = region.param_pkid ";
                sql += " where  a.comp_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Companym();
                    mRow.comp_pkid = Dr["comp_pkid"].ToString();
                    mRow.comp_code = Dr["comp_code"].ToString();
                    mRow.comp_name = Dr["comp_name"].ToString();
                    mRow.comp_short_name = Dr["comp_short_name"].ToString();
                    mRow.comp_type = Dr["comp_type"].ToString();
                    mRow.comp_parent_id = Dr["comp_parent_id"].ToString();
                    mRow.comp_address1 = Dr["comp_address1"].ToString();
                    mRow.comp_address2 = Dr["comp_address2"].ToString();
                    mRow.comp_address3 = Dr["comp_address3"].ToString();

                    mRow.comp_location = Dr["comp_location"].ToString();
                    mRow.comp_district = Dr["comp_district"].ToString();
                    mRow.comp_state = Dr["comp_state"].ToString();

                    mRow.comp_region_id = Dr["comp_region_id"].ToString();
                    mRow.comp_region_name = Dr["comp_region_name"].ToString();

                    mRow.comp_tel = Dr["comp_tel"].ToString();
                    mRow.comp_fax = Dr["comp_fax"].ToString();
                    mRow.comp_web = Dr["comp_web"].ToString();
                    mRow.comp_email = Dr["comp_email"].ToString();


                    mRow.comp_logo_name = Dr["comp_logo_name"].ToString();
                    mRow.comp_logo_uploaded = false;
                    if (Dr["comp_logo_name"].ToString().Trim().Length > 0)
                        mRow.comp_logo_uploaded = true;

                    mRow.comp_image_name = Dr["comp_image_name"].ToString();
                    mRow.comp_image_uploaded = false;
                    if (Dr["comp_image_name"].ToString().Trim().Length > 0)
                        mRow.comp_image_uploaded = true;


                    mRow.comp_approver_email = Dr["comp_approver_email"].ToString();
                    mRow.comp_receiver_email = Dr["comp_receiver_email"].ToString();

                    mRow.comp_ptc = Dr["comp_ptc"].ToString();
                    mRow.comp_mobile = Dr["comp_mobile"].ToString();
                    mRow.comp_order = Lib.Conv2Integer(Dr["comp_order"].ToString());
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

        public string AllValid(Companym Record)
        {
            string str = "";
            try
            {
                sql = "select comp_pkid from (";
                sql += "select comp_pkid  from companym a where (a.comp_code = '{CODE}')  ";
                if (Record.comp_type != "C")
                {
                    sql += " and a.rec_company_code = '" + Record._globalvariables.comp_code + "'" ;
                }
                sql += ") a where comp_pkid <> '{PKID}'";

                sql = sql.Replace("{CODE}", Record.comp_code);
                sql = sql.Replace("{PKID}", Record.comp_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";


            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Companym Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            string scode = "";
            int iOrder = 0;
            try
            {
                Con_Oracle = new DBConnection();


                if (Record.rec_mode == "ADD" && Record.comp_type == "C")
                    Lib.AddError(ref ErrorMessage, "New Company Creation is Disabled");

                if (Record.comp_code.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Code Cannot Be Empty");
                if (Record.comp_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");

                if (Record.comp_short_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Short Name Cannot Be Empty");

                if (Record.comp_type != "C")
                {
                    if (Record.comp_parent_id is null)
                        Lib.AddError(ref ErrorMessage, "Parent Company Cannot Be Blank");

                    if (Record.comp_region_id =="")
                        Lib.AddError(ref ErrorMessage, "Region Cannot Be Blank");

                }

                if (Record.comp_state.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "State Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ( (ErrorMessage = AllValid(Record)) != "" ) 
                    throw new Exception(ErrorMessage);

                Record.comp_code = Record.comp_code.ToString().Trim().Replace(" ","");

                if (Record.rec_mode == "ADD" && Record.comp_type == "C")
                {
                    sql = "select isnull(max(comp_order),0) + 1 as slno from companym where comp_type ='C'";
                    iOrder = Lib.Conv2Integer(Con_Oracle.ExecuteScalar(sql).ToString());
                }

                if (Record.rec_mode == "ADD" && Record.comp_type != "C")
                {
                    sql = "select isnull(max(comp_order),0) + 1 as slno from companym where comp_parent_id ='" +  Record.comp_parent_id+"'";
                    iOrder =  Lib.Conv2Integer(Con_Oracle.ExecuteScalar(sql).ToString());
                }
                if (Record.rec_mode == "ADD" && Record.comp_type != "C")
                {
                    sql = "select max(comp_code) as comp_code from companym where comp_pkid = '" + Record.comp_parent_id + "'";
                    scode = Con_Oracle.ExecuteScalar(sql).ToString();
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("companym", Record.rec_mode, "comp_pkid", Record.comp_pkid);
                Rec.InsertString("comp_name", Record.comp_name);
                Rec.InsertString("comp_short_name", Record.comp_short_name);

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("comp_code", Record.comp_code);
                    Rec.InsertString("comp_type", Record.comp_type);
                    Rec.InsertString("rec_printed", "N");
                    Rec.InsertString("rec_locked", "N");
                    Rec.InsertString("rec_updated", "N");
                    if (Record.comp_type == "C")
                    {
                        Rec.InsertString("rec_company_code", Record.comp_code);
                    }
                    if (Record.comp_type != "C")
                    {
                        Rec.InsertString("comp_parent_id", Record.comp_parent_id);
                        Rec.InsertString("rec_company_code", scode);
                        Rec.InsertString("rec_branch_code", Record.comp_code);
                    }
                    Rec.InsertNumeric("comp_order", iOrder.ToString());
                }
                else
                    Rec.InsertNumeric("comp_order", Record.comp_order.ToString());

                Rec.InsertString("comp_address1", Record.comp_address1);
                Rec.InsertString("comp_address2", Record.comp_address2);
                Rec.InsertString("comp_address3", Record.comp_address3);

                Rec.InsertString("comp_location", Record.comp_location);
                Rec.InsertString("comp_district", Record.comp_district);
                Rec.InsertString("comp_state", Record.comp_state);

                Rec.InsertString("comp_region_id", Record.comp_region_id);

                Rec.InsertString("comp_tel", Record.comp_tel);
                Rec.InsertString("comp_fax", Record.comp_fax);
                Rec.InsertString("comp_email", Record.comp_email, "L");
                Rec.InsertString("comp_web", Record.comp_web,"L");

                Rec.InsertString("comp_logo_name", Record.comp_logo_name);
                Rec.InsertString("comp_image_name", Record.comp_image_name);

                Rec.InsertString("comp_ptc", Record.comp_ptc);
                Rec.InsertString("comp_mobile", Record.comp_mobile);

                Rec.InsertString("comp_approver_email", Record.comp_approver_email, "L");
                Rec.InsertString("comp_receiver_email", Record.comp_receiver_email, "L");



                sql = Rec.UpdateRow();
               
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                if ( Record.rec_mode == "ADD" && Record.comp_type == "C")
                    CreateModules( Record._globalvariables.comp_code,Record.comp_code, Record.comp_pkid, Record._globalvariables.year_code);
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


        private void CreateModules(string oldcompcode, string newcompcode, string newcompid, string yearcode)
        {
            DataTable dt = new DataTable();
            DataTable dtmenu = new DataTable();
            string sql1 = "";
            string sql = "";

            string sid = "";

            // Create Admin User
            sql = "";
            sql += " insert into userm( ";
            sql += " USER_PKID, USER_CODE, USER_NAME, USER_PASSWORD, USER_EMAIL, USER_ISSUPERVISOR, USER_ISLOCKED, REC_DELETED, REC_UPDATED, REC_ORIGIN, REC_PRINTED, REC_LOCKED, REC_COMPANY_CODE, REC_BRANCH_CODE, REC_CATEGORY, REC_CREATED_BY, REC_CREATED_DATE, REC_EDITED_BY, REC_EDITED_DATE, REC_DELETED_BY, REC_DELETED_DATE, REC_VERSION, USER_BRANCH_ID, USER_COMPANY_ID ";
            sql += " ) ";
            sql += " select ";
            sql += "'" + System.Guid.NewGuid().ToString().ToUpper() + "' as  USER_PKID, ";
            sql += " USER_CODE, USER_NAME, USER_PASSWORD, USER_EMAIL, USER_ISSUPERVISOR, USER_ISLOCKED, REC_DELETED, ";
            sql += " REC_UPDATED, REC_ORIGIN, REC_PRINTED, REC_LOCKED, ";
            sql += "'" + newcompcode + "' as  REC_COMPANY_CODE, ";
            sql += " REC_BRANCH_CODE, REC_CATEGORY, ";
            sql += " REC_CREATED_BY, REC_CREATED_DATE, REC_EDITED_BY, REC_EDITED_DATE, REC_DELETED_BY, REC_DELETED_DATE, REC_VERSION, ";
            sql += " NULL, ";
            sql += "'" + newcompid + "' as USER_COMPANY_ID ";
            sql += " from userm where user_code = 'ADMIN' and rec_company_code = '" + oldcompcode + "'";
            Con_Oracle.ExecuteNonQuery(sql);


            // Create Modules
            sql1 = "";
            sql1 += " insert into modulem(module_pkid, module_name, module_order, rec_company_code) ";
            sql1 += " values ('{module_pkid}', '{module_name}', {module_order}, '{rec_company_code}')";

            sql = "";
            sql += "select module_pkid, '' as id, module_name, module_order, rec_company_code from modulem ";
            sql += " where rec_company_code = '" + oldcompcode + "' order by module_order";
            dt = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow Dr in dt.Rows)
            {
                sql = sql1;
                sid = System.Guid.NewGuid().ToString().ToUpper();
                Dr["id"] = sid;
                sql = sql.Replace("{module_pkid}", sid);
                sql = sql.Replace("{module_name}", Dr["module_name"].ToString());
                sql = sql.Replace("{module_order}", Dr["module_order"].ToString());
                sql = sql.Replace("{rec_company_code}", newcompcode);
                Con_Oracle.ExecuteNonQuery(sql);
            }

            // Create Menu
            sql1 = "";
            sql1 += " insert into menum(menu_module_id, menu_pkid, menu_code, menu_name, menu_route1, menu_route2,menu_type, menu_displayed, menu_order, rec_company_code) ";
            sql1 += " values ('{menu_module_id}', '{menu_pkid}', '{menu_code}', '{menu_name}', '{menu_route1}', '{menu_route2}','{menu_type}', '{menu_displayed}', {menu_order}, '{rec_company_code}')";

            sql = "";
            sql += "select menu_module_id, menu_pkid, menu_code, menu_name, menu_route1, menu_route2, menu_type, menu_displayed, menu_order, rec_company_code ";
            sql += " from menum where rec_company_code = '" + oldcompcode + "' and menu_code not like '~%'";
            sql += " order by menu_order ";

            dtmenu = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow Dr in dtmenu.Rows)
            {
                sid = "";
                foreach (DataRow dr1 in dt.Select("module_pkid='" + Dr["menu_module_id"].ToString() + "'"))
                    sid = dr1["id"].ToString();
                if (sid == "")
                    continue;
                sql = sql1;
                sql = sql.Replace("{menu_pkid}", System.Guid.NewGuid().ToString().ToUpper());
                sql = sql.Replace("{menu_module_id}", sid);
                sql = sql.Replace("{menu_code}", Dr["menu_code"].ToString());
                sql = sql.Replace("{menu_name}", Dr["menu_name"].ToString());
                sql = sql.Replace("{menu_route1}", Dr["menu_route1"].ToString());
                sql = sql.Replace("{menu_route2}", Dr["menu_route2"].ToString());
                sql = sql.Replace("{menu_type}", Dr["menu_type"].ToString());
                sql = sql.Replace("{menu_displayed}", Dr["menu_displayed"].ToString());
                sql = sql.Replace("{menu_order}", Dr["menu_order"].ToString());
                sql = sql.Replace("{rec_company_code}", newcompcode);
                Con_Oracle.ExecuteNonQuery(sql);
            }


            // Create Settings
            sql1 = "";
            sql1 += " insert into settings(rec_company_code,parentid, tablename,caption, id, code, name, tabletype) ";
            sql1 += " select '" + newcompcode + "' as rec_company_code, '" + newcompcode + "' as parentid, tablename,caption, id, code, name, tabletype from settings  ";
            sql1 += " where rec_company_code = '" + oldcompcode + "'";
            sql1 +="  and parentid = '" + oldcompcode + "'";

            Con_Oracle.ExecuteNonQuery(sql1);

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
                sql += " select comp_pkid,comp_name from companym where comp_type ='C' order by comp_name";

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
            RetData.Add("list", mList);

            return RetData;
        }


        public IDictionary<string, object> Delete(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();

            string pkid = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string comp_type = SearchData["comp_type"].ToString();

            DataTable Dt_Tables = new DataTable();


            sql = " select count(*) as tot from companym  where comp_parent_id = '" + pkid + "'";
            int iTot =  (int) Con_Oracle.ExecuteScalar(sql);
            if ( iTot > 0 )
            {
                throw new Exception("Record Exists In Child records");
            }

            sql = " select count(*) as tot from pim_docm  where doc_store_id = '" + pkid + "'";
            iTot = (int)Con_Oracle.ExecuteScalar(sql);
            if (iTot > 0)
            {
                throw new Exception("Record Exists In Child Table");
            }

            sql = " select count(*) as tot from  pim_spotm  where spot_store_id = '" + pkid + "'";
            iTot = (int)Con_Oracle.ExecuteScalar(sql);
            if (iTot > 0)
            {
                throw new Exception("Record Exists In Spot Table");
            }


            try
            {
                Con_Oracle.BeginTransaction();

                if (comp_type == "C")
                {
                    sql = " select tab_table_name  from tablesm  where rec_company_code = '" + comp_code + "'";
                    Dt_Tables = Con_Oracle.ExecuteQuery(sql);

                    foreach (DataRow Dr in Dt_Tables.Rows)
                    {
                        sql = "DROP TABLE " + Dr["tab_table_name"];
                        Con_Oracle.ExecuteNonQuery(sql);
                    }

                    sql = "delete from tablesd where rec_company_code = '" + comp_code + "'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "delete from tablesm where rec_company_code = '" + comp_code + "'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "delete from param where rec_company_code = '" + comp_code + "'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "delete from pim_groupm where rec_company_code = '" + comp_code + "'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "delete from modulem where rec_company_code = '" + comp_code + "'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "delete from userm where rec_company_code = '" + comp_code + "'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "delete from menum where rec_company_code = '" + comp_code + "'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "delete from userrights  where rights_company_id = '" + pkid + "'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "delete from settings where parentid = '" + comp_code + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                }

                sql = "delete from companym where comp_pkid = '" + pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            return RetData;
        }


        public IDictionary<string, object> LoadUserStore(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Companym> mList = new List<Companym>();
            Companym mRow;

            string type = SearchData["type"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string user_pkid = SearchData["user_pkid"].ToString();

            string userid = SearchData["userid"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            //Boolean rights_admin = (Boolean)SearchData["rights_admin"];

            string region_id = SearchData["region_id"].ToString();
            string vendor_id = SearchData["vendor_id"].ToString();
            string role_name = SearchData["role_name"].ToString();

            string sql2 = "";

            long startrow = 0;
            long endrow = 0;

            string user_role_name = "";
            string user_region_id = "";
            try
            {
                DataTable Dt_userm = new DataTable();
                sql = "select user_region_id,role.param_name as user_role_name from userm ";
                sql += " left join param role on user_role_id = role.param_pkid ";
                sql += " where user_pkid = '" + userid + "' ";
                Dt_userm = Con_Oracle.ExecuteQuery(sql);
                foreach ( DataRow Dr in Dt_userm.Rows)
                {
                    user_region_id = Dr["user_region_id"].ToString();
                    user_role_name = Dr["user_role_name"].ToString();
                    break;
                }

                sWhere = " where a.rec_company_code ='" + comp_code + "' and a.comp_type ='S' ";
                if (user_role_name != "SUPER ADMIN")
                {
                    sWhere += " and a.comp_region_id  = '" + user_region_id +"'";
                }
                if (role_name == "ZONE ADMIN" || role_name == "VENDOR")
                {
                    sql2 = " inner join userd c on a.comp_pkid = c.user_branch_id and c.user_id = '" + user_pkid + "'";
                }


                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  a.comp_name like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or region.param_name like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    if (Con_Oracle.DB == "SQL")
                    {
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total ";
                        sql += " from companym a ";
                        sql += " left join userd b on a.comp_pkid = b.user_branch_id  and b.user_branch_id = '" + userid + "'";
                        sql += "  left join param region on a.comp_region_id = region.param_pkid ";
                        sql += sql2;

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
                sql += " select * from ( ";
                sql += "  select  a.comp_pkid, a.comp_code, a.comp_name, a.comp_address1, a.comp_state,  ";
                sql += "  region.param_name as comp_region_name,";
                sql += "  b.user_branch_id, b.user_id, row_number() over(order by a.comp_name) rn ";
                sql += "  from companym a ";
                sql += "  left join userd b on a.comp_pkid = b.user_branch_id and b.user_id = '" + userid + "'";
                sql += "  left join param region on a.comp_region_id = region.param_pkid ";
                sql += sql2;
                sql +=    sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by comp_name ";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Companym();


                    mRow.comp_pkid = Dr["comp_pkid"].ToString();
                    mRow.comp_code = Dr["comp_code"].ToString();
                    mRow.comp_name = Dr["comp_name"].ToString();
                    mRow.comp_address1 = Dr["comp_address1"].ToString();
                    mRow.comp_state = Dr["comp_state"].ToString();
                    mRow.comp_region_name = Dr["comp_region_name"].ToString();
                    mRow.selected = true;
                    if (Dr["user_id"].Equals(DBNull.Value))
                        mRow.selected = false;

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


        public Dictionary<string, object> SaveUserStore(Companym Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {
                Con_Oracle = new DBConnection();
                Con_Oracle.BeginTransaction();


                sql = "delete from userd where rec_company_code = '" + Record._globalvariables.comp_code + "'";
                sql += " and user_id = '" + Record.user_id + "'";
                sql += " and user_branch_id = '" + Record.comp_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                if (Record.selected)
                {
                    sql = "insert into userd(user_id, user_branch_id, rec_type,rec_company_code) values (";
                    sql += " '" + Record.user_id + "',";
                    sql += " '" + Record.comp_pkid + "',";
                    sql += " 'S',";
                    sql += " '" + Record._globalvariables.comp_code + "'";
                    sql += ")";
                    Con_Oracle.ExecuteNonQuery(sql);
                }

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



        public IDictionary<string, object> ListApproval(Dictionary<string, object> SearchData)
        {
         
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<approvald> mList = new List<approvald>();
            approvald mRow;

            string pkid = SearchData["pkid"].ToString();
            try
            {

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select ad_pkid, ad_parent_id, ad_by, ad_date, ad_remarks, ad_status from approvald where ad_parent_id ='" + pkid + "' order by ad_date";

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new approvald();
                    mRow.ad_pkid = Dr["ad_pkid"].ToString();
                    mRow.ad_parent_id = Dr["ad_parent_id"].ToString();
                    mRow.ad_by = Dr["ad_by"].ToString();
                    mRow.ad_remarks = Dr["ad_remarks"].ToString();
                    mRow.ad_status = Dr["ad_status"].ToString();
                    mRow.ad_date = Lib.DatetoStringDisplayformat(Dr["ad_date"]);
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


        public Dictionary<string, object> SaveApproval(approvald Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string sql2 = "";
            string sql3 = "";

            try
            {
                Con_Oracle = new DBConnection();

                /*
                if (Record.comp_code.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Code Cannot Be Empty");
                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);
                */

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("approvald", "ADD", "ad_pkid", System.Guid.NewGuid().ToString().ToUpper());
                Rec.InsertString("ad_parent_id", Record.ad_parent_id);
                Rec.InsertString("ad_by", Record.ad_by);
                Rec.InsertString("ad_remarks", Record.ad_remarks);
                Rec.InsertString("ad_status", Record.ad_status);
                Rec.InsertFunction("ad_date", "getdate()");
                sql = Rec.UpdateRow();

                sql2 = "delete from approvalm where am_pkid ='" + Record.ad_parent_id + "'";

                Rec = new DBRecord();
                Rec.CreateRow("approvalm", "ADD", "am_pkid", Record.ad_parent_id);
                Rec.InsertString("am_by", Record.ad_by);
                Rec.InsertString("am_remarks", Record.ad_remarks);
                Rec.InsertString("am_status", Record.ad_status);
                Rec.InsertFunction("am_date", "getdate()");

                sql3 = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.ExecuteNonQuery(sql2);
                Con_Oracle.ExecuteNonQuery(sql3);
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


        public void test1()
        {

        }


    }
}
