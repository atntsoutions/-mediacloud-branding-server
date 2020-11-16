using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase.Connections;


namespace DataBase
{
    public class LovService : BL_Base
    {
        // This is for Combo Loading
        public IDictionary<string, object> Lov(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>(); 

            string Table = "";
            string pkid = "";
            string comp_code = "";
            string branch_code = "";

            string subtype= "";

            if (SearchData.ContainsKey("table"))
                Table = SearchData["table"].ToString().ToUpper();

            if (SearchData.ContainsKey("pkid"))
                pkid = SearchData["pkid"].ToString().ToUpper();

            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            if (SearchData.ContainsKey("branch_code"))
                branch_code = SearchData["branch_code"].ToString();

            if (SearchData.ContainsKey("subtype"))
                subtype = SearchData["subtype"].ToString();


            DataTable Dt_List = new DataTable();

            Con_Oracle = new DBConnection();

            try
            {

                if (Table == "BRANCH")
                {
                    sql = "select  comp_pkid ,comp_code, comp_name from companym a";
                    sql += " where a.rec_company_code = '" + comp_code + "'";
                    sql += " and a.comp_type = 'B'";
                    sql += " order by a.comp_name ";

                    List<Companym> mList = new List<Companym>();
                    Companym Row;
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        Row = new Companym();
                        Row.comp_pkid = Dr["comp_pkid"].ToString();
                        Row.comp_code = Dr["comp_code"].ToString();
                        Row.comp_name = Dr["comp_name"].ToString();
                        mList.Add(Row);
                    }
                    RetData.Add(Table.ToLower(), mList);
                }

                if (Table == "PARAM")
                {
                    sql = "select param_pkid, param_code, param_name   from param  a  ";
                    sql += " where a.rec_company_code = '" + comp_code + "'";
                    if (SearchData.ContainsKey("param_pkid"))
                        sql += " and a.param_pkid = '" + SearchData["param_pkid"].ToString() + "'";
                    if (SearchData.ContainsKey("param_type"))
                        sql += " and a.param_type = '" + SearchData["param_type"].ToString() + "'";
                    if (SearchData.ContainsKey("param_code"))
                        sql += " and a.param_code = '" + SearchData["param_code"].ToString() + "'";
                    if (SearchData.ContainsKey("param_name"))
                        sql += " and a.param_name = '" + SearchData["param_name"].ToString() + "'";

                    if (SearchData["param_type"].ToString() == "ESANCHITDOC")
                    {
                        if ( Con_Oracle.DB == "ORACLE")
                            sql += " and nvl(a.rec_locked,'N') = 'N' ";
                        else
                            sql += " and isnull(a.rec_locked,'N') = 'N' ";

                    }


                    sql += " order by a.param_name ";

                    List<Param> mList = new List<Param>();
                    Param Row;
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        Row = new Param();
                        Row.param_pkid = Dr["param_pkid"].ToString();
                        Row.param_code = Dr["param_code"].ToString();
                        Row.param_name = Dr["param_name"].ToString();
                        mList.Add(Row);
                    }
                    RetData.Add(Table.ToLower(), mList);
                }
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

        // This is for searching records
        public IDictionary<string, object> SearchRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string Table = "";


            string []_str;
            string _code = "";
            int _year = 0;

            Boolean isOldYear = false;


            if (SearchData.ContainsKey("table"))
                Table = SearchData["table"].ToString().ToUpper();


            DataTable Dt_List = new DataTable();

            Con_Oracle = new DBConnection();

            try
            {
                if (Table == "PARAM")
                {
                    sql = "select param_pkid, param_type, param_code, param_name,param_id1   from param  a  ";
                    sql += " where 1=1 ";
                    if (SearchData.ContainsKey("comp_code"))
                        sql += " and a.rec_company_code = '" + SearchData["comp_code"].ToString() + "'";
                    if (SearchData.ContainsKey("param_pkid"))
                        sql += " and a.param_pkid = '" + SearchData["param_pkid"].ToString() + "'";
                    if (SearchData.ContainsKey("param_type"))
                        sql += " and a.param_type = '" + SearchData["param_type"].ToString() + "'";
                    if (SearchData.ContainsKey("param_code"))
                        sql += " and a.param_code = '" + SearchData["param_code"].ToString() + "'";
                    if (SearchData.ContainsKey("param_name"))
                        sql += " and a.param_name = '" + SearchData["param_name"].ToString() + "'";
                    sql += " order by a.param_name ";

                    List<Param> ParamList = new List<Param>();
                    Param paramRow;
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        paramRow = new Param();
                        paramRow.param_pkid = Dr["param_pkid"].ToString();
                        paramRow.param_type = Dr["param_type"].ToString();
                        paramRow.param_code = Dr["param_code"].ToString();
                        paramRow.param_name = Dr["param_name"].ToString();
                        paramRow.param_id1 = Dr["param_id1"].ToString();
                        ParamList.Add(paramRow);
                    }
                    RetData.Add(Table.ToLower(), ParamList);
                }


           
                if (Table == "HISTORY")
                {
                    List<Auditlog> aList = new List<Auditlog>();

                    string pkid = SearchData["pkid"].ToString().Trim();
                    if (pkid != "")
                    {
                        sql = " select to_char(audit_date,'DD/MM/YYYY HH24:MI:SS') as auditdate,audit_action,audit_user_code";
                        sql += " ,audit_type,audit_refno,audit_remarks  ";
                        sql += " from auditlog ";
                        sql += " where audit_pkey = '" + pkid + "'";
                        sql += " order by audit_date ";

                        DataTable dt_audit = new DataTable();
                        dt_audit = Con_Oracle.ExecuteQuery(sql);
                        Auditlog _aRec;
                        foreach (DataRow Dr in dt_audit.Rows)
                        {
                            _aRec = new Auditlog();
                            _aRec.audit_date = Dr["auditdate"].ToString();
                            _aRec.audit_action = Dr["audit_action"].ToString();
                            _aRec.audit_user_code = Dr["audit_user_code"].ToString();
                            _aRec.audit_type = Dr["audit_type"].ToString();
                            _aRec.audit_refno = Dr["audit_refno"].ToString();
                            _aRec.audit_remarks = Dr["audit_remarks"].ToString();
                            aList.Add(_aRec);
                        }
                    }
                    RetData.Add("list", aList);
                }
                
                if (Table == "PASTEDATA")
                {

                    char schar = '\t';
                    decimal nAmt = 0;
                    string[] srow;
                    string[] sdata;
                    string stype = "";
                    string snos = "";
                    string comp_code = "";
                    string branch_code = "";
                    string year_code = "";
                    string isoldyear = "N";
                    decimal nTotal = 0;

                    List<LovTable> mList = new List<LovTable>();
                    LovTable _Rec = null;

                    DataTable dt_job = new DataTable();

                    stype = SearchData["type"].ToString().ToUpper();
                    comp_code = SearchData["comp_code"].ToString().ToUpper();
                    branch_code = SearchData["branch_code"].ToString().ToUpper();
                    year_code = SearchData["year_code"].ToString();

                    if (SearchData.ContainsKey("isoldyear"))
                        isoldyear = SearchData["isoldyear"].ToString();

                    if (SearchData["cbdata"].ToString().Contains(","))
                        schar = ',';

                    srow = SearchData["cbdata"].ToString().Split('\n');

                    foreach (string str in srow)
                    {
                        if (str.Length > 0)
                        {
                            sdata = str.Split(schar);
                            if (sdata.Length > 0)
                            {
                                if (snos.Length > 0)
                                    snos += ",";
                                if (stype == "JOB SEA EXPORT" || stype == "JOB AIR EXPORT" || stype == "SI SEA IMPORT" || stype == "SI AIR IMPORT" || stype == "SI SEA EXPORT" || stype == "SI AIR EXPORT"|| stype == "GENERAL JOB")
                                    snos += sdata[0].ToString();
                                else
                                    snos += "'" + sdata[0].ToString() + "'";
                            }
                        }
                    }
                    if (stype == "JOB SEA EXPORT" || stype == "JOB AIR EXPORT")
                    {
                        sql = " select job_pkid as id, job_docno as code, job_docno as name, job_year as year, 0 as amt from jobm  where ";
                        sql += " rec_branch_code = '" + branch_code + "'";
                        sql += " and job_year = " + year_code;
                        if (stype == "JOB SEA EXPORT")
                            sql += " and rec_category ='SEA EXPORT'  ";
                        if (stype == "JOB AIR EXPORT")
                            sql += " and rec_category ='AIR EXPORT'  ";
                        sql += " and job_docno in (" + snos + ")";
                        sql += " order by job_docno ";
                    }

                    if (stype == "SI SEA IMPORT" || stype == "SI AIR IMPORT" || stype == "SI SEA EXPORT" || stype == "SI AIR EXPORT")
                    {
                        sql = " select hbl_pkid as id, hbl_no as code, hbl_no as name, hbl_year as year, 0 as amt from hblm  where ";
                        sql += " rec_branch_code = '" + branch_code + "'";
                        sql += " and hbl_year = " + year_code;
                        if (stype == "SI SEA EXPORT")
                            sql += " and hbl_type ='HBL-SE' ";
                        if (stype == "SI AIR EXPORT")
                            sql += " and hbl_type ='HBL-AE' ";
                        if (stype == "SI SEA IMPORT")
                            sql += " and hbl_type ='HBL-SI' ";
                        if (stype == "SI AIR IMPORT")
                            sql += " and hbl_type ='HBL-AI' ";

                        sql += " and hbl_no in (" + snos + ")";
                        sql += " order by hbl_no ";
                    }

                    if (stype == "GENERAL JOB")
                    {
                        sql = " select hbl_pkid as id, hbl_no as code, hbl_prefix as name, hbl_year as year, 0 as amt from hblm  where ";
                        sql += " rec_branch_code = '" + branch_code + "'";
                        sql += " and hbl_year = " + year_code;
                        sql += " and hbl_type ='JOB-GN' ";
                        sql += " and hbl_no in (" + snos + ")";
                        sql += " order by hbl_no ";
                    }


                    if (stype == "CNTR SEA EXPORT")
                    {
                        sql = " select cntr_pkid as id,cntr_no as code, cntr_no as name ,cntr_year as year, 0 as amt  from containerm ";
                        sql += " where rec_branch_code = '" + branch_code + "'";
                        sql += " and cntr_year = " + year_code;
                        sql += " and cntr_no in (" + snos + ")";
                        sql += " order by cntr_no ";
                    }

                    if (stype == "EMPLOYEE")
                    {
                        sql = " select cc_pkid as id , cc_code as code , cc_name as name, 0 as year, 0 as amt  ";
                        sql += " from costcenterm where rec_company_code = '" + comp_code + "' and cc_type = 'EMPLOYEE' ";
                        sql += " and cc_code in (" + snos + ")";
                        sql += " order by cc_code ";
                    }

                    dt_job = new DataTable();
                    dt_job = Con_Oracle.ExecuteQuery(sql);

                    foreach (string str in srow)
                    {
                        if (str.Length > 0)
                        {
                            sdata = str.Split(schar);
                            if (sdata.Length > 0)
                            {
                                if (snos.Length > 0)
                                {
                                    foreach (DataRow Dr in dt_job.Rows)
                                    {
                                        if (stype == "JOB SEA EXPORT" || stype == "JOB AIR EXPORT" || stype == "SI SEA EXPORT" || stype == "SI AIR EXPORT" || stype == "SI SEA IMPORT" || stype == "SI AIR IMPORT" || stype == "GENERAL JOB")
                                        {
                                            if (Lib.Conv2Integer(sdata[0].ToString()) == Lib.Conv2Integer(Dr["code"].ToString()))
                                            {
                                                nAmt = Lib.Conv2Decimal(Dr["amt"].ToString());
                                                Dr["amt"] = nAmt + Lib.Conv2Decimal(sdata[1].ToString());
                                                nTotal += Lib.Conv2Decimal(sdata[1].ToString());
                                                break;
                                            }
                                        }
                                        if (stype == "CNTR SEA EXPORT" || stype == "EMPLOYEE")
                                        {
                                            if (sdata[0].ToString().Trim() == Dr["code"].ToString())
                                            {
                                                nAmt = Lib.Conv2Decimal(Dr["amt"].ToString());
                                                Dr["amt"] = nAmt + Lib.Conv2Decimal(sdata[1].ToString());
                                                nTotal += Lib.Conv2Decimal(sdata[1].ToString());
                                                break;
                                            }
                                        }


                                    }
                                }
                            }
                        }
                    }





                    foreach (DataRow Dr in dt_job.Rows)
                    {
                        _Rec = new LovTable();
                        _Rec.id = Dr["id"].ToString();
                        _Rec.type = stype;
                        _Rec.code = Dr["code"].ToString();
                        if (isoldyear == "Y")
                            _Rec.code = Dr["code"].ToString() + "/" + year_code;
                       _Rec.name = Dr["name"].ToString();
                        _Rec.rate = Lib.Convert2Decimal(Dr["amt"].ToString());
                        mList.Add(_Rec);
                    }
                    dt_job.Rows.Clear();
                    RetData.Add("total", nTotal.ToString());
                    RetData.Add(Table.ToLower(), mList);

                }
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


        

        // This is used by AutoComplete Search Box
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            int rows_starting_number = 1;
            int rows_ending_number = 10;
            int rows_total = 0;

            if (SearchData.ContainsKey("rows_starting_number"))
                rows_starting_number = Lib.Conv2Integer(SearchData["rows_starting_number"].ToString());
            if (SearchData.ContainsKey("rows_ending_number"))
                rows_ending_number = Lib.Conv2Integer(SearchData["rows_ending_number"].ToString());

            string Type = "";
            if (SearchData.ContainsKey("type"))
                Type = SearchData["type"].ToString();

            string SubType = "";
            if (SearchData.ContainsKey("subtype"))
                SubType = SearchData["subtype"].ToString();

            string SearchString = "";
            if (SearchData.ContainsKey("searchstring"))
                SearchString = SearchData["searchstring"].ToString();

            string ParentId = "";
            if (SearchData.ContainsKey("parentid"))
                ParentId = SearchData["parentid"].ToString();

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            string branch_code = "";
            if (SearchData.ContainsKey("branch_code"))
                branch_code = SearchData["branch_code"].ToString();


            string where = "";
            if (SearchData.ContainsKey("where"))
                where = SearchData["where"].ToString();

            Boolean isAdmin = false;
            if (SearchData.ContainsKey("user_admin"))
                isAdmin = (Boolean) SearchData["user_admin"];

            Boolean parentmandatory = false;
            if (SearchData.ContainsKey("parentmandatory"))
                parentmandatory = (Boolean)SearchData["parentmandatory"];

            if (SearchString == null)
                SearchString = "";

            if (ParentId == null)
                ParentId = "";

            List<LovTable> mList = new List<LovTable>();
            LovTable mRow;

            DataTable Dt_List = new DataTable();

            Con_Oracle = new DBConnection();

            try
            {
                if (Type == "COMPANY")
                {
                    sql = " select * from ( ";
                    sql += " select  comp_id as id, comp_code as code,  comp_name as name ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by comp_name) rn ";
                    sql += " from companym a";
                    sql += " where  comp_type = 'C'";
                    sql += " and (comp_name like '%" + SearchString.ToUpper() + "%')";
                    sql += " ) a  where rn between {START} and {END} order by name ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());
                }
                else if (Type == "BRANCH")
                {
                    sql = " select * from ( ";
                    sql += " select  comp_pkid as id , comp_code as code, comp_name as name ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by comp_name) rn ";
                    sql += " from companym a";
                    sql += " where comp_type = 'B' and rec_company_code ='" + comp_code + "' ";
                    sql += " and ( ";
                    sql += " comp_code like '%" + SearchString.ToUpper() + "%'";
                    sql += " or comp_name like '" + SearchString.ToUpper() + "%'";
                    sql += ")";
                    //sql += " order by a.comp_name ";
                    sql += " ) a  where rn between {START} and {END} order by name ";

                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());

                }
                else if (Type == "STORE")
                {
                    sql = " select * from ( ";
                    sql += " select  comp_pkid as id , comp_code as code, comp_name as name ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by comp_name) rn ";
                    sql += " from companym a";
                    if ( !isAdmin && parentmandatory)
                        sql += " inner join userd b on a.comp_pkid =  b.user_branch_id and b.user_id = '" + ParentId + "'";
                    sql += " where a.rec_company_code ='" + comp_code + "' and comp_type = 'S' ";
                    sql += " and ( ";
                    sql += " comp_code like '%" + SearchString.ToUpper() + "%'";
                    sql += " or comp_name like '" + SearchString.ToUpper() + "%'";
                    sql += ")";
                    sql += " ) a  where rn between {START} and {END} order by name ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());
                }
                else if (Type == "VENDOR")
                {
                    sql = " select * from ( ";
                    sql += " select  comp_pkid as id , comp_code as code, comp_name as name ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by comp_name) rn ";
                    sql += " from companym a";
                    sql += " where a.rec_company_code ='" + comp_code + "' and comp_type = 'V' ";
                    if (where != "")
                        sql += " and " + where;
                    sql += " and ( ";
                    sql += " comp_code like '%" + SearchString.ToUpper() + "%'";
                    sql += " or comp_name like '" + SearchString.ToUpper() + "%'";
                    sql += ")";
                    sql += " ) a  where rn between {START} and {END} order by name ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());
                }
                else if (Type == "USER")
                {
                    sql = " select * from ( ";
                    sql += " select user_pkid as id ,user_code as code, user_name as name ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by user_name) rn ";
                    sql += " from userm a  where user_code <> 'ADMIN' ";
                    sql += " and a.rec_company_code = '" + comp_code + "'";
                    sql += " and user_name like '%" + SearchString.ToUpper() + "%'";
                    //sql += " order by a.user_name ";
                    sql += " ) a  where rn between {START} and {END} order by name ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());
                }
                else if (Type == "PIM-GROUP")
                {
                    sql = " select * from ( ";
                    sql += " select grp_pkid as id ,grp_level_name as code, grp_level_name as name ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by grp_level_name) rn ";
                    sql += " from pim_groupm a  ";
                    sql += " where a.rec_company_code = '" + comp_code + "'";
                    sql += " and a.grp_table_name  = '" + SubType + "'";
                    sql += " and grp_level_name like '%" + SearchString.ToUpper() + "%'";
                    sql += " ) a  where rn between {START} and {END} order by name ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());
                }
                else if (Type == "TABLECOLUMNS")
                {
                    sql += " select * from ( ";
                    sql += " select a.*, count(*) over() rowscount,row_number() over(order by code) rn from ( ";
                    sql += " select 'STORE_NAME' as id, 'STORE_NAME' as code, 'STORE_NAME' as name ";
                    sql += " union all ";
                    sql += " select 'STORE_APPROVER' as id, 'STORE_APPROVER' as code, 'STORE_APPROVER' as name ";
                    sql += " union all ";
                    sql += " select 'STORE_RECEIVER' as id, 'STORE_RECEIVER' as code, 'STORE_RECEIVER' as name ";

                    sql += " union all ";
                    sql += " select 'PRODUCT' as id, 'PRODUCT' as code, 'PRODUCT' as name ";

                    sql += " union all ";
                    sql += " select 'LOGO_DEFAULT' as id, 'LOGO_DEFAULT' as code, 'LOGO_DEFAULT' as name ";
                    sql += " union all ";
                    sql += " select tabd_col_name as id, tabd_col_name as code, tabd_col_name as name ";
                    sql += " from tablesm a inner ";
                    sql += " join tablesd b on a.tab_pkid = b.tabd_parent_id ";
                    sql += " where a.rec_company_code = '" + comp_code + "' and a.tab_pkid = '" + SubType + "'";
                    sql += " ) a  where code like  '%" + SearchString.ToUpper() + "%'";
                    sql += " ) a  where rn between {START} and {END} order by name ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());

                }
                else if (Type == "TABLESM")
                {
                    sql = " select * from ( ";
                    sql += " select tab_pkid as id ,tab_name as code, tab_name as name ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by tab_name) rn ";
                    sql += " from tablesm a  ";
                    sql += " where a.rec_company_code = '" + comp_code + "'";
                    sql += " and tab_name like '%" + SearchString.ToUpper() + "%'";
                    sql += " ) a  where rn between {START} and {END} order by code ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());
                }

                else if (Type == "COUNTRY" || Type == "STATE" || Type == "GENERAL JOB TYPES" || Type == "UNIT" || Type == "SHIPMENT TYPE" || Type == "INCOTERM" || Type == "CONTAINER TYPE" || Type == "SEA PORT" || Type == "AIR PORT" || Type == "SEA CARRIER" || Type == "AIR CARRIER" || Type == "VESSEL" || Type == "CONTACT TYPE" || Type == "SALESMAN" || Type == "LOCATION" || Type == "PRE CARRIAGE" || Type == "COMMODITY" || Type == "CITY" || Type == "SCHEME CODE" || Type == "END USE" || Type == "SAC" || Type == "CHALIC" || Type == "PAN" || Type == "TAN" || Type == "PORT" || Type == "COURIER COMPANY" || Type == "ESANCHITDOC" || Type == "ACCOUNTS MAIN CODE" || Type == "SERVICE CONTRACT" || Type == "TABLES")
                {
                    sql = " select * from (";

                    if (Type == "SCHEME CODE" || Type == "END USE" || Type == "SAC" || Type == "PAN" || Type == "TAN")
                        sql += "select param_pkid as id , param_code as code, param_code ||'-' || param_name  as name   ";
                    else
                        sql += "select param_pkid as id , param_code as code, param_name  as name  ";

                    sql += " ,count(*) over() rowscount,row_number() over(order by param_name) rn ";
                    sql += " from param  a  ";
                    sql += " where a.rec_company_code = '" + comp_code + "'";

                    if (Type == "PORT")
                        sql += " and param_type in ('SEA PORT','AIR PORT')";
                    else
                        sql += " and param_type ='" + Type + "'";
                    sql += " and (";
                    sql += " param_code like '%" + SearchString.ToUpper() + "%'";
                    sql += " or param_name like '%" + SearchString.ToUpper() + "%'";
                    sql += " )";

                    if (where != "")
                    {
                        sql += " and (" + where + ")";
                    }

                    //sql += " order by a.param_name ";
                    sql += " ) a  where rn between {START} and {END} order by name ";

                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());

                }

                else if (Type == "CURRENCY")
                {
                    sql = " select * from ( ";
                    sql += "select param_pkid as id , param_code as code, param_name  as name, param_rate as rate ,param_id1 as col1 ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by param_name) rn ";
                    sql += " from param  a  ";
                    sql += " where a.rec_company_code = '" + comp_code + "'";
                    if (Type == "CURRENCY")
                        sql += " and param_type ='CURRENCY'";
                    sql += " and ( ";
                    sql += " param_code like '%" + SearchString.ToUpper() + "%'";
                    sql += " or param_name like '%" + SearchString.ToUpper() + "%'";
                    sql += " )";
                    //sql += " order by a.param_name ";
                    sql += " ) a  where rn between {START} and {END} order by name ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());
                }

                else if (Type == "PARAM")
                {
                    sql = " select * from ( ";
                    sql += "select param_pkid as id , param_code as code, param_name  as name, param_rate as rate ,param_id1 as col1 ";
                    sql += " ,count(*) over() rowscount,row_number() over(order by param_name) rn ";
                    sql += " from param  a  ";
                    sql += " where a.rec_company_code = '" + comp_code + "'";
                    sql += " and param_type ='" + SubType + "'";
                    if (where != "")
                        sql += " and " + where;
                    sql += " and ( ";
                    sql += " param_code like '%" + SearchString.ToUpper() + "%'";
                    sql += " or param_name like '%" + SearchString.ToUpper() + "%'";
                    sql += " )";
                    //sql += " order by a.param_name ";
                    sql += " ) a  where rn between {START} and {END} order by name ";
                    sql = sql.Replace("{START}", rows_starting_number.ToString());
                    sql = sql.Replace("{END}", rows_ending_number.ToString());
                }


                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                if (Dt_List.Rows.Count > 0)
                    rows_total = Lib.Conv2Integer(Dt_List.Rows[0]["rowscount"].ToString());

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new LovTable();

                    mRow.id = Dr["id"].ToString();
                    mRow.code = Dr["code"].ToString();
                    mRow.name = Dr["name"].ToString();
                    mRow.rate = 0;
                    mRow.col1 = "";
                    mRow.col2 = "";
                    mRow.col3 = "";
                    mRow.col4 = "";
                    mRow.col5 = "";
                    mRow.col6 = "";
                    mRow.col7 = "";

                    if (Type == "CURRENCY")
                    {
                        mRow.rate = Lib.Conv2Decimal(Dr["rate"].ToString()); //Fwd Rate
                        mRow.col1 = Dr["col1"].ToString();//Clr Rate
                    }
                    if (Type == "STRREFUNDM")
                    {
                        mRow.rate = Lib.Conv2Decimal(Dr["rate"].ToString());
                    }
                    if (Type == "ACCTM")
                    {
                        mRow.col1 = Dr["col1"].ToString();
                        mRow.col2 = Dr["col2"].ToString();
                        mRow.col3 = Dr["col3"].ToString();
                        mRow.col4 = Dr["col4"].ToString();
                        mRow.col5 = Dr["col5"].ToString();
                        mRow.col6 = Dr["col6"].ToString();
                    }
                    if (Type == "CUSTOMERADDRESS")
                    {
                        mRow.col1 = Dr["add_gstin"].ToString();
                        mRow.col2 = Dr["add_state_id"].ToString();
                        mRow.col3 = Dr["state_code"].ToString();
                        mRow.col4 = Dr["state_code"].ToString() + "-" + Dr["state_name"].ToString();
                        mRow.col5 = "N";
                        if (Dr["add_sepz_unit"].ToString() == "Y")
                            mRow.col5 = "Y";

                        mRow.col6 = Dr["add_email"].ToString();

                        mRow.col7 = "N";
                        if (Dr["add_is_export"].ToString() == "Y")
                            mRow.col7 = "Y";

                    }
                    if (Type == "EMPLOYEE")
                    {
                        mRow.col1 = Dr["col1"].ToString();//emp status
                    }
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
            RetData.Add("rows_total", rows_total);

            return RetData;
        }

     

        // Global Settings 
        public string getParamValue(string comp_code, string id, string columnName)
        {
            try
            {

                sql = "select "+ columnName +" from param where rec_company_code = '"+ comp_code +"' and  param_pkid = '" + id + "'";


                Con_Oracle = new DBConnection();
                DataTable Dt_test = new DataTable();

                Dt_test = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                if (Dt_test.Rows.Count > 0)
                    return Dt_test.Rows[0][columnName].ToString();
                else
                    return "";
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

   
    }
}


