using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;


using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLAccounts

{
    public class CostRecoService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<Costreco> mList = new List<Costreco>();
        Costreco mrow;
        int iRow = 0;
        int iCol = 0;
        string type = "";


        int iCodeCount = 0;
        string[] aCodes;
        string sql2 = "";

        string report_folder = "";
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string searchtype = "";
        string searchstring = "";
        string rec_category = "";

        string type_date = "";
        string from_date = "";
        string to_date = "";
        string code = "";
        string ErrorMessage = "";
        Boolean main_code = false;
        decimal tot_debit = 0;
        decimal tot_credit = 0;
        decimal tot_deference = 0;
        decimal jv_balance = 0;


        decimal tot_debit0 = 0;
        decimal tot_credit0 = 0;
        decimal tot_debit1= 0;
        decimal tot_credit1 = 0;
        decimal tot_debit2 = 0;
        decimal tot_credit2 = 0;

        string hide_ho_entries = "N";

        string mID = "";

        Boolean IsClr  =false;
        Boolean IsImp = false;

        Boolean Format2 = false;

        string _code = "";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<Costreco>();
            ErrorMessage = "";
            try
            {

                type = SearchData["type"].ToString();
                rec_category = SearchData["rec_category"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();
                type_date = SearchData["type_date"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();
                _code = SearchData["code"].ToString();
                main_code = (Boolean)SearchData["main_code"];
                Format2 = (Boolean)SearchData["format2"];

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);

                hide_ho_entries = SearchData["hide_ho_entries"].ToString();

                IsClr = false;
                IsImp = false;
                if (_code.StartsWith("1101") || _code.StartsWith("1102") || _code.StartsWith("1103") || _code.StartsWith("1104"))
                    IsClr = true ;
                if (_code.StartsWith("1201") || _code.StartsWith("1202") || _code.StartsWith("1203") || _code.StartsWith("1204"))
                    IsClr = true;


                if (_code.StartsWith("1301") || _code.StartsWith("1302") || _code.StartsWith("1303") || _code.StartsWith("1304"))
                {
                    IsClr = true;
                    IsImp = true;
                }

                if (_code.StartsWith("1401") || _code.StartsWith("1402") || _code.StartsWith("1403") || _code.StartsWith("1404"))
                {
                    IsClr = true;
                    IsImp = true;
                }
                
                code = _code;
                aCodes = code.Split(',');
                iCodeCount = aCodes.Length;
                if (_code.ToString().Contains(","))
                {
                   code = _code.ToString().Replace(  ",", "','");
                }

                if (from_date == "NULL" || to_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");







                /*
                if (type == "SCREEN" && from_date != "NULL" && to_date != "NULL")
                {
                    DateTime dt_frm = DateTime.Parse(from_date);
                    DateTime dt_to = DateTime.Parse(to_date);
                    int days = (dt_to - dt_frm).Days;

                    if (days > 31)
                        Lib.AddError(ref ErrorMessage, " | Only one month data range can be used,use excel to download");
                }
                */

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }



                sql = "";

                // DATE, TYPE, VRNO, MBLSL#, MBLNO, BOOKING#, CATEGORY, DEBIT , CREDIT, NARRATION

                if (IsClr)
                {
                    iCodeCount = 1;

                    sql = "";
                    sql += " select jvh_pkid,jvh_vrno, cc_code as mblslno,null as MBLBKNO, null as mblno,null as mblbookno, jvh_date, jvh_type, cc_type as hbl_type, cc_code, cc_name,jvh_narration, null as mstat, null as hstat, ";
                    sql += " sum(case when jv_drcr = 'DR' then ct_amount else 0 end) as jv_debit, ";
                    sql += " sum(case when jv_drcr = 'CR' then ct_amount else 0 end) as jv_credit ";
                    sql += " from ledgerh a inner ";
                    sql += " join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm on jv_acc_id = acc_pkid ";
                    sql += " left join costcentert on jv_pkid = ct_jv_id and ct_posted = 'Y' ";
                    sql += " left join costcenterm on ct_cost_id = cc_pkid ";
                    sql += " where ";
                    if (main_code)
                        sql += " acc_main_code in( '{CODE}' ) ";
                    else
                        sql += " acc_code in( '{CODE}' ) ";
                    sql += " and a.rec_branch_code = '{BRCODE}' ";
                    sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')  ";
                    if (hide_ho_entries == "Y")
                        sql += "  and jvh_type not in ('HO','IN-ES') ";
                    sql += " group by jvh_pkid, jvh_vrno, jvh_date, jvh_type,cc_type, cc_pkid, cc_code, cc_name, jvh_narration";
                    sql += " order by cc_type, cc_code, cc_name, jvh_date ";

                }
                else
                {
                    sql = " select * from (";

                    /*
                    // 1107 Cntr / Job 
                    if (_code.StartsWith("1107"))
                    {
                        sql += " select mbl.hbl_no as MBLSLNO, mbl.hbl_bl_no as mblno,mbl.hbl_book_no as mblbookno,";
                        sql += " hbl.hbl_type as hbl_type, jvh_date,jvh_type, jvh_vrno, jvh_narration, mbl.hbl_terms as mstat, hbl.hbl_terms  as hstat, ";
                        sql += " sum(jv_debit) as jv_debit, sum(jv_credit) as jv_credit";
                        sql += " from ";
                        sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " inner join acctm on jv_acc_id = acc_pkid";
                        sql += " inner join hblm hbl on a.jvh_cc_id = hbl.hbl_pkid";
                        sql += " inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid ";
                        sql += " where jvh_type in('IN' ) and  jvh_cc_category not in ('GENERAL JOB') ";
                        if (main_code)
                            sql += " and acc_main_code in( '{CODE}' ) ";
                        else
                            sql += " and acc_code in('{CODE}')";
                        sql += " and a.rec_branch_code = '{BRCODE}'";
                        sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";
                        sql += " group by mbl.hbl_no, mbl.hbl_bl_no ,mbl.hbl_book_no,";
                        sql += " hbl.hbl_type , jvh_date,jvh_type, jvh_vrno, jvh_narration, mbl.hbl_terms, hbl.hbl_terms  ";
                    }
                    else
                    { 
                        sql += " select mbl.hbl_no as MBLSLNO, mbl.hbl_bl_no as mblno,mbl.hbl_book_no as mblbookno,";
                        sql += " hbl.hbl_type as hbl_type, jvh_date,jvh_type, jvh_vrno, jvh_narration, mbl.hbl_terms as mstat, hbl.hbl_terms  as hstat, ";
                        sql += " sum(jv_debit) as jv_debit, sum(jv_credit) as jv_credit";
                        sql += " from ";
                        sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " inner join acctm on jv_acc_id = acc_pkid";
                        sql += " inner join hblm hbl on a.jvh_cc_id = hbl.hbl_pkid";
                        sql += " left join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";
                        sql += " where jvh_type in('IN' ) and  jvh_cc_category not in ('GENERAL JOB') ";
                        if (main_code)
                            sql += " and acc_main_code in( '{CODE}' ) ";
                        else
                            sql += " and acc_code in('{CODE}')";
                        sql += " and a.rec_branch_code = '{BRCODE}'";
                        sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";
                        sql += " group by mbl.hbl_no, mbl.hbl_bl_no ,mbl.hbl_book_no,";
                        sql += " hbl.hbl_type , jvh_date,jvh_type, jvh_vrno, jvh_narration, mbl.hbl_terms, hbl.hbl_terms  ";
                    }

                    */

                    sql += " select mbl.hbl_no as MBLSLNO, mbl.hbl_bl_no as mblno,mbl.hbl_book_no as mblbookno,";
                    sql += " hbl.hbl_type as hbl_type, jvh_date,jvh_type, jvh_vrno, jvh_narration, mbl.hbl_terms as mstat, hbl.hbl_terms  as hstat, ";
                    sql += " sum(jv_debit) as jv_debit, sum(jv_credit) as jv_credit,";

                    if (iCodeCount > 1)
                    {
                        for (int i = 0; i < aCodes.Length; i++)
                        {
                            sql += " sum (case when acc_main_code ='" + aCodes[i] + "'" + " then jv_debit  else 0 end) as coldr" + i.ToString() + ", ";
                            sql += " sum (case when acc_main_code ='" + aCodes[i] + "'" + " then jv_credit else 0 end) as colcr" + i.ToString() + ", ";
                        }
                    }

                    sql += " max(inv_source) as inv_source, max(hbl.hbl_job_nos) as hbl_job_nos ";

                    sql += " from ";
                    sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm on jv_acc_id = acc_pkid";
                    sql += " inner join hblm hbl on a.jvh_cc_id = hbl.hbl_pkid";
                    sql += " left join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";

                    sql += " left join jobincome jc  on jv_pkid = inv_pkid";

                    sql += " where jvh_type in('IN' ) and  jvh_cc_category not in ('GENERAL JOB') ";
                    if (main_code)
                        sql += " and acc_main_code in( '{CODE}' ) ";
                    else
                        sql += " and acc_code in('{CODE}')";
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                    sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";
                    /*
                    if (hide_ho_entries == "Y")
                        sql += "  and jvh_type <> 'HO' ";
                    */
                    sql += " group by mbl.hbl_no, mbl.hbl_bl_no ,mbl.hbl_book_no,";
                    sql += " hbl.hbl_type , jvh_date,jvh_type, jvh_vrno, jvh_narration, mbl.hbl_terms, hbl.hbl_terms  ";





                    sql += " union all";

                    sql += " select ";
                    sql += " mbl.hbl_no as MBLBKNO, mbl.hbl_bl_no as mblno,mbl.hbl_book_no as mblbookno,";
                    sql += " mbl.hbl_type, jvh_date,jvh_type, jvh_vrno,  jvh_narration,mbl.hbl_terms as mtat, null  as hstat,";
                    sql += " sum(jv_debit)  as jv_debit , sum(jv_credit) as jv_credit,";

                    if (iCodeCount > 1)
                    {
                        for (int i = 0; i < aCodes.Length; i++)
                        {
                            sql += " sum (case when acc_main_code ='" + aCodes[i] + "'" + " then jv_debit  else 0 end) as coldr" + i.ToString() + ", ";
                            sql += " sum (case when acc_main_code ='" + aCodes[i] + "'" + " then jv_credit else 0 end) as colcr" + i.ToString() + ", ";
                        }
                    }


                    sql += " null as inv_source, null as hbl_job_nos";
                    sql += " from ";
                    sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm on jv_acc_id = acc_pkid";
                    sql += " inner join hblm mbl on a.jvh_cc_id = mbl.hbl_pkid";
                    sql += " left join customerm agnt on hbl_agent_id = agnt.cust_pkid";
                    sql += " where jvh_type in('PN') and  jvh_cc_category not in ('GENERAL JOB') ";
                    /*
                    if (hide_ho_entries == "Y")
                        sql += "  and jvh_type <> 'HO' ";
                    */

                    if (main_code)
                        sql += " and acc_main_code in( '{CODE}' ) ";
                    else
                        sql += " and acc_code in( '{CODE}' )";
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                    sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')  ";
                    sql += " group by ";
                    sql += " mbl.hbl_no , mbl.hbl_bl_no ,mbl.hbl_book_no ,";
                    sql += " mbl.hbl_type, jvh_date,jvh_type, jvh_vrno,  jvh_narration, mbl.hbl_terms";

                    sql += " union all";

                    if (hide_ho_entries == "N")
                    {
                        sql += " select ";
                        sql += " mbl.hbl_no as MBLBKNO, mbl.hbl_bl_no as mblno,mbl.hbl_book_no as mblbookno,";
                        sql += " mbl.hbl_type, jvh_date,jvh_type, jvh_vrno,  jvh_narration,mbl.hbl_terms as mtat, null  as hstat,";
                        sql += " sum(jv_debit) as jv_debit, sum(jv_credit) as jv_credit,";

                        if (iCodeCount > 1)
                        {
                            for (int i = 0; i < aCodes.Length; i++)
                            {
                                sql += " sum (case when acc_main_code ='" + aCodes[i] + "'" + " then jv_debit  else 0 end) as coldr" + i.ToString() + ", ";
                                sql += " sum (case when acc_main_code ='" + aCodes[i] + "'" + " then jv_credit else 0 end) as colcr" + i.ToString() + ", ";
                            }
                        }

                        sql += " null as inv_source, null as hbl_job_nos";
                        sql += " from ";
                        sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " inner join acctm on jv_acc_id = acc_pkid";
                        sql += " left join hblm mbl on a.jvh_cc_id = mbl.hbl_pkid";


                        sql += " where jvh_type in('HO','IN-ES') ";
                        if (main_code)
                            sql += " and  acc_main_code in( '{CODE}' ) ";
                        else
                            sql += " and acc_code in( '{CODE}' )";

                        sql += " and a.rec_branch_code = '{BRCODE}'  ";
                        sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";
                        sql += " group by ";
                        sql += " mbl.hbl_no , mbl.hbl_bl_no ,mbl.hbl_book_no ,mbl.hbl_type, jvh_date,jvh_type, jvh_vrno,  jvh_narration,mbl.hbl_terms";

                        sql += " union all";
                    }
                    // GENERAL JOB AND JV, BR, BP, CP,CR

                    sql += " select";
                    sql += " m.hbl_no as MBLBKNO, m.hbl_bl_no as mblno,m.hbl_book_no as mblbookno,";
                    sql += " h.hbl_type, jvh_date,jvh_type, jvh_vrno,   ";
                    sql += " jvh_narration,m.hbl_terms as mstat, h.hbl_terms  as hstat,";
                    sql += " sum(case when jv_drcr = 'DR' then ct_amount else 0 end ) as jv_debit,";
                    sql += " sum(case when jv_drcr = 'CR' then ct_amount else 0 end ) as jv_credit,";

                    if (iCodeCount > 1)
                    {
                        for (int i = 0; i < aCodes.Length; i++)
                        {
                            sql += " sum (case when jv_drcr = 'DR' and acc_main_code ='" + aCodes[i] + "'" + " then ct_amount  else 0 end) as coldr" + i.ToString() + ", ";
                            sql += " sum (case when jv_drcr = 'CR' and acc_main_code ='" + aCodes[i] + "'" + " then ct_amount else 0 end) as colcr" + i.ToString() + ", ";
                        }
                    }



                    sql += " null as inv_source, null as hbl_job_nos";
                    sql += " from ";
                    sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm on jv_acc_id = acc_pkid";
                    sql += " inner join costcentert on jv_pkid = ct_jv_id and ct_type ='M'";
                    sql += " left join hblm h on ct_cost_id = h.hbl_pkid";
                    sql += " left join hblm m on h.hbl_mbl_id = m.hbl_pkid";
                    sql += " where ( jvh_cc_category ='GENERAL JOB'  or jvh_type in ( 'JV', 'BR', 'BP', 'CP', 'CR','DN','CN','CI','DI'))  ";

                    if (main_code)
                        sql += " and acc_main_code in( '{CODE}' ) ";
                    else
                        sql += " and acc_code in( '{CODE}' )";

                    sql += " and a.rec_branch_code = '{BRCODE}'";
                    sql += " and (ct_category like 'SI%' or ct_category like 'GEN%' )";
                    sql += " and jvh_type not in ('HO', 'IN-ES') ";
                    sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";
                    sql += " group by ";
                    sql += " m.hbl_no ,m.hbl_bl_no ,m.hbl_book_no ,m.hbl_type, h.hbl_type, jvh_date,jvh_type, jvh_vrno,   jvh_narration, m.hbl_terms, h.hbl_terms ";

                    sql += " union all";

                    sql += " select";
                    sql += " m.hbl_no as MBLBKNO, m.hbl_bl_no as mblno,m.hbl_book_no as mblbookno,m.hbl_type, jvh_date,jvh_type, jvh_vrno,  jvh_narration, m.hbl_terms as mstat, null as hstat,";
                    sql += " sum(case when jv_drcr = 'DR' then ct_amount else 0 end) as jv_debit,";
                    sql += " sum(case when jv_drcr = 'CR' then ct_amount else 0 end) as jv_credit,";

                    if (iCodeCount > 1)
                    {
                        for (int i = 0; i < aCodes.Length; i++)
                        {
                            sql += " sum (case when jv_drcr = 'DR' and acc_main_code ='" + aCodes[i] + "'" + " then ct_amount  else 0 end) as coldr" + i.ToString() + ", ";
                            sql += " sum (case when jv_drcr = 'CR' and acc_main_code ='" + aCodes[i] + "'" + " then ct_amount else 0 end) as colcr" + i.ToString() + ", ";
                        }
                    }



                    sql += " null as inv_source, null as hbl_job_nos";
                    sql += " from ";
                    sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm on jv_acc_id = acc_pkid";
                    sql += " inner join costcentert on jv_pkid = ct_jv_id and ct_type ='M'";
                    sql += " inner join containerm on ct_cost_id = cntr_pkid";
                    sql += " left join hblm m on cntr_booking_id = m.hbl_pkid";
                    sql += " where ( jvh_cc_category ='GENERAL JOB'  or jvh_type in ( 'JV', 'BR', 'BP', 'CP', 'CR','DN','CN','CI','DI'))  ";


                    if (main_code)
                        sql += " and acc_main_code in( '{CODE}' )";
                    else
                        sql += " and acc_code in( '{CODE}' )";
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                    sql += " and ct_category like 'CNTR%'";
                    sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";
                    sql += " group by m.hbl_no , m.hbl_bl_no ,m.hbl_book_no ,m.hbl_type, jvh_date,jvh_type, jvh_vrno,  jvh_narration, m.hbl_terms";

                    sql += " union all";

                    sql += " select";
                    sql += " null  as MBLBKNO, cast(job_docno as nvarchar2(25)) as  mblno,null  as mblbookno,";
                    sql += " cast('JOB' as  nvarchar2(10)) as hbl_type, jvh_date,jvh_type, jvh_vrno,";
                    sql += " jvh_narration, null as mstat,  null as hstat, ";
                    sql += " sum(case when jv_drcr = 'DR' then ct_amount else 0 end) as jv_debit,";
                    sql += " sum(case when jv_drcr = 'CR' then ct_amount else 0 end) as jv_credit,";

                    if (iCodeCount > 1)
                    {
                        for (int i = 0; i < aCodes.Length; i++)
                        {
                            sql += " sum (case when jv_drcr = 'DR' and acc_main_code ='" + aCodes[i] + "'" + " then ct_amount  else 0 end) as coldr" + i.ToString() + ", ";
                            sql += " sum (case when jv_drcr = 'CR' and acc_main_code ='" + aCodes[i] + "'" + " then  ct_amount else 0 end) as colcr" + i.ToString() + ", ";
                        }
                    }


                    sql += " null as inv_source, null as hbl_job_nos";
                    sql += " from ";
                    sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm on jv_acc_id = acc_pkid";
                    sql += " inner join costcentert on jv_pkid = ct_jv_id and ct_type ='M'";
                    sql += " inner join jobm on ct_cost_id = job_pkid";
                    //sql += " where ( jvh_cc_category ='GENERAL JOB'  or jvh_type in ('IN', 'JV', 'BR', 'BP', 'CP', 'CR','DN','CN','CI','DI')) ";

                    sql += " where ( jvh_cc_category ='GENERAL JOB'  or jvh_type in ('JV', 'BR', 'BP', 'CP', 'CR','DN','CN','CI','DI')) ";

                    if (main_code)
                        sql += " and acc_main_code in( '{CODE}' ) ";
                    else
                        sql += " and acc_code in( '{CODE}' ) ";
                    sql += " and a.rec_branch_code = '{BRCODE}' and ct_category like 'JOB%'";
                    sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')  ";
                    sql += " group by job_docno,jvh_date,jvh_type, jvh_vrno, jvh_narration  ";


                    sql += " union all";

                    sql += " select";
                    sql += " null  as MBLBKNO, cast('' as nvarchar2(25)) as  mblno,null  as mblbookno,";
                    sql += " cast(jvh_type as  nvarchar2(10)) as hbl_type, jvh_date,jvh_type, jvh_vrno,";
                    sql += " jvh_narration, null as mstat,  null as hstat, ";
                    sql += " sum(case when jv_drcr = 'DR' then jv_debit  else 0 end) as jv_debit,";
                    sql += " sum(case when jv_drcr = 'CR' then jv_credit else 0 end) as jv_credit,";

                    if (iCodeCount > 1)
                    {
                        for (int i = 0; i < aCodes.Length; i++)
                        {
                            sql += " sum (case when jv_drcr = 'DR' and acc_main_code ='" + aCodes[i] + "'" + " then jv_debit  else 0 end) as coldr" + i.ToString() + ", ";
                            sql += " sum (case when jv_drcr = 'CR' and acc_main_code ='" + aCodes[i] + "'" + " then jv_credit else 0 end) as colcr" + i.ToString() + ", ";
                        }
                    }

                    sql += " null as inv_source, null as hbl_job_nos";
                    sql += " from ";
                    sql += " ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm on jv_acc_id = acc_pkid";
                    sql += " left join costcentert on jv_pkid = ct_jv_id and ct_type ='M'";
                    sql += " where (jvh_type in ( 'JV', 'BR', 'BP', 'CP', 'CR','DN','CN','CI','DI')) ";

                    if (main_code)
                        sql += " and acc_main_code in( '{CODE}' ) ";
                    else
                        sql += " and acc_code in( '{CODE}' ) ";
                    sql += " and a.rec_branch_code = '{BRCODE}' and ct_category is null ";
                    sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')  ";
                    sql += " group by jvh_date,jvh_type, jvh_vrno, jvh_narration  ";



                    sql += " ) a ";
                    sql += " order by mblslno, jvh_vrno";

                    if (Format2)
                    {
                        /*
                        sql += " select mbl.hbl_no as MBLSLNO, mbl.hbl_bl_no as mblno,mbl.hbl_book_no as mblbookno,";
                        sql += " hbl.hbl_type as hbl_type, jvh_date,jvh_type, jvh_vrno, jvh_narration,null as mstat, null as hstat,";
                        sql += " sum(jv_debit) as jv_debit, sum(jv_credit) as jv_credit,null as inv_source, null as hbl_job_nos";
                        */

                        sql = "";
                        sql += " select * from( ";
                        sql += " select m.hbl_no as mblslno, m.hbl_bl_no as mblno,m.hbl_book_no as mblbookno, ";
                        sql += " cc_type as hbl_type, jvh_date, jvh_type, jvh_vrno, jvh_narration, ";
                        sql += " h.hbl_no as sino ,h.hbl_bl_no as hblno, cc_code, ";
                        sql += " max(jv_debit) as ho_dr, max(jv_credit) as ho_cr,  ";
                        sql += " sum(case when jv_drcr = 'DR' then abs(ct_amount) else 0 end) as jv_debit, ";
                        sql += " sum(case when jv_drcr = 'CR' then abs(ct_amount) else 0 end) as jv_credit,null as inv_source, null as hbl_job_nos ";
                        sql += " from  ledgerh a  ";
                        sql += " inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " inner join acctm on jv_acc_id = acc_pkid ";
                        sql += " left join costcentert on jv_pkid = ct_jv_id and ct_posted = 'Y' ";
                        sql += " left join costcenterm on ct_cost_id = cc_pkid ";
                        sql += " left join hblm h on ct_cost_id = h.hbl_pkid ";
                        sql += " left join hblm m on h.hbl_mbl_id = m.hbl_pkid ";
                        sql += " left join hblm mcc on jvh_cc_id = mcc.hbl_pkid ";
                        sql += " where ";
                        if (main_code)
                            sql += " acc_main_code in( '{CODE}' ) ";
                        else
                            sql += " acc_code in( '{CODE}' ) ";
                        sql += " and a.rec_branch_code = '{BRCODE}' ";
                        sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')  ";

                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in( 'HO', 'IN-ES') ";


                        sql += " group by  a.rec_branch_code, jvh_pkid, jvh_vrno, jvh_date, jvh_type,cc_type, cc_pkid, cc_code,jvh_narration, ct_cost_id, m.hbl_no, m.hbl_bl_no,m.hbl_book_no,  ";
                        sql += " h.hbl_no,h.hbl_bl_no ";
                        sql += " ) a ";
                        sql += " order by mblno, hblno ";

                    }
                }

                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{CODE}", code);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                tot_credit = 0;
                tot_debit = 0;
                tot_deference = 0;
                jv_balance = 0;

                if ( Dt_List.Rows.Count > 0)
                {
                    mID = Dt_List.Rows[0]["MBLSLNO"].ToString();
                }

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    if ( mID != Dr["MBLSLNO"].ToString())
                    {
                        if ( mrow != null)
                        {
                            mrow.jv_balance = jv_balance;
                        }
                        jv_balance = 0;
                        mID = Dr["MBLSLNO"].ToString();
                    }

                    mrow = new Costreco();
                    mrow.row_type = "DETAIL";
                    mrow.row_colour = "BLACK";
                    mrow.mbl_no = Dr["MBLSLNO"].ToString();
                    mrow.mbl_bl_no = Dr["mblno"].ToString();
                    mrow.mbl_book_no = Dr["mblbookno"].ToString();
                    mrow.hbl_type = Dr["hbl_type"].ToString();
                    mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mrow.jvh_type = Dr["jvh_type"].ToString();
                    mrow.jvh_vrno = Dr["jvh_vrno"].ToString();

                    if ( code == "1107")
                    {
                        if ( Dr["inv_source"].ToString() == "LOCAL CHARGES")
                        {
                            mrow.mbl_no = Dr["hbl_job_nos"].ToString();
                            mrow.hbl_type = "JOB";
                        }
                    }



                    if (Format2)
                    {
                        mrow.mstat = "";
                        mrow.hstat = "";
                    }
                    else
                    {
                        mrow.mstat = Dr["mstat"].ToString().Replace("FREIGHT", "");
                        mrow.hstat = Dr["hstat"].ToString().Replace("FREIGHT", "");
                    }

                    if ( Format2 && !IsClr)
                    {
                        mrow.hbl_type = Dr["hbl_type"].ToString() + "-" + Dr["cc_code"].ToString();
                        if (Dr["jvh_type"].ToString() == "HO")
                        {
                            if (Lib.Conv2Decimal(Dr["jv_debit"].ToString()) == 0 && Lib.Conv2Decimal(Dr["ho_dr"].ToString()) > 0)
                                Dr["jv_debit"] = Lib.Conv2Decimal(Dr["ho_dr"].ToString());
                            if (Lib.Conv2Decimal(Dr["jv_credit"].ToString()) == 0 && Lib.Conv2Decimal(Dr["ho_cr"].ToString()) > 0)
                                Dr["jv_credit"] = Lib.Conv2Decimal(Dr["ho_cr"].ToString());
                        }
                    }

                    mrow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                    mrow.jvh_narration = Dr["jvh_narration"].ToString();
                    jv_balance += Lib.Conv2Decimal(Dr["jv_credit"].ToString()) - Lib.Conv2Decimal(Dr["jv_debit"].ToString());


                    if (iCodeCount > 1)
                    {
                        if (aCodes.Length >= 1)
                        {
                            mrow.coldr0 = Lib.Conv2Decimal(Dr["coldr0"].ToString());
                            mrow.colcr0 = Lib.Conv2Decimal(Dr["colcr0"].ToString());
                            tot_debit0 += Lib.Conv2Decimal(mrow.coldr0.ToString());
                            tot_credit0 += Lib.Conv2Decimal(mrow.colcr0.ToString());
                        }
                        if (aCodes.Length >= 2)
                        {
                            mrow.coldr1 = Lib.Conv2Decimal(Dr["coldr1"].ToString());
                            mrow.colcr1 = Lib.Conv2Decimal(Dr["colcr1"].ToString());
                            tot_debit1 += Lib.Conv2Decimal(mrow.coldr1.ToString());
                            tot_credit1 += Lib.Conv2Decimal(mrow.colcr1.ToString());
                        }
                        if (aCodes.Length >= 3)
                        {
                            mrow.coldr2 = Lib.Conv2Decimal(Dr["coldr2"].ToString());
                            mrow.colcr2 = Lib.Conv2Decimal(Dr["colcr2"].ToString());
                            tot_debit2 += Lib.Conv2Decimal(mrow.coldr2.ToString());
                            tot_credit2 += Lib.Conv2Decimal(mrow.colcr2.ToString());
                        }
                    }





                    mList.Add(mrow);

                    tot_debit += Lib.Conv2Decimal(mrow.jv_debit.ToString());
                    tot_credit += Lib.Conv2Decimal(mrow.jv_credit.ToString());

                }
                if (mList.Count > 1)
                {
                    if (mrow != null)
                    {
                        mrow.jv_balance = jv_balance;
                    }

                    mrow = new Costreco();
                    mrow.row_type = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.mbl_no = "TOTAL";
                    mrow.jv_debit = Lib.Conv2Decimal(Lib.NumericFormat(tot_debit.ToString(), 2));
                    mrow.jv_credit = Lib.Conv2Decimal(Lib.NumericFormat(tot_credit.ToString(), 2));

                    tot_deference = tot_credit - tot_debit;
                    mrow.jv_balance = Lib.Conv2Decimal(Lib.NumericFormat(tot_deference.ToString(), 2));


                    if (iCodeCount > 1)
                    {
                        if (aCodes.Length >= 1)
                        {
                            mrow.coldr0 = Lib.Conv2Decimal(Lib.NumericFormat(tot_debit0.ToString(), 2));
                            mrow.colcr0 = Lib.Conv2Decimal(Lib.NumericFormat(tot_credit0.ToString(), 2));
                        }
                        if (aCodes.Length >= 2)
                        {
                            mrow.coldr1 = Lib.Conv2Decimal(Lib.NumericFormat(tot_debit1.ToString(), 2));
                            mrow.colcr1 = Lib.Conv2Decimal(Lib.NumericFormat(tot_credit1.ToString(), 2));
                        }
                        if (aCodes.Length >= 3)
                        {
                            mrow.coldr1 = Lib.Conv2Decimal(Lib.NumericFormat(tot_debit1.ToString(), 2));
                            mrow.colcr1 = Lib.Conv2Decimal(Lib.NumericFormat(tot_credit1.ToString(), 2));
                        }
                    }


                    mList.Add(mrow);


                    /*
                    mrow = new Costreco();
                    mrow.row_type = "DIFFERENCE";
                    mrow.row_colour = "RED";
                    mrow.mbl_no = "DIFFERENCE";

                    if (tot_debit > tot_credit)
                    {
                        tot_deference = tot_debit - tot_credit;
                        mrow.jv_debit = 0;
                        mrow.jv_credit = Lib.Conv2Decimal(Lib.NumericFormat(tot_deference.ToString(), 2));
                    }
                    else
                    {
                        tot_deference = tot_credit - tot_debit;
                        mrow.jv_credit = 0;
                        mrow.jv_debit = Lib.Conv2Decimal(Lib.NumericFormat(tot_deference.ToString(), 2));
                    }
                    mList.Add(mrow);
                    */

                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintCostRecoReport();
                }
                Dt_List.Rows.Clear();

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("type", type);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", mList);
            RetData.Add("isclr", IsClr);
            RetData.Add("isimp", IsImp);
            RetData.Add("codecount", iCodeCount);
            return RetData;
        }

        private void PrintCostRecoReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                mSearchData.Add("branch_code", branch_code);

                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "CostRecoReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 15;
                WS.Columns[15].Width = 256 * 15;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 15;

                WS.Columns[20].Width = 256 * 15;
                WS.Columns[21].Width = 256 * 15;
                WS.Columns[22].Width = 256 * 15;
                WS.Columns[23].Width = 256 * 15;
                WS.Columns[24].Width = 256 * 15;
                WS.Columns[25].Width = 256 * 15;



                iRow = 0; iCol = 1;
                WS.Columns[8].Style.NumberFormat = "#0.00";
                WS.Columns[9].Style.NumberFormat = "#0.00";
                WS.Columns[10].Style.NumberFormat = "#0.00";

                _Size = 14;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 12;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "COSTING RECONCILE ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;

                if (IsClr)
                {
                    if ( IsImp)
                        Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    else 
                        Lib.WriteData(WS, iRow, iCol++, "JOB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                else
                {
                    Lib.WriteData(WS, iRow, iCol++, "MBLSL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "MBLNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "BOOKING#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }

                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                if (!IsClr)
                {
                    Lib.WriteData(WS, iRow, iCol++, "M-STAT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "H-STAT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }






                if (iCodeCount > 1)
                {
                    if (aCodes.Length >= 1)
                    {
                        Lib.WriteData(WS, iRow, iCol++, aCodes[0]+ "-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, aCodes[0]+ "-CR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                    }
                    if (aCodes.Length >= 2)
                    {
                        Lib.WriteData(WS, iRow, iCol++, aCodes[1] + "-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, aCodes[1] + "-CR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                    }
                    if (aCodes.Length >= 3)
                    {
                        Lib.WriteData(WS, iRow, iCol++, aCodes[2] + "-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, aCodes[2] + "-CR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                    }
                }

                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "NARRATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);



                decimal val = 0;
                foreach (Costreco Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {

                        if (IsClr)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        else
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.mbl_book_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }

                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_vrno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        if (!IsClr)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.mstat, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.hstat, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }


                        if (iCodeCount > 1)
                        {
                            if (aCodes.Length >= 1)
                            {
                                Lib.WriteData(WS, iRow, iCol++, Rec.coldr0, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, Rec.colcr0, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            }
                            if (aCodes.Length >= 2)
                            {
                                Lib.WriteData(WS, iRow, iCol++, Rec.coldr1, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, Rec.colcr1, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                            }
                            if (aCodes.Length >= 3)
                            {
                                Lib.WriteData(WS, iRow, iCol++, Rec.coldr2, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, Rec.colcr2, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            }
                        }

                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_debit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.jv_balance != null ? Lib.Conv2Decimal(Rec.jv_balance.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_narration, _Color, false, "", "L", "", _Size, false, 325, "", true);



                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if (IsClr) {
                            Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, true, "T", "L", "", _Size, false, 325, "", true);
                        }
                        else
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, true, "T", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                       
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);

                        if (!IsClr)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        }


                        if (iCodeCount > 1)
                        {
                            if (aCodes.Length >= 1)
                            {
                                Lib.WriteData(WS, iRow, iCol++, Rec.coldr0, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, Rec.colcr0, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            }
                            if (aCodes.Length >= 2)
                            {
                                Lib.WriteData(WS, iRow, iCol++, Rec.coldr1, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, Rec.colcr1, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                            }
                            if (aCodes.Length >= 3)
                            {
                                Lib.WriteData(WS, iRow, iCol++, Rec.coldr2, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, Rec.colcr2, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            }
                        }

                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_debit, _Color, true, "T", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, true, "T", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.jv_balance != null ? Lib.Conv2Decimal(Rec.jv_balance.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, true, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);




                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);

                    }
                    if (Rec.row_type == "DIFFERENCE")
                    {
                        if (IsClr) {
                            Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, true, "B", "L", "", _Size, false, 325, "", true);
                        }
                        else
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, true, "B", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                       
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);


                        if (iCodeCount > 1)
                        {
                            if (aCodes.Length >= 1)
                            {
                                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            }
                            if (aCodes.Length >= 2)
                            {
                                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                            }
                            if (aCodes.Length >= 3)
                            {
                                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            }
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_debit, _Color, true, "B", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, true, "B", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "B", "L", "", _Size, false, 325, "", true);



                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);

                    }
                }
                // iRow++;

                WB.SaveXls(File_Name);
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
