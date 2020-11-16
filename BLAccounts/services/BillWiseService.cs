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
    public class BillWiseService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        DataTable Dt_List2 = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<BillWise> mList = new List<BillWise>();
        BillWise mrow;
        int iRow = 0;
        int iCol = 0;
        string type = "";

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
        string searchexpid = "";
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string code = "";
        string id = "";
        string acc_name = "";
        string jobno = "";
        string sbno = "";
        string sbdt = "";
        string invno = "";
        string cominvno = "";
        string ACC_ID = "";
        string volume = "";
        string netwt = "", grwt = "";
        string ErrorMessage = "";
        Boolean main_code = false;
        Boolean all = false;
        decimal tot_tot_amt = 0;
        decimal tot_cgst_amt = 0;
        decimal tot_sgst_amt = 0;
        decimal tot_igst_amt = 0;
        decimal tot_gst_amt = 0;
        decimal tot_net_amt = 0;
        decimal cntr_20 = 0, cntr_40 = 0;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<BillWise>();
            ErrorMessage = "";
            try
            {
                type = SearchData["type"].ToString();

                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();


                type_date = SearchData["type_date"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();

                ACC_ID = SearchData["acc_id"].ToString();

                all = (Boolean)SearchData["all"];

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);

                if (from_date == "NULL" || to_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");

                //id = SearchData["acc_id"].ToString();
                //acc_name = SearchData["acc_name"].ToString();

                //if (type == "SCREEN" && from_date != "NULL" && to_date != "NULL")
                //{
                //    DateTime dt_frm = DateTime.Parse(from_date);
                //    DateTime dt_to = DateTime.Parse(to_date);
                //    int days = (dt_to - dt_frm).Days;

                //    if (days > 31)
                //        Lib.AddError(ref ErrorMessage, " | Only one month data range can be used,use excel to download");
                //}

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (type_date == "INVOICE" || type_date == "PURCHASE")
                {
                    sql = "";

                    sql = " select ";
                    sql += " a.rec_branch_code,jvh_date,jvh_vrno, ";
                    sql += " jvh_type,jvh_sez,c.acc_name as acc_name,jvh_gstin,";
                    sql += " jvh_rc,jvh_gst_type,jvh_cc_category, hbl_no,";
                    sql += " jvh_tot_amt,jvh_cgst_amt, jvh_sgst_amt,jvh_igst_amt,jvh_gst_amt, jvh_net_amt";
                    sql += " from  ledgerh a ";
                    sql += " left  join hblm on jvh_cc_id = hbl_pkid ";
                    sql += " left join acctm c on a.jvh_acc_id = c.acc_pkid";
                    sql += " where a.rec_company_code = '{COMPCODE}' ";
                    if(!all)
                    {
                        sql += " and a.rec_branch_code=  '{BRCODE}' ";
                    }
                    
                    sql += " and a.jvh_year ='{YEAR}'";
                    sql += " and jvh_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";

                    if (type_date == "INVOICE")
                        sql += " and  jvh_type = 'IN'";
                    else
                        sql += " and  jvh_type = 'PN'";

                    if (ACC_ID != "")
                        sql += " and a.jvh_acc_id = '{PKID}'";

                    sql += " order by a.rec_branch_code,a.jvh_date, a.jvh_type, a.jvh_vrno";


                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{PKID}", ACC_ID);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    tot_tot_amt = 0;
                    tot_cgst_amt = 0;
                    tot_sgst_amt = 0;
                    tot_igst_amt = 0;
                    tot_gst_amt = 0;
                    tot_net_amt = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new BillWise();
                        mrow.rowtype = "DETAIL";
                        mrow.rowcolor = "BLACK";
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.jvh_vrno = Dr["jvh_vrno"].ToString();
                        mrow.jvh_type = Dr["jvh_type"].ToString();
                        mrow.jvh_sez = Dr["jvh_sez"].ToString();
                        mrow.acc_name = Dr["acc_name"].ToString();
                        mrow.jvh_gstin = Dr["jvh_gstin"].ToString();
                        mrow.jvh_rc = Dr["jvh_rc"].ToString();
                        mrow.jvh_gst_type = Dr["jvh_gst_type"].ToString();
                        mrow.jvh_cc_category = Dr["jvh_cc_category"].ToString();
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.jvh_tot_amt = Lib.Conv2Decimal(Dr["jvh_tot_amt"].ToString());
                        mrow.jvh_cgst_amt = Lib.Conv2Decimal(Dr["jvh_cgst_amt"].ToString());
                        mrow.jvh_sgst_amt = Lib.Conv2Decimal(Dr["jvh_sgst_amt"].ToString());
                        mrow.jvh_igst_amt = Lib.Conv2Decimal(Dr["jvh_igst_amt"].ToString());
                        mrow.jvh_gst_amt = Lib.Conv2Decimal(Dr["jvh_gst_amt"].ToString());
                        mrow.jvh_net_amt = Lib.Conv2Decimal(Dr["jvh_net_amt"].ToString());
                        mrow.branch = Dr["rec_branch_code"].ToString();

                        mList.Add(mrow);

                        tot_tot_amt += Lib.Conv2Decimal(mrow.jvh_tot_amt.ToString());
                        tot_cgst_amt += Lib.Conv2Decimal(mrow.jvh_cgst_amt.ToString());
                        tot_sgst_amt += Lib.Conv2Decimal(mrow.jvh_sgst_amt.ToString());
                        tot_igst_amt += Lib.Conv2Decimal(mrow.jvh_igst_amt.ToString());
                        tot_gst_amt += Lib.Conv2Decimal(mrow.jvh_gst_amt.ToString());
                        tot_net_amt += Lib.Conv2Decimal(mrow.jvh_net_amt.ToString());
                    }
                    if (mList.Count > 1)
                    {
                        mrow = new BillWise();
                        mrow.rowtype = "TOTAL";
                        mrow.rowcolor = "RED";
                        mrow.jvh_date = "TOTAL";
                        mrow.jvh_tot_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_tot_amt.ToString(), 2));
                        mrow.jvh_cgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_cgst_amt.ToString(), 2));
                        mrow.jvh_sgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_sgst_amt.ToString(), 2));
                        mrow.jvh_igst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_igst_amt.ToString(), 2));
                        mrow.jvh_gst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_gst_amt.ToString(), 2));
                        mrow.jvh_net_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_net_amt.ToString(), 2));
                        mList.Add(mrow);


                    }


                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintBillWiseReport();
                    }
                    Dt_List.Rows.Clear();
                }

                if (type_date == "SEAEXPORT-TAXREPORT")
                {


                    sql = " select 1 as roworder, '' as BR_NAME,a.rec_branch_code, hbl_pkid,hbl_no, jvh_docno,jvh_date,shpr.acc_name as shipper_name, cons.cust_name as consignee_name,";
                    sql += " HBL_NTWT as NTWT,HBL_GRWT as GRWT,";
                    sql += "  pod.param_name as pod_name, hbl_book_cntr_20 as cntr_20,hbl_book_cntr_40 as cntr_40, jv_frt, jv_frt_gst,jv_thc,jv_thc_gst, ";
                    sql += " jv_detn,jv_detn_gst, jv_others,jv_others_gst, ";
                    sql += " nvl(jv_frt_gst,0)+nvl(jv_thc_gst,0) +nvl(jv_detn_gst,0) +nvl(jv_others_gst,0) as total_gst   from (   ";
                    sql += " select a.rec_branch_code,jvh_vrno, jvh_date,jvh_docno,jvh_acc_id,jvh_cc_id,  ";
                    sql += " sum(case when d.acc_code in('1105001') then jv_credit  else 0 end) as jv_frt ,  ";
                    sql += " sum(case when d.acc_code in('1105001') then jv_gst_amt  else 0 end) as jv_frt_gst,  ";

                    sql += " sum(case when d.acc_code in('1106040') then jv_credit  else 0 end) as jv_thc,  ";
                    sql += " sum(case when d.acc_code in('1106040') then  jv_gst_amt  else 0 end) as jv_thc_gst,  ";

                    sql += " sum(case when d.acc_code in('1106017')  then jv_credit  else 0 end) as jv_detn,  ";
                    sql += " sum(case when d.acc_code in('1106017')  then jv_gst_amt  else 0 end) as jv_detn_gst,   ";

                    sql += " sum(case when d.acc_main_code not in('5501','5505')  and d.acc_code not in('1105001','1106040','1106017') then jv_credit  else 0 end) as jv_others,  ";
                    sql += " sum(case when d.acc_main_code not in('5501','5505')  and d.acc_code not in('1105001','1106040','1106017') then jv_gst_amt  else 0 end) as jv_others_gst   ";

                    sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id  ";

                    sql += " inner join acctm c on a.jvh_acc_id = c.acc_pkid  ";
                    sql += " inner join acctm d on b.jv_acc_id = d.acc_pkid ";

                    sql += " inner join hblm on a.jvh_cc_id = hbl_pkid  ";
                    sql += " where a.rec_company_code = '{COMPCODE}' ";
                    if(!all)
                    {
                        sql += " and a.rec_branch_code = '{BRCODE}' ";
                    }
                    
                    sql += " and hbl_type ='HBL-SE'  ";
                    sql += " and jvh_date between '{FDATE}' and '{EDATE}' and jv_credit > 0       ";

                    if (ACC_ID != "")
                        sql += " and a.jvh_acc_id = '{PKID}'";

                    sql += " group by a.rec_branch_code,jvh_vrno, jvh_date , jvh_docno,jvh_acc_id,jvh_cc_id, hbl_pkid";
                    sql += "  ";
                    sql += " ) a ";
                    sql += " left join hblm on jvh_cc_id = hbl_pkid ";
                    sql += " left join acctm shpr on jvh_acc_id = shpr.acc_pkid ";
                    sql += " left join customerm cons on hbl_imp_id = cons.cust_pkid ";
                    sql += " left join param pod on hbl_pod_id = pod.param_pkid   ";
                    sql += " order by a.rec_branch_code,jvh_docno, jvh_date";


                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{PKID}", ACC_ID);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);

                    sql = "";
                    sql = " select hbl_pkid, job_pkid, job_docno,jexp_invoice_no,jexp_comm_invoice_no,opr_sbill_no, opr_sbill_date  ";
                    sql += " from  ledgerh a   ";
                    sql += " inner join  hblm on a.jvh_cc_id= hbl_pkid ";
                    sql += " inner join  jobm on hbl_pkid = jobs_hbl_id ";
                    sql += " inner join  jobexpm on job_pkid= jexp_job_id   ";
                    sql += " inner join  joboperationsm on job_pkid= opr_job_id   ";
                    sql += " where a.rec_company_code ='{COMPCODE}' ";
                    if(!all)
                    {
                        sql += " and  a.rec_branch_code = '{BRCODE}' ";
                    }
                    sql += " and jvh_date between '{FDATE}' and '{EDATE}' and hbl_type ='HBL-SE' ";
                    sql += " order by a.rec_branch_code,hbl_pkid,job_pkid";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    // sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);



                    Dt_List2 = new DataTable();
                    Dt_List2 = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    cntr_20 = 0; cntr_40 = 0;
                    volume = "";
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new BillWise();
                        mrow.rowtype = "DETAIL";
                        mrow.rowcolor = "BLACK";
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);

                        mrow.consignee = Dr["consignee_name"].ToString();
                        mrow.pod = Dr["pod_name"].ToString();
                        cntr_20 = Lib.Conv2Decimal(Dr["cntr_20"].ToString());
                        cntr_40 = Lib.Conv2Decimal(Dr["cntr_40"].ToString());

                        if (cntr_20 > 0)
                            volume = cntr_20 + "*20";
                        if (cntr_40 > 0)
                        {
                            if (volume != "")
                                volume += ",";
                            volume += cntr_40 + "*40";
                        }
                        mrow.volume = volume.ToString();
                        mrow.hbl_ntwt = Lib.Conv2Decimal(Dr["NTWT"].ToString());
                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["GRWT"].ToString());
                        mrow.jv_frt = Lib.Conv2Decimal(Dr["jv_frt"].ToString());

                        mrow.jv_frt_gst = Lib.Conv2Decimal(Dr["jv_frt_gst"].ToString());
                        mrow.jv_thc = Lib.Conv2Decimal(Dr["jv_thc"].ToString());

                        mrow.jv_thc_gst = Lib.Conv2Decimal(Dr["jv_thc_gst"].ToString());
                        mrow.jv_detn = Lib.Conv2Decimal(Dr["jv_detn"].ToString());

                        mrow.jv_detn_gst = Lib.Conv2Decimal(Dr["jv_detn_gst"].ToString());
                        mrow.jv_others = Lib.Conv2Decimal(Dr["jv_others"].ToString());

                        mrow.jv_others_gst = Lib.Conv2Decimal(Dr["jv_others_gst"].ToString());
                        mrow.total_gst = Lib.Conv2Decimal(Dr["total_gst"].ToString());

                        jobno = "";
                        invno = "";
                        cominvno = "";
                        sbno = "";
                        sbdt = "";
                        foreach (DataRow Dr1 in Dt_List2.Select("hbl_pkid='" + Dr["hbl_pkid"].ToString() + "'"))
                        {
                            if (jobno != "")
                                jobno += ",";
                            jobno += Dr1["job_docno"].ToString();

                            if (invno != "")
                                invno += ",";
                            invno += Dr1["jexp_invoice_no"].ToString();

                            if (cominvno != "")
                                cominvno += ",";
                            cominvno += Dr1["jexp_comm_invoice_no"].ToString();

                            if (sbno != "")
                                sbno += ",";
                            sbno += Dr1["opr_sbill_no"].ToString();
                            if (!Dr1["opr_sbill_date"].Equals(DBNull.Value))
                            {
                                sbdt = ((DateTime) Dr1["opr_sbill_date"]).ToString("dd-MM-yyyy");
                            }
                        }

                        mrow.job_docno = jobno;
                        mrow.jexp_invoice_no = invno;
                        mrow.jexp_comm_invoice_no = cominvno;


                        mrow.branch = Dr["rec_branch_code"].ToString();

                        mrow.job_sbno = sbno;
                        mrow.job_sbdt = sbdt;

                        mList.Add(mrow);
                        cntr_20 = 0; cntr_40 = 0; volume = "";

                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintSeaExportTaxReport();
                    }
                    Dt_List.Rows.Clear();
                }


                if (type_date == "AIREXPORT-TAXREPORT")
                {

                    sql = " select 1 as roworder, '' as BR_NAME,a.rec_branch_code, hbl_pkid,hbl_no, jvh_docno,jvh_date,shpr.acc_name as shipper_name, cons.cust_name as consignee_name,  ";
                    sql += " pod.param_name as pod_name,  hbl_chwt as volume,HBL_NTWT as NTWT,HBL_GRWT as GRWT, jv_frt,  jv_frt_gst,jv_other, jv_other_gst, ";
                    sql += " nvl(jv_frt_gst,0) + nvl(jv_other_gst,0) as total_gst   from (   ";
                    sql += " select jvh_vrno, jvh_date,jvh_docno,jvh_acc_id,jvh_cc_id,a.rec_branch_code,  ";
                    sql += " sum(case when d.acc_code in('1205001') then jv_credit  else 0 end) as jv_frt ,  ";
                    sql += " sum(case when d.acc_code in('1205001')  then jv_gst_amt  else 0 end) as jv_frt_gst,  ";
                    sql += " sum(case when d.acc_main_code not in ('5501','5505') and  d.acc_code <> '1205001'  then jv_credit  else 0 end) as jv_other,  ";
                    sql += " sum(case when d.acc_main_code not in('5501','5505') and  d.acc_code <> '1205001' then jv_gst_amt  else 0 end) as jv_other_gst  ";
                    sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id  ";
                    sql += " inner join acctm c on a.jvh_acc_id = c.acc_pkid  ";

                    sql += " inner join acctm d on b.jv_acc_id = d.acc_pkid ";

                    sql += " left join hblm on a.jvh_cc_id= hbl_pkid  ";
                    sql += " where a.rec_company_code = '{COMPCODE}' ";
                    if(!all)
                    {
                        sql += " and a.rec_branch_code = '{BRCODE}' ";
                    }
                   
                    sql += " and hbl_type ='HBL-AE'";
                    sql += " and jvh_date between '{FDATE}' and '{EDATE}' and jv_credit > 0       ";

                    if (ACC_ID != "")
                        sql += " and a.jvh_acc_id = '{PKID}'";


                    sql += " group by a.rec_branch_code,jvh_vrno, jvh_date , jvh_docno,jvh_acc_id,jvh_cc_id, hbl_pkid";
                    sql += " ) a ";
                    sql += " left join hblm on jvh_cc_id = hbl_pkid ";
                    sql += " left join acctm shpr on jvh_acc_id = shpr.acc_pkid ";
                    sql += " left join customerm cons on hbl_imp_id = cons.cust_pkid ";
                    sql += " left join param pod on hbl_pod_id = pod.param_pkid   ";

                    sql += " order by a.rec_branch_code,jvh_docno, jvh_date";




                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    // sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{PKID}", ACC_ID);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);

                    sql = "";

                    sql = " select hbl_pkid, job_pkid, job_docno,jexp_invoice_no,jvh_docno,jexp_comm_invoice_no, JOB_NTWT as NTWT,JOB_GRWT as GRWT , job_chwt as chwt,opr_sbill_no, opr_sbill_date ";
                    sql += " from  ledgerh a   ";
                    sql += " inner join  hblm on a.jvh_cc_id= hbl_pkid ";
                    sql += " inner join  jobm on hbl_pkid = jobs_hbl_id ";
                    sql += " inner join  jobexpm on job_pkid= jexp_job_id   ";
                    sql += " inner join  joboperationsm on job_pkid= opr_job_id   ";
                    sql += " where a.rec_company_code ='{COMPCODE}' ";
                    if(!all)
                    {
                        sql += " and  a.rec_branch_code = '{BRCODE}'";
                    }
                    
                    sql += " and jvh_date between '{FDATE}' and '{EDATE}' and hbl_type ='HBL-AE' ";
                    sql += " order by a.rec_branch_code,hbl_pkid,job_pkid ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    // sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);



                    Dt_List2 = new DataTable();
                    Dt_List2 = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    cntr_20 = 0; cntr_40 = 0;
                    volume = "";
                    netwt = ""; grwt = "";
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new BillWise();
                        mrow.rowtype = "DETAIL";
                        mrow.rowcolor = "BLACK";
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.consignee = Dr["consignee_name"].ToString();
                        mrow.pod = Dr["pod_name"].ToString();
                        mrow.hbl_chwt = Lib.Conv2Decimal(Dr["volume"].ToString());
                        mrow.hbl_ntwt = Lib.Conv2Decimal(Dr["NTWT"].ToString());
                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["GRWT"].ToString());
                        mrow.jv_frt = Lib.Conv2Decimal(Dr["jv_frt"].ToString());
                        mrow.jv_frt_gst = Lib.Conv2Decimal(Dr["jv_frt_gst"].ToString());
                        mrow.jv_others = Lib.Conv2Decimal(Dr["jv_other"].ToString());
                        mrow.jv_others_gst = Lib.Conv2Decimal(Dr["jv_other_gst"].ToString());
                        mrow.total_gst = Lib.Conv2Decimal(Dr["total_gst"].ToString());

                        jobno = "";
                        invno = "";
                        cominvno = "";
                        sbno = "";
                        sbdt = "";
                        foreach (DataRow Dr1 in Dt_List2.Select("hbl_pkid='" + Dr["hbl_pkid"].ToString() + "'"))
                        {
                            if (jobno != "")
                                jobno += ",";
                            jobno += Dr1["job_docno"].ToString();

                            if (invno != "")
                                invno += ",";
                            invno += Dr1["jexp_invoice_no"].ToString();

                            if (cominvno != "")
                                cominvno += ",";
                            cominvno += Dr1["jexp_comm_invoice_no"].ToString();

                            if (sbno != "")
                                sbno += ",";
                            sbno += Dr1["opr_sbill_no"].ToString();
                            if (!Dr1["opr_sbill_date"].Equals(DBNull.Value))
                            {
                                sbdt = ((DateTime)Dr1["opr_sbill_date"]).ToString("dd-MM-yyyy");
                            }


                        }
                        mrow.job_docno = jobno;
                        mrow.jexp_invoice_no = invno;
                        mrow.jexp_comm_invoice_no = cominvno;
                        mrow.job_sbno = sbno;
                        mrow.job_sbdt = sbdt;
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mList.Add(mrow);
                        cntr_20 = 0; cntr_40 = 0; volume = "";

                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintAirExportTaxReport();
                    }
                    Dt_List.Rows.Clear();
                }


                if (type_date == "SEAEXPORT-TAXREPORT-2")
                {


                    sql = " select 1 as roworder, '' as BR_NAME,a.rec_branch_code, hbl_pkid,hbl_no, jvh_docno,jvh_date,shpr.acc_name as shipper_name, cons.cust_name as consignee_name,";
                    sql += " HBL_NTWT as NTWT,HBL_GRWT as GRWT,";
                    sql += "  pod.param_name as pod_name, hbl_book_cntr_20 as cntr_20,hbl_book_cntr_40 as cntr_40,  ";
                    sql += " jv_dest_truck, jv_bl_amend, jv_bl_surr, jv_detn, jv_bl_reissue, jv_via_charge, jv_detn2,";
                    sql += " jv_others,jv_total_gst as total_gst   from (   ";
                    sql += " select a.rec_branch_code,jvh_vrno, jvh_date,jvh_docno,jvh_acc_id,jvh_cc_id,  ";

                    sql += " sum(case when d.acc_code in('1105027')  then jv_credit  else 0 end) as jv_dest_truck,  ";
                    sql += " sum(case when d.acc_code in('1106004')  then jv_credit  else 0 end) as jv_bl_amend,  ";
                    sql += " sum(case when d.acc_code in('1106005')  then jv_credit  else 0 end) as jv_bl_surr,  ";
                    sql += " sum(case when d.acc_code in('1106017')  then jv_credit  else 0 end) as jv_detn,  ";
                    sql += " sum(case when d.acc_code in('1106031')  then jv_credit  else 0 end) as jv_bl_reissue,  ";
                    sql += " sum(case when d.acc_code in('1106043')  then jv_credit  else 0 end) as jv_via_charge,  ";
                    sql += " sum(case when d.acc_code in('1306015')  then jv_credit  else 0 end) as jv_detn2,  ";

                    sql += " sum(case when d.acc_code not in('1105027','1106004','1106005', '1106017','1106031','1106043','1306015') then jv_credit  else 0 end) as jv_others,  ";
                    sql += " sum(case when d.acc_main_code not in('5501','5505')  then jv_gst_amt  else 0 end) as jv_total_gst   ";

                    sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id  ";
                    sql += " inner join acctm c on a.jvh_acc_id = c.acc_pkid  ";
                    sql += " inner join acctm d on b.jv_acc_id = d.acc_pkid ";

                    sql += " inner join hblm on a.jvh_cc_id = hbl_pkid  ";
                    sql += " where a.rec_company_code = '{COMPCODE}' ";
                    if (!all)
                    {
                        sql += " and a.rec_branch_code = '{BRCODE}' ";
                    }

                    sql += " and hbl_type ='HBL-SE'  ";
                    sql += " and jvh_date between '{FDATE}' and '{EDATE}' and jv_credit > 0       ";

                    if (ACC_ID != "")
                        sql += " and a.jvh_acc_id = '{PKID}'";

                    sql += " group by a.rec_branch_code,jvh_vrno, jvh_date , jvh_docno,jvh_acc_id,jvh_cc_id, hbl_pkid";
                    sql += "  ";
                    sql += " ) a ";
                    sql += " left join hblm on jvh_cc_id = hbl_pkid ";
                    sql += " left join acctm shpr on jvh_acc_id = shpr.acc_pkid ";
                    sql += " left join customerm cons on hbl_imp_id = cons.cust_pkid ";
                    sql += " left join param pod on hbl_pod_id = pod.param_pkid   ";
                    sql += " order by a.rec_branch_code,jvh_docno, jvh_date";


                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{PKID}", ACC_ID);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);

                    sql = "";
                    sql = " select hbl_pkid, job_pkid, job_docno,jexp_invoice_no,jexp_comm_invoice_no,opr_sbill_no, opr_sbill_date  ";
                    sql += " from  ledgerh a   ";
                    sql += " inner join  hblm on a.jvh_cc_id= hbl_pkid ";
                    sql += " inner join  jobm on hbl_pkid = jobs_hbl_id ";
                    sql += " inner join  jobexpm on job_pkid= jexp_job_id   ";
                    sql += " inner join  joboperationsm on job_pkid= opr_job_id   ";
                    sql += " where a.rec_company_code ='{COMPCODE}' ";
                    if (!all)
                    {
                        sql += " and  a.rec_branch_code = '{BRCODE}' ";
                    }
                    sql += " and jvh_date between '{FDATE}' and '{EDATE}' and hbl_type ='HBL-SE' ";
                    sql += " order by a.rec_branch_code,hbl_pkid,job_pkid";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    // sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);



                    Dt_List2 = new DataTable();
                    Dt_List2 = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    cntr_20 = 0; cntr_40 = 0;
                    volume = "";
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new BillWise();
                        mrow.rowtype = "DETAIL";
                        mrow.rowcolor = "BLACK";
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);

                        mrow.consignee = Dr["consignee_name"].ToString();
                        mrow.pod = Dr["pod_name"].ToString();
                        cntr_20 = Lib.Conv2Decimal(Dr["cntr_20"].ToString());
                        cntr_40 = Lib.Conv2Decimal(Dr["cntr_40"].ToString());

                        if (cntr_20 > 0)
                            volume = cntr_20 + "*20";
                        if (cntr_40 > 0)
                        {
                            if (volume != "")
                                volume += ",";
                            volume += cntr_40 + "*40";
                        }
                        mrow.volume = volume.ToString();
                        mrow.hbl_ntwt = Lib.Conv2Decimal(Dr["NTWT"].ToString());
                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["GRWT"].ToString());

                        mrow.jv_dest_truck = Lib.Conv2Decimal(Dr["jv_dest_truck"].ToString());
                        mrow.jv_bl_amend = Lib.Conv2Decimal(Dr["jv_bl_amend"].ToString());
                        mrow.jv_bl_surr = Lib.Conv2Decimal(Dr["jv_bl_surr"].ToString());
                        mrow.jv_bl_reissue = Lib.Conv2Decimal(Dr["jv_bl_reissue"].ToString());
                        mrow.jv_via_charge = Lib.Conv2Decimal(Dr["jv_via_charge"].ToString());
                        mrow.jv_detn = Lib.Conv2Decimal(Dr["jv_detn"].ToString());
                        mrow.jv_detn2 = Lib.Conv2Decimal(Dr["jv_detn2"].ToString());
                        mrow.jv_others = Lib.Conv2Decimal(Dr["jv_others"].ToString());
                        mrow.total_gst = Lib.Conv2Decimal(Dr["total_gst"].ToString());

                        jobno = "";
                        invno = "";
                        cominvno = "";
                        sbno = "";
                        sbdt = "";
                        foreach (DataRow Dr1 in Dt_List2.Select("hbl_pkid='" + Dr["hbl_pkid"].ToString() + "'"))
                        {
                            if (jobno != "")
                                jobno += ",";
                            jobno += Dr1["job_docno"].ToString();

                            if (invno != "")
                                invno += ",";
                            invno += Dr1["jexp_invoice_no"].ToString();

                            if (cominvno != "")
                                cominvno += ",";
                            cominvno += Dr1["jexp_comm_invoice_no"].ToString();
                            if (sbno != "")
                                sbno += ",";
                            sbno += Dr1["opr_sbill_no"].ToString();
                            if (!Dr1["opr_sbill_date"].Equals(DBNull.Value))
                            {
                                sbdt = ((DateTime)Dr1["opr_sbill_date"]).ToString("dd-MM-yyyy");
                            }
                        }
                        mrow.job_docno = jobno;
                        mrow.jexp_invoice_no = invno;
                        mrow.jexp_comm_invoice_no = cominvno;
                        mrow.job_sbno = sbno;
                        mrow.job_sbdt = sbdt;
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mList.Add(mrow);
                        cntr_20 = 0; cntr_40 = 0; volume = "";
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintSeaExportTaxReport2();
                    }
                    Dt_List.Rows.Clear();
                }


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
            return RetData;
        }

        private void PrintBillWiseReport()
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
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                

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

                File_Display_Name = "BillWiseReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                if(!all)
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 15;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 8;
                    WS.Columns[4].Width = 256 * 8;
                    WS.Columns[5].Width = 256 * 25;
                    WS.Columns[6].Width = 256 * 20;
                    WS.Columns[7].Width = 256 * 8;
                    WS.Columns[8].Width = 256 * 15;
                    WS.Columns[9].Width = 256 * 15;
                    WS.Columns[10].Width = 256 * 15;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 15;
                    WS.Columns[14].Width = 256 * 15;
                    WS.Columns[15].Width = 256 * 15;
                    WS.Columns[16].Width = 256 * 15;

                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 15;
                    WS.Columns[4].Width = 256 * 8;
                    WS.Columns[5].Width = 256 * 8;
                    WS.Columns[6].Width = 256 * 25;
                    WS.Columns[7].Width = 256 * 20;
                    WS.Columns[8].Width = 256 * 8;
                    WS.Columns[9].Width = 256 * 15;
                    WS.Columns[10].Width = 256 * 15;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 15;
                    WS.Columns[14].Width = 256 * 15;
                    WS.Columns[15].Width = 256 * 15;
                    WS.Columns[16].Width = 256 * 15;
                    WS.Columns[17].Width = 256 * 15;

                }
                
                iRow = 0; iCol = 1;
               
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
                Lib.WriteData(WS, iRow, 1, "BILL WISE ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SEZ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ACC-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GSTIN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RC", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NET", _Color, true, "BT", "R", "", _Size, false, 325, "", true);



                foreach (BillWise Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_vrno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_sez, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.acc_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gstin, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_rc, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_cc_category, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_tot_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_cgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_sgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_igst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_net_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);



                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_date, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_tot_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_cgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_sgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_igst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_net_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

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




        private void PrintSeaExportTaxReport()
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
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                

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

                File_Display_Name = "SeaExportTaxReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                if(!all)
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 8;
                    WS.Columns[2].Width = 256 * 11;

                    WS.Columns[3].Width = 256 * 15;
                    WS.Columns[4].Width = 256 * 15;

                    WS.Columns[5].Width = 256 * 23;
                    WS.Columns[6].Width = 256 * 13;
                    WS.Columns[7].Width = 256 * 15;
                    WS.Columns[8].Width = 256 * 11;
                    WS.Columns[9].Width = 256 * 25;
                    WS.Columns[10].Width = 256 * 22;
                    WS.Columns[11].Width = 256 * 9;
                    WS.Columns[12].Width = 256 * 12;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[22].Width = 256 * 12;
                    WS.Columns[23].Width = 256 * 12;
                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 13;
                    WS.Columns[2].Width = 256 * 8;
                    WS.Columns[3].Width = 256 * 11;

                    WS.Columns[4].Width = 256 * 15;
                    WS.Columns[5].Width = 256 * 15;

                    WS.Columns[6].Width = 256 * 23;
                    WS.Columns[7].Width = 256 * 13;
                    WS.Columns[8].Width = 256 * 15;
                    WS.Columns[9].Width = 256 * 11;
                    WS.Columns[10].Width = 256 * 25;
                    WS.Columns[11].Width = 256 * 22;
                    WS.Columns[12].Width = 256 * 9;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[22].Width = 256 * 12;
                    WS.Columns[23].Width = 256 * 12;
                    WS.Columns[24].Width = 256 * 12;
                }
               

                iRow = 0; iCol = 1;

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
                Lib.WriteData(WS, iRow, 1, "SEA-EXPORT TAX REPORT FROM " + from_date + " TO " + to_date, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "SI-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                

                Lib.WriteData(WS, iRow, iCol++, "SB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SB-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "JOBNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "COMM.INVNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTM.INVNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VOLUME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FRT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "THC", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DETN", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OTHERS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL-GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);



                foreach (BillWise Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if(all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_sbno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_sbdt, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jexp_invoice_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jexp_comm_invoice_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.volume, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_frt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_frt_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_thc, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_thc_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_detn, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_detn_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_others, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_others_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.total_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);

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



        private void PrintAirExportTaxReport()
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
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                

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

                File_Display_Name = "AirExportTaxReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                if(!all)
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 8;
                    WS.Columns[2].Width = 256 * 11;

                    WS.Columns[3].Width = 256 * 15;
                    WS.Columns[4].Width = 256 * 15;


                    WS.Columns[5].Width = 256 * 23;
                    WS.Columns[6].Width = 256 * 13;
                    WS.Columns[7].Width = 256 * 15;
                    WS.Columns[8].Width = 256 * 11;
                    WS.Columns[9].Width = 256 * 25;
                    WS.Columns[10].Width = 256 * 22;
                    WS.Columns[11].Width = 256 * 9;
                    WS.Columns[12].Width = 256 * 12;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[22].Width = 256 * 12;
                    WS.Columns[23].Width = 256 * 12;

                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 13;
                    WS.Columns[2].Width = 256 * 8;
                    WS.Columns[3].Width = 256 * 11;

                    WS.Columns[4].Width = 256 * 15;
                    WS.Columns[5].Width = 256 * 15;


                    WS.Columns[6].Width = 256 * 23;
                    WS.Columns[7].Width = 256 * 13;
                    WS.Columns[8].Width = 256 * 15;
                    WS.Columns[9].Width = 256 * 11;
                    WS.Columns[10].Width = 256 * 25;
                    WS.Columns[11].Width = 256 * 22;
                    WS.Columns[12].Width = 256 * 9;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[22].Width = 256 * 12;
                    WS.Columns[23].Width = 256 * 12;
                    WS.Columns[24].Width = 256 * 12;

                }

                iRow = 0; iCol = 1;

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
                Lib.WriteData(WS, iRow, 1, "AIR-EXPORT TAX REPORT FROM " + from_date + " TO " + to_date, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "SI-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "SB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SB-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "JOBNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMM.INVNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTM.INVNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FRT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OTHERS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL-GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);



                foreach (BillWise Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if(all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_sbno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_sbdt, _Color, false, "", "L", "", _Size, false, 325, "", true);


                    Lib.WriteData(WS, iRow, iCol++, Rec.job_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jexp_invoice_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jexp_comm_invoice_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_frt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_frt_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_others, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_others_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.total_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);

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


        private void PrintSeaExportTaxReport2()
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
                if (!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }


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

                File_Display_Name = "SeaExportTaxReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                if (!all)
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 8;
                    WS.Columns[2].Width = 256 * 11;

                    WS.Columns[3].Width = 256 * 15;
                    WS.Columns[4].Width = 256 * 15;

                    WS.Columns[5].Width = 256 * 23;
                    WS.Columns[6].Width = 256 * 13;
                    WS.Columns[7].Width = 256 * 15;
                    WS.Columns[8].Width = 256 * 11;
                    WS.Columns[9].Width = 256 * 25;
                    WS.Columns[10].Width = 256 * 22;
                    WS.Columns[11].Width = 256 * 9;
                    WS.Columns[12].Width = 256 * 12;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[22].Width = 256 * 12;
                    WS.Columns[23].Width = 256 * 12;
                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 13;
                    WS.Columns[2].Width = 256 * 8;
                    WS.Columns[3].Width = 256 * 11;

                    WS.Columns[4].Width = 256 * 15;
                    WS.Columns[5].Width = 256 * 15;

                    WS.Columns[6].Width = 256 * 23;
                    WS.Columns[7].Width = 256 * 13;
                    WS.Columns[8].Width = 256 * 15;
                    WS.Columns[9].Width = 256 * 11;
                    WS.Columns[10].Width = 256 * 25;
                    WS.Columns[11].Width = 256 * 22;
                    WS.Columns[12].Width = 256 * 9;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[22].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[23].Width = 256 * 12;
                }


                iRow = 0; iCol = 1;

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
                Lib.WriteData(WS, iRow, 1, "SEA-EXPORT TAX REPORT FROM " + from_date + " TO " + to_date, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "SI-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SB-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "JOBNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMM.INVNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTM.INVNOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VOLUME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "DEST.TRUCK", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BL.AMEND", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BL.SURR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BL.REISSUE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DETN.1106", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VIA.CHARGE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DETN.1306", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OTHERS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL.GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                foreach (BillWise Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_sbno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_sbdt, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jexp_invoice_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jexp_comm_invoice_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.volume, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_dest_truck, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_bl_amend, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_bl_surr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_bl_reissue, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_detn, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_via_charge, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_detn2, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_others, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.total_gst, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);

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
