using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

using XL.XSheet;

namespace BLReport1
{
    public class GstReportService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<GstReport> mList = new List<GstReport>();
        GstReport mrow;
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
        string searchexpid = "";
        string format_type = "";
        string from_date = "";
        string to_date = "";
        string ErrorMessage = "";

        decimal tot_inv_amt = 0;
        decimal tot_taxable_amt = 0;
        decimal tot_cgst_amt = 0;
        decimal tot_sgst_amt = 0;
        decimal tot_igst_amt = 0;
        decimal tot_gst_amt = 0;

        Boolean all = false;

        Dictionary<int, string> DocInvCountDic;
        Dictionary<int, string> DocDNCountDic;
        Dictionary<int, string> DocCNCountDic;
        Dictionary<int, string> DocInvExpCountDic;

        DataTable Dt_B2CL = null;
        DataTable Dt_B2SM = null;
        DataTable Dt_HSN = null;
        DataTable Dt_CDNR = null;
        DataTable Dt_CDNUR = null;
        DataTable Dt_EXP = null;
        string GenerateMsg = "";
       
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<GstReport>();
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
                format_type = SearchData["format_type"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();

                all = (Boolean)SearchData["all"];

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);


                if (from_date == "NULL" || to_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");

                //if ((type == "SCREEN" || type == "GSTR1") && from_date != "NULL" && to_date != "NULL")
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

                if (format_type == "GSTR1")
                {
                    /*
                    sql = " select jvh_docno  ";
                    sql += "   ,h.jvh_cc_category  ";
                    sql += "   ,jvh_date ";
                    sql += "   ,jvh_gstin  ";
                    sql += "   ,max(jvh_net_amt) as jvh_net_amt";
                    sql += "   ,sum(jv_net_total) as jv_net_total"; //inv_amt
                    sql += "   ,sum(jv_credit) as jv_Credit";
                    sql += "   ,sum(jv_taxable_amt) as jv_taxable_amt";
                    sql += "   ,jvh_gst_type  ";
                    sql += "   ,jv_gst_rate  ";
                    sql += "   ,sum(jv_cgst_amt) as jv_cgst_amt";
                    sql += "   ,sum(jv_sgst_amt) as jv_sgst_amt";
                    sql += "   ,sum(jv_igst_amt) as jv_igst_amt";
                    sql += "   ,sum(jv_gst_amt) as jv_gst_amt";
                    sql += "   ,jvh_sez";
                    sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
                    sql += "   ,'N' as RC";
                    sql += "   ,case when jvh_sez='Y' then";
                    sql += "    case when jvh_gst='Y' and sum(jv_gst_amt)>0 then 'SEZ supplies with payment' else 'SEZ supplies without payment' end ";
                    sql += "    else 'Regular' end as jvh_invoice_Type";
                    sql += "   ,'' as ecomgstn";
                    sql += "   ,0 as cess";
                    sql += "   ,jvh_gst";
                    sql += "   ,max(party.acc_name) as jvh_party_name ";
                    sql += "   from ledgerh h";
                    sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += "   left join param st on h.jvh_state_id = st.param_pkid";
                    sql += "   left join acctm party on h.jvh_acc_id = party.acc_pkid";
                    sql += "   where h.rec_company_code = '{COMPCODE}'";
                    if (this.branch_code != "")
                        sql += "   and h.rec_branch_code = '{BRCODE}'";
                    sql += "   and h.jvh_type in ('IN') ";
                    sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(party.acc_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(h.jvh_gstin) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "   group by jvh_docno, h.jvh_cc_category, jvh_date, jvh_gstin,jvh_gst_type, jv_gst_rate,jvh_sez, st.param_code, st.param_name,jvh_gst";
                    sql += "   order by jvh_docno,jv_gst_rate";
                    */

                    sql = " select jvh_docno,h.rec_branch_code  ";
                    sql += "   ,h.jvh_cc_category  ";
                    sql += "   ,jvh_date ";
                    sql += "   ,jvh_gstin  ";
                    //sql += "   ,max(jvh_tot_amt) as inv_amt";//inv_amt
                    sql += "   ,max(jvh_net_amt) as inv_amt";//inv_amt
                    sql += "   ,sum(jv_credit) as taxable_amt"; //taxble amt
                    sql += "   ,jvh_gst_type  ";
                    sql += "   ,0 as jv_gst_rate  ";
                    sql += "   ,sum(jv_cgst_amt) as jv_cgst_amt";
                    sql += "   ,sum(jv_sgst_amt) as jv_sgst_amt";
                    sql += "   ,sum(jv_igst_amt) as jv_igst_amt";
                    sql += "   ,sum(jv_gst_amt) as jv_gst_amt";
                    sql += "   ,jvh_sez";
                    sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
                    sql += "   ,'N' as RC";
                    sql += "   ,case when jvh_sez='Y' then";
                    sql += "    case when jvh_gst='Y' and sum(jv_gst_amt)>0 then 'SEZ supplies with payment' else 'SEZ supplies without payment' end ";
                    sql += "    else 'Regular' end as jvh_invoice_Type";
                    sql += "   ,'' as ecomgstn";
                    sql += "   ,0 as cess";
                    sql += "   ,jvh_gst";
                    sql += "   ,max(party.acc_name) as jvh_party_name ";
                    sql += "   from ledgerh h";
                    sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += "   left join param st on h.jvh_state_id = st.param_pkid";
                    sql += "   left join acctm party on h.jvh_acc_id = party.acc_pkid";
                    sql += "   where h.rec_company_code = '{COMPCODE}'";
                    if (!all)
                        sql += "   and h.rec_branch_code = '{BRCODE}'";
                    sql += "   and h.jvh_type in ('IN') ";
                    sql += "   and h.jvh_gst = 'N' ";
                    sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(party.acc_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(h.jvh_gstin) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "   group by h.rec_branch_code,jvh_docno, h.jvh_cc_category, jvh_date, jvh_gstin,jvh_gst_type,jvh_sez, st.param_code, st.param_name,jvh_gst";

                    sql += " Union all";


                    sql += " select jvh_docno,h.rec_branch_code  ";
                    sql += "   ,h.jvh_cc_category  ";
                    sql += "   ,jvh_date ";
                    sql += "   ,jvh_gstin  ";
                   // sql += "   ,max(jvh_tot_amt) as inv_amt";//inv_amt
                    sql += "   ,max(jvh_net_amt) as inv_amt";//inv_amt
                    sql += "   ,sum(jv_credit) as taxable_amt"; //taxable amt
                    sql += "   ,jvh_gst_type  ";
                    sql += "   ,jv_gst_rate  ";
                    sql += "   ,sum(jv_cgst_amt) as jv_cgst_amt";
                    sql += "   ,sum(jv_sgst_amt) as jv_sgst_amt";
                    sql += "   ,sum(jv_igst_amt) as jv_igst_amt";
                    sql += "   ,sum(jv_gst_amt) as jv_gst_amt";
                    sql += "   ,jvh_sez";
                    sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
                    sql += "   ,'N' as RC";
                    sql += "   ,case when jvh_sez='Y' then";
                    sql += "    case when jvh_gst='Y' and sum(jv_gst_amt)>0 then 'SEZ supplies with payment' else 'SEZ supplies without payment' end ";
                    sql += "    else 'Regular' end as jvh_invoice_Type";
                    sql += "   ,'' as ecomgstn";
                    sql += "   ,0 as cess";
                    sql += "   ,jvh_gst";
                    sql += "   ,max(party.acc_name) as jvh_party_name ";
                    sql += "   from ledgerh h";
                    sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += "   left join param st on h.jvh_state_id = st.param_pkid";
                    sql += "   left join acctm party on h.jvh_acc_id = party.acc_pkid";
                    sql += "   where h.rec_company_code = '{COMPCODE}'";
                    if (!all)
                        sql += "   and h.rec_branch_code = '{BRCODE}'";
                    sql += "   and h.jvh_type in ('IN') ";
                    sql += "   and h.jvh_gst = 'Y' ";
                   // sql += "   and jv_gst_rate > 0 ";
                    sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(party.acc_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(h.jvh_gstin) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "   group by h.rec_branch_code,jvh_docno, h.jvh_cc_category, jvh_date, jvh_gstin,jvh_gst_type,jv_gst_rate ,jvh_sez, st.param_code, st.param_name,jvh_gst";


                    sql += "   order by rec_branch_code,jvh_docno,jv_gst_rate";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    tot_inv_amt = 0;
                    tot_taxable_amt = 0;
                    tot_cgst_amt = 0;
                    tot_sgst_amt = 0;
                    tot_igst_amt = 0;
                    tot_gst_amt = 0;

                    DocInvCountDic = new Dictionary<int, string>();
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        if (!DocInvCountDic.ContainsValue(Dr["jvh_docno"].ToString().Trim()))
                            DocInvCountDic.Add(DocInvCountDic.Count, Dr["jvh_docno"].ToString().Trim());
     
                        mrow = new GstReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_cc_category = Dr["jvh_cc_category"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        if (!Dr["jvh_date"].Equals(DBNull.Value))
                            mrow.jvh_date_gstr1 = ((DateTime)Dr["jvh_date"]).ToString("dd-MMM-yyyy");
                        else
                            mrow.jvh_date_gstr1 = "";
                        mrow.jvh_gstin = Dr["jvh_gstin"].ToString();
                        mrow.inv_amt = Lib.Conv2Decimal(Dr["inv_amt"].ToString());
                        mrow.taxable_amt = Lib.Conv2Decimal(Dr["taxable_amt"].ToString());
                        mrow.jvh_gst_type = Dr["jvh_gst_type"].ToString();
                        mrow.jv_gst_rate = Lib.Conv2Decimal(Dr["jv_gst_rate"].ToString());
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Dr["jv_cgst_amt"].ToString());
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Dr["jv_sgst_amt"].ToString());
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Dr["jv_igst_amt"].ToString());
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Dr["jv_gst_amt"].ToString());
                        mrow.jvh_sez = Dr["jvh_sez"].ToString();
                        mrow.jvh_state_name = Dr["jvh_state_name"].ToString();
                        mrow.rc = Dr["rc"].ToString();
                        mrow.jvh_invoice_type = Dr["jvh_invoice_type"].ToString();
                        mrow.ecomgstn = Dr["ecomgstn"].ToString();
                        mrow.cess = Lib.Conv2Decimal(Dr["cess"].ToString());
                        mrow.jvh_gst = Dr["jvh_gst"].ToString();
                        mrow.jvh_party_name = Dr["jvh_party_name"].ToString();
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mList.Add(mrow);

                        tot_inv_amt += Lib.Conv2Decimal(mrow.inv_amt.ToString());
                        tot_taxable_amt += Lib.Conv2Decimal(mrow.taxable_amt.ToString());
                        tot_cgst_amt += Lib.Conv2Decimal(mrow.jv_cgst_amt.ToString());
                        tot_sgst_amt += Lib.Conv2Decimal(mrow.jv_sgst_amt.ToString());
                        tot_igst_amt += Lib.Conv2Decimal(mrow.jv_igst_amt.ToString());
                        tot_gst_amt += Lib.Conv2Decimal(mrow.jv_gst_amt.ToString());
                    }
                    if (mList.Count > 1)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.jvh_invoice_type = "TOTAL";
                        mrow.inv_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_inv_amt.ToString(), 2));
                        mrow.taxable_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_taxable_amt.ToString(), 2));
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_cgst_amt.ToString(), 2));
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_sgst_amt.ToString(), 2));
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_igst_amt.ToString(), 2));
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_gst_amt.ToString(), 2));
                        mList.Add(mrow);
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintGSTR1Report();
                    }
                    if (type == "GSTR1")
                    {
                        if (mList != null)
                            GenerateGSTR1();

                    }
                }

                if (format_type == "INVOICE")
                {
                    sql = " select jvh_docno,h.rec_branch_code ";
                    sql += "   ,h.jvh_cc_category  ";
                    sql += "   ,jvh_date ";
                    sql += "   ,jvh_gstin  ";
                    sql += "   ,max(jvh_net_amt) as jvh_net_amt";
                    sql += "   ,sum(jv_credit) as jv_Credit";
                    sql += "   ,sum(jv_net_total) as jv_net_total";//inv_amt
                    sql += "   ,sum(jv_credit) as jv_taxable_amt";
                    sql += "   ,jvh_gst_type  ";
                    sql += "   ,sum(jv_cgst_amt) as jv_cgst_amt";
                    sql += "   ,sum(jv_sgst_amt) as jv_sgst_amt";
                    sql += "   ,sum(jv_igst_amt) as jv_igst_amt";
                    sql += "   ,sum(jv_gst_amt) as jv_gst_amt";
                    sql += "   ,jvh_sez";
                    sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
                    sql += "   ,'N' as RC";
                    sql += "   ,case when jvh_sez='Y' then";
                    sql += "    case when jvh_gst='Y' and sum(jv_gst_amt)>0 then 'SEZ supplies with payment' else 'SEZ supplies without payment' end ";
                    sql += "    else 'Regular' end as jvh_invoice_Type";
                    sql += "   ,jvh_gst";
                    sql += "   ,max(party.acc_name) as jvh_party_name ";
                    sql += "   from ledgerh h";
                    sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += "   left join param st on h.jvh_state_id = st.param_pkid";
                    sql += "   left join acctm party on h.jvh_acc_id = party.acc_pkid";
                    sql += "   where h.rec_company_code = '{COMPCODE}'";
                    if (!all )
                        sql += "   and h.rec_branch_code = '{BRCODE}'";
                    sql += "   and h.jvh_type in ('IN','IN-ES') ";
                    sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(party.acc_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(h.jvh_gstin) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "   group by h.rec_branch_code,jvh_docno, h.jvh_cc_category, jvh_date, jvh_gstin,jvh_gst_type,jvh_sez, st.param_code, st.param_name,jvh_gst";
                    sql += "   order by h.rec_branch_code,jvh_docno ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    tot_inv_amt = 0;
                    tot_taxable_amt = 0;
                    tot_cgst_amt = 0;
                    tot_sgst_amt = 0;
                    tot_igst_amt = 0;
                    tot_gst_amt = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_cc_category = Dr["jvh_cc_category"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.jvh_gstin = Dr["jvh_gstin"].ToString();
                        mrow.jvh_net_amt = Lib.Conv2Decimal(Dr["jvh_net_amt"].ToString());
                        mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.jv_net_total = Lib.Conv2Decimal(Dr["jv_net_total"].ToString());
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Dr["jv_taxable_amt"].ToString());
                        mrow.jvh_gst_type = Dr["jvh_gst_type"].ToString();
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Dr["jv_cgst_amt"].ToString());
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Dr["jv_sgst_amt"].ToString());
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Dr["jv_igst_amt"].ToString());
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Dr["jv_gst_amt"].ToString());
                        mrow.jvh_sez = Dr["jvh_sez"].ToString();
                        mrow.jvh_state_name = Dr["jvh_state_name"].ToString();
                        mrow.rc = Dr["rc"].ToString();
                        mrow.jvh_invoice_type = Dr["jvh_invoice_type"].ToString();
                        mrow.jvh_gst = Dr["jvh_gst"].ToString();
                        mrow.jvh_party_name = Dr["jvh_party_name"].ToString();
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mList.Add(mrow);

                        tot_inv_amt += Lib.Conv2Decimal(mrow.jv_net_total.ToString());
                        tot_taxable_amt += Lib.Conv2Decimal(mrow.jv_taxable_amt.ToString());
                        tot_cgst_amt += Lib.Conv2Decimal(mrow.jv_cgst_amt.ToString());
                        tot_sgst_amt += Lib.Conv2Decimal(mrow.jv_sgst_amt.ToString());
                        tot_igst_amt += Lib.Conv2Decimal(mrow.jv_igst_amt.ToString());
                        tot_gst_amt += Lib.Conv2Decimal(mrow.jv_gst_amt.ToString());
                    }
                    if (mList.Count > 1)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.jvh_invoice_type = "TOTAL";
                        mrow.jv_net_total = Lib.Conv2Decimal(Lib.NumericFormat(tot_inv_amt.ToString(), 2));
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_taxable_amt.ToString(), 2));
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_cgst_amt.ToString(), 2));
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_sgst_amt.ToString(), 2));
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_igst_amt.ToString(), 2));
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_gst_amt.ToString(), 2));
                        mList.Add(mrow);
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintInvoiceReport();
                    }

                }

                if (format_type == "INVOICE-DETAILS")
                {
                    sql = " select jvh_docno,h.rec_branch_code  ";
                    sql += "   ,h.jvh_cc_category  ";
                    sql += "   ,jvh_date ";
                    sql += "   ,jvh_gstin  ";
                    sql += "   ,jvh_net_amt";
                    sql += "   ,jv_Credit";
                    sql += "   ,jv_net_total";
                    sql += "   ,jv_credit as jv_taxable_amt";
                    sql += "   ,jvh_gst_type  ";
                    sql += "   ,jv_gst_rate  ";
                    sql += "   ,jv_cgst_rate";
                    sql += "   ,jv_sgst_rate";
                    sql += "   ,jv_igst_rate";
                    sql += "   ,jv_cgst_amt";
                    sql += "   ,jv_sgst_amt";
                    sql += "   ,jv_igst_amt";
                    sql += "   ,jv_gst_amt";
                    sql += "   ,jvh_sez";
                    sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
                    sql += "   ,'N' as RC";
                    sql += "   ,case when jvh_sez='Y' then";
                    sql += "    case when jvh_gst='Y' and  jv_gst_amt > 0 then 'SEZ supplies with payment' else 'SEZ supplies without payment' end ";
                    sql += "    else 'Regular' end as jvh_invoice_Type";
                    sql += "   ,'' as ecomgstn";
                    sql += "   ,0 as cess";
                    sql += "   ,jvh_gst";
                    sql += "   ,party.acc_name as jvh_party_name ";
                    sql += "   ,sac.param_code as jv_sac_code ";
                    sql += "   ,acc.acc_code as jv_acc_code ";
                    sql += "   ,acc.acc_name as jv_acc_name ";
                    sql += "   from ledgerh h";
                    sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += "   left join param st on h.jvh_state_id = st.param_pkid";
                    sql += "   left join acctm party on h.jvh_acc_id = party.acc_pkid";
                    sql += "   left join acctm acc on t.jv_acc_id = acc.acc_pkid";
                    sql += "   left join param sac on t.jv_sac_id = sac.param_pkid";
                    sql += "   where h.rec_company_code = '{COMPCODE}'";
                    if (!all)
                        sql += "   and h.rec_branch_code = '{BRCODE}'";
                    sql += "   and h.jvh_type in ('IN','IN-ES') ";
                    sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(party.acc_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(h.jvh_gstin) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "   order by h.rec_branch_code,jvh_docno ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    tot_inv_amt = 0;
                    tot_taxable_amt = 0;
                    tot_cgst_amt = 0;
                    tot_sgst_amt = 0;
                    tot_igst_amt = 0;
                    tot_gst_amt = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_cc_category = Dr["jvh_cc_category"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.jvh_gstin = Dr["jvh_gstin"].ToString();
                        mrow.jvh_net_amt = Lib.Conv2Decimal(Dr["jvh_net_amt"].ToString());
                        mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.jv_net_total = Lib.Conv2Decimal(Dr["jv_net_total"].ToString());
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Dr["jv_taxable_amt"].ToString());
                        mrow.jvh_gst_type = Dr["jvh_gst_type"].ToString();
                        mrow.jv_gst_rate = Lib.Conv2Decimal(Dr["jv_gst_rate"].ToString());
                        mrow.jv_cgst_rate = Lib.Conv2Decimal(Dr["jv_cgst_rate"].ToString());
                        mrow.jv_sgst_rate = Lib.Conv2Decimal(Dr["jv_sgst_rate"].ToString());
                        mrow.jv_igst_rate = Lib.Conv2Decimal(Dr["jv_igst_rate"].ToString());
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Dr["jv_cgst_amt"].ToString());
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Dr["jv_sgst_amt"].ToString());
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Dr["jv_igst_amt"].ToString());
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Dr["jv_gst_amt"].ToString());
                        mrow.jvh_sez = Dr["jvh_sez"].ToString();
                        mrow.jvh_state_name = Dr["jvh_state_name"].ToString();
                        mrow.rc = Dr["rc"].ToString();
                        mrow.jvh_invoice_type = Dr["jvh_invoice_type"].ToString();
                        mrow.ecomgstn = Dr["ecomgstn"].ToString();
                        mrow.cess = Lib.Conv2Decimal(Dr["cess"].ToString());
                        mrow.jvh_gst = Dr["jvh_gst"].ToString();
                        mrow.jvh_party_name = Dr["jvh_party_name"].ToString();
                        mrow.jv_sac_code = Dr["jv_sac_code"].ToString();
                        mrow.jv_acc_name = Dr["jv_acc_name"].ToString();
                        mrow.jv_acc_code = Dr["jv_acc_code"].ToString();
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mList.Add(mrow);

                        tot_inv_amt += Lib.Conv2Decimal(mrow.jv_net_total.ToString());
                        tot_taxable_amt += Lib.Conv2Decimal(mrow.jv_taxable_amt.ToString());
                        tot_cgst_amt += Lib.Conv2Decimal(mrow.jv_cgst_amt.ToString());
                        tot_sgst_amt += Lib.Conv2Decimal(mrow.jv_sgst_amt.ToString());
                        tot_igst_amt += Lib.Conv2Decimal(mrow.jv_igst_amt.ToString());
                        tot_gst_amt += Lib.Conv2Decimal(mrow.jv_gst_amt.ToString());
                    }
                    if (mList.Count > 1)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.jvh_invoice_type = "TOTAL";
                        mrow.jv_net_total = Lib.Conv2Decimal(Lib.NumericFormat(tot_inv_amt.ToString(), 2));
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_taxable_amt.ToString(), 2));
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_cgst_amt.ToString(), 2));
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_sgst_amt.ToString(), 2));
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_igst_amt.ToString(), 2));
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_gst_amt.ToString(), 2));
                        mList.Add(mrow);
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintInvoicedetReport();
                    }

                }

                if (format_type == "PURCHASE")
                {
                    sql = " select jvh_docno,h.rec_branch_code ";
                    sql += "   ,h.jvh_cc_category  ";
                    sql += "   ,jvh_date ";
                    sql += "   ,h.jvh_org_invno ";
                    sql += "   ,h.jvh_org_invdt ";
                    sql += "   ,jvh_gstin  ";
                    sql += "   ,max(jvh_net_amt) as jvh_net_amt";
                    sql += "   ,sum(jv_credit) as jv_Credit";
                    sql += "   ,sum(jv_net_total) as jv_net_total";
                    sql += "   ,sum(case when jvh_type = 'PN' then jv_taxable_amt else jv_debit end ) as jv_taxable_amt";
                    sql += "   ,jvh_gst_type  ";
                    sql += "   ,sum(jv_cgst_amt) as jv_cgst_amt";
                    sql += "   ,sum(jv_sgst_amt) as jv_sgst_amt";
                    sql += "   ,sum(jv_igst_amt) as jv_igst_amt";
                    sql += "   ,sum(jv_gst_amt) as jv_gst_amt";
                    sql += "   ,jvh_sez";
                    sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
                    sql += "   ,jvh_rc as RC";
                    sql += "   ,case when jvh_sez='Y' then";
                    sql += "    case when jvh_gst='Y' and sum(jv_gst_amt)>0 then 'SEZ supplies with payment' else 'SEZ supplies without payment' end ";
                    sql += "    else 'Regular' end as jvh_invoice_Type";
                    sql += "   ,jvh_gst";
                    sql += "   ,max(party.acc_name) as jvh_party_name ";
                    sql += "   from ledgerh h";
                    sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_debit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += "   left join param st on h.jvh_state_id = st.param_pkid";
                    sql += "   left join acctm party on h.jvh_acc_id = party.acc_pkid";
                    sql += "   where h.rec_company_code = '{COMPCODE}'";
                    if (!all)
                        sql += "   and h.rec_branch_code = '{BRCODE}'";
                    sql += "   and h.jvh_type in ('PN','JV', 'BR','BP','CR','CP') ";
                   // sql += "   and nvl(h.jvh_rc,'N') = 'N'";
                    sql += "   and t.jv_gst_amt > 0 ";
                    sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(party.acc_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(h.jvh_gstin) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "   group by h.rec_branch_code,jvh_docno, h.jvh_cc_category, jvh_date,h.jvh_org_invno,h.jvh_org_invdt, jvh_gstin,jvh_gst_type,jvh_sez, st.param_code, st.param_name,jvh_gst,jvh_rc";
                    sql += "   order by h.rec_branch_code,jvh_docno,jvh_date ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    tot_inv_amt = 0;
                    tot_taxable_amt = 0;
                    tot_cgst_amt = 0;
                    tot_sgst_amt = 0;
                    tot_igst_amt = 0;
                    tot_gst_amt = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_cc_category = Dr["jvh_cc_category"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.jvh_org_invno = Dr["jvh_org_invno"].ToString();
                        mrow.jvh_org_invdt = Lib.DatetoStringDisplayformat(Dr["jvh_org_invdt"]);
                        mrow.jvh_gstin = Dr["jvh_gstin"].ToString();
                        mrow.jvh_net_amt = Lib.Conv2Decimal(Dr["jvh_net_amt"].ToString());
                        mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.jv_net_total = Lib.Conv2Decimal(Dr["jv_net_total"].ToString());
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Dr["jv_taxable_amt"].ToString());
                        mrow.jvh_gst_type = Dr["jvh_gst_type"].ToString();
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Dr["jv_cgst_amt"].ToString());
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Dr["jv_sgst_amt"].ToString());
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Dr["jv_igst_amt"].ToString());
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Dr["jv_gst_amt"].ToString());
                        mrow.jvh_sez = Dr["jvh_sez"].ToString();
                        mrow.jvh_state_name = Dr["jvh_state_name"].ToString();
                        mrow.rc = Dr["rc"].ToString();
                        mrow.jvh_invoice_type = Dr["jvh_invoice_type"].ToString();
                        mrow.jvh_gst = Dr["jvh_gst"].ToString();
                        mrow.jvh_party_name = Dr["jvh_party_name"].ToString();
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mList.Add(mrow);

                        tot_inv_amt += Lib.Conv2Decimal(mrow.jv_net_total.ToString());
                        tot_taxable_amt += Lib.Conv2Decimal(mrow.jv_taxable_amt.ToString());
                        tot_cgst_amt += Lib.Conv2Decimal(mrow.jv_cgst_amt.ToString());
                        tot_sgst_amt += Lib.Conv2Decimal(mrow.jv_sgst_amt.ToString());
                        tot_igst_amt += Lib.Conv2Decimal(mrow.jv_igst_amt.ToString());
                        tot_gst_amt += Lib.Conv2Decimal(mrow.jv_gst_amt.ToString());
                    }
                    if (mList.Count > 1)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.jvh_invoice_type = "TOTAL";
                        mrow.jv_net_total = Lib.Conv2Decimal(Lib.NumericFormat(tot_inv_amt.ToString(), 2));
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_taxable_amt.ToString(), 2));
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_cgst_amt.ToString(), 2));
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_sgst_amt.ToString(), 2));
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_igst_amt.ToString(), 2));
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_gst_amt.ToString(), 2));
                        mList.Add(mrow);
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintpurchaseReport();
                    }

                }

                if (format_type == "PURCHASE-DETAILS")
                {
                    sql = " select jvh_docno,h.rec_branch_code ";
                    sql += "   ,h.jvh_cc_category  ";
                    sql += "   ,h.jvh_type  ";
                    sql += "   ,jvh_date ";
                    sql += "   ,h.jvh_org_invno ";
                    sql += "   ,h.jvh_org_invdt ";
                    sql += "   ,jvh_gstin  ";
                    sql += "   ,jvh_net_amt";
                    sql += "   ,jv_Credit";
                    sql += "   ,jv_net_total";
                    sql += "   ,case when jvh_type = 'PN' then jv_taxable_amt else jv_debit end  as jv_taxable_amt";
                    sql += "   ,jvh_gst_type  ";
                    sql += "   ,jv_gst_rate  ";
                    sql += "   ,jv_cgst_rate";
                    sql += "   ,jv_sgst_rate";
                    sql += "   ,jv_igst_rate";
                    sql += "   ,jv_cgst_amt";
                    sql += "   ,jv_sgst_amt";
                    sql += "   ,jv_igst_amt";
                    sql += "   ,jv_gst_amt";
                    sql += "   ,jvh_sez";
                    sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
                    sql += "   ,jvh_rc as RC";
                    sql += "   ,case when jvh_sez='Y' then";
                    sql += "    case when jvh_gst='Y' and  jv_gst_amt > 0 then 'SEZ supplies with payment' else 'SEZ supplies without payment' end ";
                    sql += "    else 'Regular' end as jvh_invoice_Type";
                    sql += "   ,'' as ecomgstn";
                    sql += "   ,0 as cess";
                    sql += "   ,jvh_gst";
                    sql += "   ,party.acc_name as jvh_party_name ";
                    sql += "   ,sac.param_code as jv_sac_code ";
                    sql += "   ,acc.acc_code as jv_acc_code ";
                    sql += "   ,acc.acc_name as jv_acc_name ";
                    sql += "   from ledgerh h";
                    sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_debit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += "   left join param st on h.jvh_state_id = st.param_pkid";
                    sql += "   left join acctm party on h.jvh_acc_id = party.acc_pkid";
                    sql += "   left join acctm acc on t.jv_acc_id = acc.acc_pkid";
                    sql += "   left join param sac on t.jv_sac_id = sac.param_pkid";
                    sql += "   where h.rec_company_code = '{COMPCODE}'";
                    if (!all)
                        sql += "   and h.rec_branch_code = '{BRCODE}'";
                    sql += "   and h.jvh_type in ('PN','JV', 'BR','BP','CR','CP') ";
                   // sql += "   and nvl(h.jvh_rc,'N') = 'N'";
                    sql += "   and t.jv_gst_amt > 0 ";
                    sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(party.acc_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(h.jvh_gstin) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "   order by h.rec_branch_code,jvh_docno,jvh_date ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    tot_inv_amt = 0;
                    tot_taxable_amt = 0;
                    tot_cgst_amt = 0;
                    tot_sgst_amt = 0;
                    tot_igst_amt = 0;
                    tot_gst_amt = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_type = Dr["jvh_type"].ToString();
                        mrow.jvh_cc_category = Dr["jvh_cc_category"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.jvh_org_invno = Dr["jvh_org_invno"].ToString();
                        mrow.jvh_org_invdt = Lib.DatetoStringDisplayformat(Dr["jvh_org_invdt"]);
                        mrow.jvh_gstin = Dr["jvh_gstin"].ToString();
                        mrow.jvh_net_amt = Lib.Conv2Decimal(Dr["jvh_net_amt"].ToString());
                        mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.jv_net_total = Lib.Conv2Decimal(Dr["jv_net_total"].ToString());
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Dr["jv_taxable_amt"].ToString());
                        mrow.jvh_gst_type = Dr["jvh_gst_type"].ToString();
                        mrow.jv_gst_rate = Lib.Conv2Decimal(Dr["jv_gst_rate"].ToString());
                        mrow.jv_cgst_rate = Lib.Conv2Decimal(Dr["jv_cgst_rate"].ToString());
                        mrow.jv_sgst_rate = Lib.Conv2Decimal(Dr["jv_sgst_rate"].ToString());
                        mrow.jv_igst_rate = Lib.Conv2Decimal(Dr["jv_igst_rate"].ToString());
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Dr["jv_cgst_amt"].ToString());
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Dr["jv_sgst_amt"].ToString());
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Dr["jv_igst_amt"].ToString());
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Dr["jv_gst_amt"].ToString());
                        mrow.jvh_sez = Dr["jvh_sez"].ToString();
                        mrow.jvh_state_name = Dr["jvh_state_name"].ToString();
                        mrow.rc = Dr["rc"].ToString();
                        mrow.jvh_invoice_type = Dr["jvh_invoice_type"].ToString();
                        mrow.ecomgstn = Dr["ecomgstn"].ToString();
                        mrow.cess = Lib.Conv2Decimal(Dr["cess"].ToString());
                        mrow.jvh_gst = Dr["jvh_gst"].ToString();
                        mrow.jvh_party_name = Dr["jvh_party_name"].ToString();
                        mrow.jv_sac_code = Dr["jv_sac_code"].ToString();
                        mrow.jv_acc_name = Dr["jv_acc_name"].ToString();
                        mrow.jv_acc_code = Dr["jv_acc_code"].ToString();
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mList.Add(mrow);

                        tot_inv_amt += Lib.Conv2Decimal(mrow.jv_net_total.ToString());
                        tot_taxable_amt += Lib.Conv2Decimal(mrow.jv_taxable_amt.ToString());
                        tot_cgst_amt += Lib.Conv2Decimal(mrow.jv_cgst_amt.ToString());
                        tot_sgst_amt += Lib.Conv2Decimal(mrow.jv_sgst_amt.ToString());
                        tot_igst_amt += Lib.Conv2Decimal(mrow.jv_igst_amt.ToString());
                        tot_gst_amt += Lib.Conv2Decimal(mrow.jv_gst_amt.ToString());
                    }
                    if (mList.Count > 1)
                    {
                        mrow = new GstReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.jvh_invoice_type = "TOTAL";
                        mrow.jv_net_total = Lib.Conv2Decimal(Lib.NumericFormat(tot_inv_amt.ToString(), 2));
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_taxable_amt.ToString(), 2));
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_cgst_amt.ToString(), 2));
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_sgst_amt.ToString(), 2));
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_igst_amt.ToString(), 2));
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_gst_amt.ToString(), 2));
                        mList.Add(mrow);
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintPurchaseDetReport();
                    }

                }

                if (format_type == "FORM 3B")
                {
                    string swhere = "";

                    swhere = " where h.rec_company_code = '{COMPCODE}'";
                    if (this.branch_code != "")
                        swhere += "   and h.rec_branch_code = '{BRCODE}'";
                    swhere += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                    // SALES NON ZERO
                    // Regular and Rate > 0
                    sql = " select 10 as slno, 'OUT' as grp,'3.1 OUTWARD SUPPLY (OTHER THAN ZERO RATED)' as stype, null as state_name, ";
                   // sql += " sum( case when nvl(jv_gst_rate,0) > 0 then jv_taxable_amt else 0 end ) as taxable_amt, ";
                    sql += " sum( case when nvl(jv_gst_rate,0) > 0 then jv_credit else 0 end ) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jvh_gst = 'Y' ";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " and nvl(jvh_sez,'N')='N'  ";

                    sql += " union all";

                    //SALES ZERO
                    // Zero Rated other than sez
                    sql += " select 20 as slno, 'OUT' as grp, '3.1 OUTWARD SUPPLY (NIL/EXEMPT)' as stype, null as state_name,  ";
                    sql += " sum(jv_credit) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN','IN-ES')";
                    /*
                    if (company_code == "SGT")
                        sql += " and jvh_gst = 'N' ";
                    else
                        sql += " and jvh_gst = 'Y' ";
                    */
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " and nvl(jvh_sez,'N')='N'  ";
                    sql += " and nvl(jv_gst_rate,0) <= 0  ";

                    sql += " union all";

                    // Sez without tax
                    sql += " select 30 as slno,'OUT' as grp, '3.1 OUTWARD SUPPLY (SPECIAL ECONOMIC ZONE)' as stype, null as state_name,  ";
                    sql += " sum(jv_credit) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " and nvl(jvh_sez,'N')='Y'  ";
                    sql += " and nvl(jv_gst_rate,0) <= 0  ";
                    //sql += " and h.rec_company_code != 'SGT'  ";

                    sql += " union all";

                    // Sez with tax
                    sql += " select 35 as slno, 'OUT' as grp,'3.1.1 OUTWARD SUPPLY (SPECIAL ECONOMIC ZONE)' as stype, null as state_name, ";
                    sql += " sum( case when nvl(jv_gst_rate,0) > 0 then jv_credit else 0 end ) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jv_credit > 0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " and nvl(jvh_sez,'N')='Y'  ";
                    sql += " and jvh_gst = 'Y' ";
                    sql += " and nvl(jv_gst_rate,0) > 0  ";

                    sql += " union all";

                    //PURCHASE ONLY RC

                    sql += " select 70 as slno,'NA' as grp,'3.1 INWARD SUPPLY (LIABLE TO REVERSE CHARGE)' as stype, null as state_name,  ";
                    sql += " sum(case when jvh_type = 'PN' then jv_taxable_amt else jv_debit end) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('PN','DN','JV', 'BR','BP','CR','CP')";
                    sql += " and jvh_gst = 'Y'  ";
                    sql += " and jvh_rc = 'Y'";
                    sql += " and jv_debit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";

                    sql += " union all";

                    //PURCHASE WITHOUT RC

                    sql += " select  40 as slno,'IN' as grp,'4. ALL OTHER ITC' as stype, null as state_name,  ";
                    sql += " sum(case when jvh_type = 'PN' then jv_taxable_amt else jv_debit end) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('PN','DN','JV', 'BR','BP','CR','CP')";
                    sql += " and jvh_gst = 'Y'  ";
                    sql += " and jvh_rc = 'N'";
                    sql += " and jv_debit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";

                    sql += " union all";

                    //SUPPLY TO UNREGI. INTER-STATE 

                    sql += " select 60 as slno,'NA' as grp,'3.2 SUPPLIES TO UN-REG. (INTER-STATE)' as stype,to_char(st.param_code || '-' || initcap(st.param_name)) as state_name, ";
                  //  sql += " sum( case when nvl(jv_gst_rate,0) > 0 then  jv_taxable_amt else 0 end ) as taxable_amt , ";
                    sql += " sum( case when nvl(jv_gst_rate,0) > 0 then  jv_credit else 0 end ) as taxable_amt , ";
                    sql += " 0 as cgst_amt, 0 as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt  ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += " left join param st on h.jvh_state_id = st.param_pkid";
                    sql += " " + swhere;
                    sql += " and jvh_gst_type  = 'INTER-STATE'  ";
                    sql += " and (jvh_gstin is null or jvh_gstin like '%.%' ) ";
                    sql += " group by  st.param_code, st.param_name ";


                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    decimal itot = 0;
                    decimal iTax_Out = 0;
                    decimal iCgst_Out = 0;
                    decimal iSgst_Out = 0;
                    decimal iIgst_Out = 0;

                    int iCtr = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        iCtr++;
                        //if (iCtr == 1 || iCtr == 2 || iCtr == 3) // 3 is removed since it is reverse charge
                        if (Dr["grp"].ToString() == "OUT")
                        {
                            iTax_Out += Lib.Convert2Decimal(Dr["TAXABLE_AMT"].ToString());
                            iCgst_Out += Lib.Convert2Decimal(Dr["CGST_AMT"].ToString());
                            iSgst_Out += Lib.Convert2Decimal(Dr["SGST_AMT"].ToString());
                            iIgst_Out += Lib.Convert2Decimal(Dr["IGST_AMT"].ToString());
                        }
                        if (Dr["grp"].ToString() == "IN")
                        {
                            iTax_Out -= Lib.Convert2Decimal(Dr["TAXABLE_AMT"].ToString());
                            iCgst_Out -= Lib.Convert2Decimal(Dr["CGST_AMT"].ToString());
                            iSgst_Out -= Lib.Convert2Decimal(Dr["SGST_AMT"].ToString());
                            iIgst_Out -= Lib.Convert2Decimal(Dr["IGST_AMT"].ToString());
                        }

                    }

                    DataRow dr1 = Dt_List.NewRow();
                    dr1["slno"] = 50;
                    dr1["STYPE"] = "TAX LIABILITY ";
                    //dr1["TAXABLE_AMT"] = iTax_Out;
                    dr1["CGST_AMT"] = iCgst_Out;
                    dr1["SGST_AMT"] = iSgst_Out;
                    dr1["IGST_AMT"] = iIgst_Out;
                    dr1["GST_AMT"] = iCgst_Out + iSgst_Out + iIgst_Out;

                    Dt_List.Rows.Add(dr1);

                    foreach (DataRow Dr in Dt_List.Select("1=1", "slno"))
                    {
                        mrow = new GstReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.jvh_gst_type = Dr["stype"].ToString();
                        mrow.jvh_state_name = Dr["state_name"].ToString();
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Dr["taxable_amt"].ToString());
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Dr["cgst_amt"].ToString());
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Dr["sgst_amt"].ToString());
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Dr["igst_amt"].ToString());
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Dr["gst_amt"].ToString());
                        mList.Add(mrow);
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintForm3BReport("");
                    }

                }
                if (format_type == "FORM 3B-RATE WISE")
                {
                    string swhere = "";

                    swhere = " where h.rec_company_code = '{COMPCODE}'";
                    if (this.branch_code != "")
                        swhere += "   and h.rec_branch_code = '{BRCODE}'";
                    swhere += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                    // SALES NON ZERO

                    sql = " select 10 as slno, 'OUT' as grp,'3.1 OUTWARD SUPPLY (OTHER THAN ZERO RATED)' as stype, null as state_name, ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum( case when nvl(jv_gst_rate,0) > 0 then jv_taxable_amt else 0 end ) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jvh_gst = 'Y' ";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";

                    //SALES ZERO

                    sql += " select 20 as slno, 'OUT' as grp, '3.1 OUTWARD SUPPLY (NIL/EXEMPT)' as stype, null as state_name,  ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum(jv_taxable_amt) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN','IN-ES')";
                    sql += " and jvh_gst = 'Y'";
                    sql += " and nvl(jv_gst_rate,0) <= 0  ";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";

                    //SALES ZERO - INTRA-STATE SUPPLIES TO REGISTERED PERSON

                    sql += " select 21 as slno, 'NA' as grp, '3.1 OUTWARD SUPPLY (NIL/EXEMPT)-REG. (INTRA-STATE)' as stype, null as state_name,  ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum(jv_taxable_amt) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jvh_gst = 'Y'";
                    sql += " and nvl(jv_gst_rate,0) <= 0  ";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " and jvh_gst_type  = 'INTRA-STATE' ";
                    sql += " and jvh_gstin is not null";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";
                    
                    //SALES ZERO - INTRA-STATE SUPPLIES TO UNREGISTERED PERSON

                    sql += " select 22 as slno, 'NA' as grp, '3.1 OUTWARD SUPPLY (NIL/EXEMPT)-UNREG. (INTRA-STATE)' as stype, null as state_name,  ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum(jv_taxable_amt) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jvh_gst = 'Y'";
                    sql += " and nvl(jv_gst_rate,0) <= 0  ";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " and jvh_gst_type  = 'INTRA-STATE' ";
                    sql += " and jvh_gstin is null";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";

                    //SALES ZERO - INTER-STATE SUPPLIES TO REGISTERED PERSON

                    sql += " select 23 as slno, 'NA' as grp, '3.1 OUTWARD SUPPLY (NIL/EXEMPT)-REG. (INTER-STATE)' as stype, null as state_name,  ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum(jv_taxable_amt) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jvh_gst = 'Y'";
                    sql += " and nvl(jv_gst_rate,0) <= 0  ";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " and jvh_gst_type  = 'INTER-STATE' ";
                    sql += " and jvh_gstin is not null";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";

                    //SALES ZERO - INTER-STATE SUPPLIES TO UNREGISTERED PERSON

                    sql += " select 24 as slno, 'NA' as grp, '3.1 OUTWARD SUPPLY (NIL/EXEMPT)-UNREG. (INTER-STATE)' as stype, null as state_name,  ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum(jv_taxable_amt) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jvh_gst = 'Y'";
                    sql += " and nvl(jv_gst_rate,0) <= 0  ";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " and jvh_gst_type  = 'INTER-STATE' ";
                    sql += " and jvh_gstin is null";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";

                    sql += " select 30 as slno, 'OUT' as grp, '3.1 OUTWARD SUPPLY (SPECIAL ECONOMIC ZONE)' as stype, null as state_name,  ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum(jv_taxable_amt) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('IN')";
                    sql += " and jvh_gst = 'N'";
                    sql += " and nvl(jv_gst_rate,0) <= 0  ";
                    sql += " and jv_credit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";

                    //PURCHASE ONLY RC

                    sql += " select 70 as slno,'NA' as grp,'3.1 INWARD SUPPLY (LIABLE TO REVERSE CHARGE)' as stype, null as state_name,  ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum(case when jvh_type = 'PN' then jv_taxable_amt else jv_debit end) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('PN','DN','JV', 'BR','BP','CR','CP')";
                    sql += " and jvh_gst = 'Y'  ";
                    sql += " and jvh_rc = 'Y'";
                    sql += " and jv_debit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";

                    //PURCHASE WITHOUT RC

                    sql += " select  40 as slno,'IN' as grp,'4. ALL OTHER ITC' as stype, null as state_name,  ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum(case when jvh_type = 'PN' then jv_taxable_amt else jv_debit end) as taxable_amt, ";
                    sql += " sum(jv_cgst_amt) as cgst_amt, ";
                    sql += " sum(jv_sgst_amt) as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h ";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id  ";
                    sql += " " + swhere;
                    sql += " and jvh_type in ('PN','DN','JV', 'BR','BP','CR','CP')";
                    sql += " and jvh_gst = 'Y'  ";
                    sql += " and jvh_rc = 'N'";
                    sql += " and jv_debit >0  ";
                    sql += " and nvl(jv_row_type,'JV') not in('HEADER','GST')   ";
                    sql += " group by jv_gst_rate ";

                    sql += " union all";

                    //SUPPLY TO UNREGI. INTER-STATE 

                    sql += " select 60 as slno,'NA' as grp,'3.2 SUPPLIES TO UN-REG. (INTER-STATE)' as stype,to_char(st.param_code || '-' || initcap(st.param_name)) as state_name, ";
                    sql += " jv_gst_rate as rate,";
                    sql += " sum( case when nvl(jv_gst_rate,0) > 0 then  jv_taxable_amt else 0 end ) as taxable_amt , ";
                    sql += " 0 as cgst_amt, 0 as sgst_amt, ";
                    sql += " sum(jv_igst_amt) as igst_amt  ";
                    sql += " ,sum(jv_gst_amt) as gst_amt ";
                    sql += " from ledgerh h";
                    sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id   and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
                    sql += " left join param st on h.jvh_state_id = st.param_pkid";
                    sql += " " + swhere;
                    sql += " and jvh_gst_type  = 'INTER-STATE'  ";
                    sql += " and (jvh_gstin is null or jvh_gstin like '%.%' ) ";
                    sql += " group by  st.param_code, st.param_name,jv_gst_rate ";


                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    decimal itot = 0;
                    decimal iTax_Out = 0;
                    decimal iCgst_Out = 0;
                    decimal iSgst_Out = 0;
                    decimal iIgst_Out = 0;

                    int iCtr = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        iCtr++;
                        //if (iCtr == 1 || iCtr == 2 || iCtr == 3) // 3 is removed since it is reverse charge
                        if (Dr["grp"].ToString() == "OUT")
                        {
                            iTax_Out += Lib.Convert2Decimal(Dr["TAXABLE_AMT"].ToString());
                            iCgst_Out += Lib.Convert2Decimal(Dr["CGST_AMT"].ToString());
                            iSgst_Out += Lib.Convert2Decimal(Dr["SGST_AMT"].ToString());
                            iIgst_Out += Lib.Convert2Decimal(Dr["IGST_AMT"].ToString());
                        }
                        if (Dr["grp"].ToString() == "IN")
                        {
                            iTax_Out -= Lib.Convert2Decimal(Dr["TAXABLE_AMT"].ToString());
                            iCgst_Out -= Lib.Convert2Decimal(Dr["CGST_AMT"].ToString());
                            iSgst_Out -= Lib.Convert2Decimal(Dr["SGST_AMT"].ToString());
                            iIgst_Out -= Lib.Convert2Decimal(Dr["IGST_AMT"].ToString());
                        }

                    }

                    DataRow dr1 = Dt_List.NewRow();
                    dr1["slno"] = 50;
                    dr1["STYPE"] = "TAX LIABILITY ";
                    //dr1["TAXABLE_AMT"] = iTax_Out;
                    dr1["CGST_AMT"] = iCgst_Out;
                    dr1["SGST_AMT"] = iSgst_Out;
                    dr1["IGST_AMT"] = iIgst_Out;
                    dr1["GST_AMT"] = iCgst_Out + iSgst_Out + iIgst_Out;

                    Dt_List.Rows.Add(dr1);

                    foreach (DataRow Dr in Dt_List.Select("1=1", "slno"))
                    {
                        mrow = new GstReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.jvh_gst_type = Dr["stype"].ToString();
                        mrow.jvh_state_name = Dr["state_name"].ToString();
                        mrow.jv_taxable_amt = Lib.Conv2Decimal(Dr["taxable_amt"].ToString());
                        mrow.jv_gst_rate = Lib.Conv2Decimal(Dr["rate"].ToString());
                        mrow.jv_cgst_amt = Lib.Conv2Decimal(Dr["cgst_amt"].ToString());
                        mrow.jv_sgst_amt = Lib.Conv2Decimal(Dr["sgst_amt"].ToString());
                        mrow.jv_igst_amt = Lib.Conv2Decimal(Dr["igst_amt"].ToString());
                        mrow.jv_gst_amt = Lib.Conv2Decimal(Dr["gst_amt"].ToString());
                        mList.Add(mrow);
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintForm3BReport("RATE");
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

            RetData.Add("type", type);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", mList);
            RetData.Add("generatemsg", GenerateMsg);
            return RetData;
        }

        private void PrintGSTR1Report()
        {
           
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
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
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }



                File_Display_Name = "GstReport.xls";
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
                    WS.Columns[1].Width = 256 * 15;
                    WS.Columns[2].Width = 256 * 12;
                    WS.Columns[3].Width = 256 * 17;
                    WS.Columns[4].Width = 256 * 45;
                    WS.Columns[5].Width = 256 * 12;
                    WS.Columns[6].Width = 256 * 5;
                    WS.Columns[7].Width = 256 * 17;
                    WS.Columns[8].Width = 256 * 12;
                    WS.Columns[9].Width = 256 * 8;
                    WS.Columns[10].Width = 256 * 3;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 10;
                    WS.Columns[13].Width = 256 * 5;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 5;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                }
                else
                {

                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 13;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 12;
                    WS.Columns[4].Width = 256 * 17;
                    WS.Columns[5].Width = 256 * 45;
                    WS.Columns[6].Width = 256 * 12;
                    WS.Columns[7].Width = 256 * 5;
                    WS.Columns[8].Width = 256 * 17;
                    WS.Columns[9].Width = 256 * 12;
                    WS.Columns[10].Width = 256 * 8;
                    WS.Columns[11].Width = 256 * 3;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 10;
                    WS.Columns[14].Width = 256 * 5;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 5;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 12;
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
                Lib.WriteData(WS, iRow, 1, "GST REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "INVOICE#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GSTIN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTOMER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SEZ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PLACE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST[Y/N]", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RC", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ECOM-GSTN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAXABLE-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CESS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (GstReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gstin, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_cc_category, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_sez, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_state_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rc, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ecomgstn, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.taxable_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cess, _Color, false, "", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.taxable_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        private void PrintInvoiceReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(all)
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                else
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                

                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }



                File_Display_Name = "GstReport.xls";
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
                    WS.Columns[2].Width = 256 * 12;
                    WS.Columns[3].Width = 256 * 17;
                    WS.Columns[4].Width = 256 * 45;
                    WS.Columns[5].Width = 256 * 12;
                    WS.Columns[6].Width = 256 * 5;
                    WS.Columns[7].Width = 256 * 17;
                    WS.Columns[8].Width = 256 * 12;
                    WS.Columns[9].Width = 256 * 8;
                    WS.Columns[10].Width = 256 * 3;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 12;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[2].Width = 256 * 12;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 12;
                    WS.Columns[4].Width = 256 * 17;
                    WS.Columns[5].Width = 256 * 45;
                    WS.Columns[6].Width = 256 * 12;
                    WS.Columns[7].Width = 256 * 5;
                    WS.Columns[8].Width = 256 * 17;
                    WS.Columns[9].Width = 256 * 12;
                    WS.Columns[10].Width = 256 * 8;
                    WS.Columns[11].Width = 256 * 3;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
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
                Lib.WriteData(WS, iRow, 1, "INVOICE REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "INVOICE#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GSTIN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTOMER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SEZ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PLACE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST[Y/N]", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RC", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAXABLE-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (GstReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gstin, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_cc_category, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_sez, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_state_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rc, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_net_total, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_net_total, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        private void PrintInvoicedetReport()
        {

            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
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
                    mSearchData.Add("branch_code","HOCPL");
                }
                

                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }



                File_Display_Name = "GstReport.xls";
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
                    WS.Columns[2].Width = 256 * 12;
                    WS.Columns[3].Width = 256 * 17;
                    WS.Columns[4].Width = 256 * 45;
                    WS.Columns[5].Width = 256 * 12;
                    WS.Columns[6].Width = 256 * 5;
                    WS.Columns[7].Width = 256 * 17;
                    WS.Columns[8].Width = 256 * 12;
                    WS.Columns[9].Width = 256 * 8;
                    WS.Columns[10].Width = 256 * 3;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 15;
                    WS.Columns[14].Width = 256 * 20;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 8;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 8;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 8;
                    WS.Columns[22].Width = 256 * 12;
                    WS.Columns[23].Width = 256 * 8;
                    WS.Columns[24].Width = 256 * 12;

                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 13;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 12;
                    WS.Columns[4].Width = 256 * 17;
                    WS.Columns[5].Width = 256 * 45;
                    WS.Columns[6].Width = 256 * 12;
                    WS.Columns[7].Width = 256 * 5;
                    WS.Columns[8].Width = 256 * 17;
                    WS.Columns[9].Width = 256 * 12;
                    WS.Columns[10].Width = 256 * 8;
                    WS.Columns[11].Width = 256 * 3;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 15;
                    WS.Columns[14].Width = 256 * 15;
                    WS.Columns[15].Width = 256 * 20;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 8;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 8;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[22].Width = 256 * 8;
                    WS.Columns[23].Width = 256 * 12;
                    WS.Columns[24].Width = 256 * 8;
                    WS.Columns[25].Width = 256 * 12;

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
                Lib.WriteData(WS, iRow, 1, "INVOICE DETAIL REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "INVOICE#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GSTIN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTOMER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SEZ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PLACE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST[Y/N]", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RC", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SAC-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ACC-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ACC-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAXABLE-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (GstReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gstin, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_cc_category, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_sez, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_state_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rc, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sac_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_acc_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_acc_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_net_total, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_net_total, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        private void PrintpurchaseReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
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
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }



                File_Display_Name = "GstReport.xls";
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
                    WS.Columns[2].Width = 256 * 12;
                    WS.Columns[3].Width = 256 * 15;
                    WS.Columns[4].Width = 256 * 12;
                    WS.Columns[5].Width = 256 * 17;
                    WS.Columns[6].Width = 256 * 45;
                    WS.Columns[7].Width = 256 * 12;
                    WS.Columns[8].Width = 256 * 5;
                    WS.Columns[9].Width = 256 * 17;
                    WS.Columns[10].Width = 256 * 12;
                    WS.Columns[11].Width = 256 * 8;
                    WS.Columns[12].Width = 256 * 8;
                    WS.Columns[13].Width = 256 * 15;
                    WS.Columns[14].Width = 256 * 12;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;

                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 13;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 12;
                    WS.Columns[4].Width = 256 * 15;
                    WS.Columns[5].Width = 256 * 12;
                    WS.Columns[6].Width = 256 * 17;
                    WS.Columns[7].Width = 256 * 45;
                    WS.Columns[8].Width = 256 * 12;
                    WS.Columns[9].Width = 256 * 5;
                    WS.Columns[10].Width = 256 * 17;
                    WS.Columns[11].Width = 256 * 12;
                    WS.Columns[12].Width = 256 * 8;
                    WS.Columns[13].Width = 256 * 8;
                    WS.Columns[14].Width = 256 * 15;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;

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
                Lib.WriteData(WS, iRow, 1, "PURCHASE REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "OUR-REF#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SUP-INV#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GSTIN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTOMER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SEZ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PLACE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST[Y/N]", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RC[Y/N]", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAXABLE-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (GstReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_org_invno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_org_invdt, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gstin, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_cc_category, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_sez, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_state_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rc, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_net_total, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_net_total, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        private void PrintPurchaseDetReport()
        {

            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
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
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }



                File_Display_Name = "GstReport.xls";
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
                    WS.Columns[2].Width = 256 * 12;
                    WS.Columns[3].Width = 256 * 5;
                    WS.Columns[4].Width = 256 * 15;
                    WS.Columns[5].Width = 256 * 12;
                    WS.Columns[6].Width = 256 * 17;
                    WS.Columns[7].Width = 256 * 45;
                    WS.Columns[8].Width = 256 * 12;
                    WS.Columns[9].Width = 256 * 5;
                    WS.Columns[10].Width = 256 * 17;
                    WS.Columns[11].Width = 256 * 12;
                    WS.Columns[12].Width = 256 * 8;
                    WS.Columns[13].Width = 256 * 8;
                    WS.Columns[14].Width = 256 * 15;
                    WS.Columns[15].Width = 256 * 15;
                    WS.Columns[16].Width = 256 * 15;
                    WS.Columns[17].Width = 256 * 20;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 8;
                    WS.Columns[21].Width = 256 * 12;
                    WS.Columns[22].Width = 256 * 8;
                    WS.Columns[23].Width = 256 * 12;
                    WS.Columns[24].Width = 256 * 8;
                    WS.Columns[25].Width = 256 * 12;
                    WS.Columns[26].Width = 256 * 8;
                    WS.Columns[27].Width = 256 * 12;
                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 12;
                    WS.Columns[4].Width = 256 * 5;
                    WS.Columns[5].Width = 256 * 15;
                    WS.Columns[6].Width = 256 * 12;
                    WS.Columns[7].Width = 256 * 17;
                    WS.Columns[8].Width = 256 * 45;
                    WS.Columns[9].Width = 256 * 12;
                    WS.Columns[10].Width = 256 * 5;
                    WS.Columns[11].Width = 256 * 17;
                    WS.Columns[12].Width = 256 * 12;
                    WS.Columns[13].Width = 256 * 8;
                    WS.Columns[14].Width = 256 * 8;
                    WS.Columns[15].Width = 256 * 15;
                    WS.Columns[16].Width = 256 * 15;
                    WS.Columns[17].Width = 256 * 15;
                    WS.Columns[18].Width = 256 * 20;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 12;
                    WS.Columns[21].Width = 256 * 8;
                    WS.Columns[22].Width = 256 * 12;
                    WS.Columns[23].Width = 256 * 8;
                    WS.Columns[24].Width = 256 * 12;
                    WS.Columns[25].Width = 256 * 8;
                    WS.Columns[26].Width = 256 * 12;
                    WS.Columns[27].Width = 256 * 8;
                    WS.Columns[28].Width = 256 * 12;
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
                Lib.WriteData(WS, iRow, 1, "PURCHASE DETAIL REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "OUR-REF#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SUP-INV#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GSTIN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTOMER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CC-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SEZ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PLACE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST[Y/N]", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RC[Y/N]", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SAC-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ACC-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ACC-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAXABLE-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (GstReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_org_invno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_org_invdt, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gstin, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_cc_category, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_sez, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_state_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rc, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sac_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_acc_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_acc_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_net_total, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_invoice_type, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_net_total, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        private void PrintForm3BReport(string stype)
        {

            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
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
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }



                File_Display_Name = "GstReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 50;
                WS.Columns[2].Width = 256 * 20;
                WS.Columns[3].Width = 256 * 12;
                WS.Columns[4].Width = 256 * 12;
                WS.Columns[5].Width = 256 * 12;
                WS.Columns[6].Width = 256 * 12;
                WS.Columns[7].Width = 256 * 12;
                WS.Columns[8].Width = 256 * 12;


                iRow = 0; iCol = 1;

                WS.Columns[3].Style.NumberFormat = "#0.00";
                WS.Columns[4].Style.NumberFormat = "#0.00";
                WS.Columns[5].Style.NumberFormat = "#0.00";
                WS.Columns[6].Style.NumberFormat = "#0.00";
                WS.Columns[7].Style.NumberFormat = "#0.00";
                WS.Columns[8].Style.NumberFormat = "#0.00";

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
                if (stype == "RATE")
                    str = "FORM 3B RATE WISE REPORT ";
                else
                    str = "FORM 3B REPORT ";
                Lib.WriteData(WS, iRow, 1,str , _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "DESCRIPTION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAXABLE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                if (stype == "RATE")
                    Lib.WriteData(WS, iRow, iCol++, "RATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CGST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SGST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IGST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (GstReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_gst_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_state_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_taxable_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        if (stype == "RATE")
                            Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_cgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_sgst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_igst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gst_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            //Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "SALES EXECUTIVE");
            //RetData.Add("smanlist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "CITY");
            //RetData.Add("citylist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "STATE");
            //RetData.Add("statelist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;
        }

        private void GenerateGSTR1()
        {
            GenerateMsg = "";
            Con_Oracle = new DBConnection();
            /*
            sql = " select * from (";
            sql += " select jvh_docno ";
            sql += " ,jvh_date";
            sql += " ,sum(jv_net_total) as jv_net_total";
            sql += " ,sum(jv_taxable_amt) as jv_taxable_amt";
            sql += " ,jv_gst_rate";
            sql += " ,sum(jv_gst_amt) as jv_gst_amt";
            sql += " ,st.param_code||'-'||initcap(st.param_name) as jvh_state_name";
            sql += " ,'' as ecomgstn";
            sql += " ,0 as cess";
            sql += " from ledgerh h";
            sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id and jv_credit>0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += " left join param st on h.jvh_state_id = st.param_pkid";
            sql += " where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += "   and h.rec_branch_code = '{BRCODE}'";
            sql += " and h.jvh_type in ('IN') ";
            sql += " and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += " and h.jvh_gst_type  = 'INTER-STATE' ";
            sql += " and h.jvh_gstin is null ";
            sql += " group by jvh_docno,jvh_date,jv_gst_rate,st.param_code,st.param_name";
            sql += " )a where jv_net_total > 250000";
            sql += "  order  by jvh_docno,jv_gst_rate";
            */

            sql = " select * from (";
            sql += " select jvh_docno ";
            sql += " ,jvh_date";
            sql += " ,max(jvh_net_amt) as inv_amt";
            sql += " ,sum(jv_credit) as taxable_amt";
            sql += " ,0 as jv_gst_rate";
            sql += " ,sum(jv_gst_amt) as jv_gst_amt";
            sql += " ,st.param_code||'-'||initcap(st.param_name) as jvh_state_name";
            sql += " ,'' as ecomgstn";
            sql += " ,0 as cess";
            sql += " from ledgerh h";
            sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id and jv_credit>0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += " left join param st on h.jvh_state_id = st.param_pkid";
            sql += " where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += "   and h.rec_branch_code = '{BRCODE}'";
            sql += " and h.jvh_type in ('IN') ";
            sql += " and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += " and h.jvh_gst_type  = 'INTER-STATE' ";
            sql += " and h.jvh_gstin is null ";
            sql += " and h.jvh_gst = 'N' ";
            sql += " group by jvh_docno,jvh_date,st.param_code,st.param_name";

            sql += " union all";

            sql += " select jvh_docno ";
            sql += " ,jvh_date";
            sql += " ,max(jvh_net_amt) as inv_amt";
            sql += " ,sum(jv_credit) as taxable_amt";
            sql += " ,jv_gst_rate";
            sql += " ,sum(jv_gst_amt) as jv_gst_amt";
            sql += " ,st.param_code||'-'||initcap(st.param_name) as jvh_state_name";
            sql += " ,'' as ecomgstn";
            sql += " ,0 as cess";
            sql += " from ledgerh h";
            sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id and jv_credit>0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += " left join param st on h.jvh_state_id = st.param_pkid";
            sql += " where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += "   and h.rec_branch_code = '{BRCODE}'";
            sql += " and h.jvh_type in ('IN') ";
            sql += " and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += " and h.jvh_gst_type  = 'INTER-STATE' ";
            sql += " and h.jvh_gstin is null ";
            sql += " and h.jvh_gst = 'Y' ";
            sql += " and jv_gst_rate > 0 ";
            sql += " group by jvh_docno,jvh_date,jv_gst_rate,st.param_code,st.param_name";

            sql += " )a where inv_amt > 250000";
            sql += "  order  by jvh_docno,jv_gst_rate";

            sql = sql.Replace("{COMPCODE}", company_code);
            sql = sql.Replace("{BRCODE}", branch_code);
            sql = sql.Replace("{FDATE}", from_date);
            sql = sql.Replace("{EDATE}", to_date);

            Dt_B2CL = new DataTable();
            Dt_B2CL = Con_Oracle.ExecuteQuery(sql);
         
            
            /*
              sql = "";
            sql += " select 'OE' as stype, jvh_state_name, jv_gst_rate, sum(jv_net_total) as jv_net_total,sum(jv_taxable_amt) as jv_taxable_amt, 0 as cess, '' as ecomgstin ";
            sql += " from ( ";
            sql += " select jvh_gst_type  ";
            sql += " ,sum(jv_net_total) as jv_net_total";
            sql += " ,sum(jv_taxable_amt) as jv_taxable_amt";
            sql += " ,jv_gst_rate  ";
            sql += " ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
            sql += " ,'' as ecomgstn";
            sql += " ,0 as cess";
            sql += " from ledgerh h";
            sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += " left join param st on h.jvh_state_id = st.param_pkid";
            sql += " where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += "  and h.rec_branch_code = '{BRCODE}'";
            sql += " and h.jvh_type in ('IN') ";
            sql += " and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += " and h.jvh_gstin is null ";
            sql += "  group by jvh_docno,jvh_date ,jvh_gst_type ,jv_gst_rate, st.param_code, st.param_name";
            sql += " ) a  ";
            sql += " where jvh_gst_type ='INTRA-STATE' or ( jvh_gst_type ='INTER-STATE' and jv_net_total <=250000 ) ";
            sql += " group by jvh_state_name,jv_gst_rate ";
            sql += " order by jvh_state_name,jv_gst_rate";

            sql = sql.Replace("{COMPCODE}", company_code);
            sql = sql.Replace("{BRCODE}", branch_code);
            sql = sql.Replace("{FDATE}", from_date);
            sql = sql.Replace("{EDATE}", to_date);
            */



            sql = "";
            sql += " select 'OE' as stype, jvh_state_name, jv_gst_rate, sum(inv_amt) as inv_amt,sum(taxable_amt) as taxable_amt, 0 as cess, '' as ecomgstin ";
            sql += " from ( ";
            sql += " select jvh_gst_type  ";
            sql += " ,max(jvh_net_amt) as inv_amt";
            sql += " ,sum(jv_credit) as taxable_amt";
            sql += " ,0 as jv_gst_rate  ";
            sql += " ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
            sql += " ,'' as ecomgstn";
            sql += " ,0 as cess";
            sql += " from ledgerh h";
            sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += " left join param st on h.jvh_state_id = st.param_pkid";
            sql += " where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += "  and h.rec_branch_code = '{BRCODE}'";
            sql += " and h.jvh_type in ('IN') ";
            sql += " and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += " and h.jvh_gstin is null ";
            sql += " and h.jvh_gst = 'N' ";
            sql += "  group by jvh_docno,jvh_date ,jvh_gst_type ,st.param_code, st.param_name";

            sql += " union all";

            sql += " select jvh_gst_type  ";
            sql += " ,max(jvh_net_amt) as inv_amt";
            sql += " ,sum(jv_credit) as taxable_amt";
            sql += " ,jv_gst_rate  ";
            sql += " ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
            sql += " ,'' as ecomgstn";
            sql += " ,0 as cess";
            sql += " from ledgerh h";
            sql += " inner join ledgert t on h.jvh_pkid = t.jv_parent_id and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += " left join param st on h.jvh_state_id = st.param_pkid";
            sql += " where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += "  and h.rec_branch_code = '{BRCODE}'";
            sql += " and h.jvh_type in ('IN') ";
            sql += " and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += " and h.jvh_gstin is null ";
            sql += " and h.jvh_gst = 'Y' ";
            sql += " and jv_gst_rate > 0 ";
            sql += "  group by jvh_docno,jvh_date ,jvh_gst_type ,jv_gst_rate, st.param_code, st.param_name";
            sql += " ) a  ";
            sql += " where jvh_gst_type ='INTRA-STATE' or ( jvh_gst_type ='INTER-STATE' and inv_amt <=250000 ) ";
            sql += " group by jvh_state_name,jv_gst_rate ";
            sql += " order by jvh_state_name,jv_gst_rate";

            sql = sql.Replace("{COMPCODE}", company_code);
            sql = sql.Replace("{BRCODE}", branch_code);
            sql = sql.Replace("{FDATE}", from_date);
            sql = sql.Replace("{EDATE}", to_date);

            Dt_B2SM = new DataTable();
            Dt_B2SM = Con_Oracle.ExecuteQuery(sql);

            //LovService lov = new LovService();
            //DataRow lovRow_Gstin = lov.getSettings(branch_code, "GSTIN");
            string BR_GSTIN = "";
            //if (lovRow_Gstin != null)
            //    BR_GSTIN = lovRow_Gstin["name"].ToString();
            /*
            sql = " select jv_sac_code,  '' as jv_sac_name,'" + BR_GSTIN + "' as our_gstin ,";
            sql += " jv_net_total, jv_taxable_amt,jv_cgst_amt, jv_sgst_amt, jv_igst_amt, jv_gst_amt,";
            sql += " 'OTH-OTHERS' as UQC,0 as TOT_QTY,0 as CESS ";
            sql += "  from ( ";
            sql += "  select substr(param_code,1,4) as jv_sac_code, ";
            sql += "  sum(jv_net_total)  as jv_net_total,  ";
            sql += "  sum(jv_taxable_amt) as jv_taxable_amt , ";
            sql += "  sum(jv_cgst_amt) as jv_cgst_amt,  ";
            sql += "  sum(jv_sgst_amt) as jv_sgst_amt,  ";
            sql += "  sum(jv_igst_amt) as jv_igst_amt,  ";
            sql += "  sum(jv_gst_amt) as jv_gst_amt   ";
            sql += "  from ledgerh h";
            sql += "  inner join ledgert t on h.jvh_pkid = t.jv_parent_id  and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += "  left join param sac on t.jv_sac_id = sac.param_pkid";
            sql += " where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += " and h.rec_branch_code = '{BRCODE}'";
            sql += " and h.jvh_type in ('IN') ";
            sql += " and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += " group by substr(param_code,1,4) ";
            sql += " ) a ";
            sql += "  order by  jv_sac_code ";

         
            */

            sql = " select jv_sac_code,  '' as jv_sac_name,'" + BR_GSTIN + "' as our_gstin ,";
            sql += " inv_amt, taxable_amt,jv_cgst_amt, jv_sgst_amt, jv_igst_amt, jv_gst_amt,";
            sql += " 'OTH-OTHERS' as UQC,0 as TOT_QTY,0 as CESS ";
            sql += "  from ( ";
            sql += "  select substr(param_code,1,4) as jv_sac_code, ";
            sql += "  sum(jv_net_total)  as inv_amt,  ";
            sql += "  sum(jv_credit) as taxable_amt , ";
            sql += "  sum(jv_cgst_amt) as jv_cgst_amt,  ";
            sql += "  sum(jv_sgst_amt) as jv_sgst_amt,  ";
            sql += "  sum(jv_igst_amt) as jv_igst_amt,  ";
            sql += "  sum(jv_gst_amt) as jv_gst_amt   ";
            sql += "  from ledgerh h";
            sql += "  inner join ledgert t on h.jvh_pkid = t.jv_parent_id  and jv_credit >0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += "  left join param sac on t.jv_sac_id = sac.param_pkid";
            sql += " where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += " and h.rec_branch_code = '{BRCODE}'";
            sql += " and h.jvh_type in ('IN','IN-ES') ";
            sql += " and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += " group by substr(param_code,1,4) ";
            sql += " ) a ";
            sql += "  order by  jv_sac_code ";

            sql = sql.Replace("{COMPCODE}", company_code);
            sql = sql.Replace("{BRCODE}", branch_code);
            sql = sql.Replace("{FDATE}", from_date);
            sql = sql.Replace("{EDATE}", to_date);

            Dt_HSN = new DataTable();
            Dt_HSN = Con_Oracle.ExecuteQuery(sql);

            sql = " select ";
            sql += "   case when nvl(jv_gst_rate,0)>0 then 'WPAY' else 'WOPAY' end as export_type";
            sql += "  ,jvh_docno ";
            sql += "  ,jvh_date";
            sql += "  ,max(jvh_net_amt) as inv_amt";
            sql += "  ,max(pol.param_code) as port_code";
            sql += "  ,'' as sb_no";
            sql += "  ,'' as sb_date";
            sql += "  ,'' as applicable_tax_rate";
            sql += "  ,jv_gst_rate as gst_rate ";
            sql += "  ,sum(jv_credit) as taxable_amt";
            sql += "  ,0 as cess_amt";
            sql += "  from ledgerh h";
            sql += "  inner join ledgert t on h.jvh_pkid = t.jv_parent_id and jv_credit>0 and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += "  left join hblm hbl on h.jvh_cc_id = hbl.hbl_pkid ";
            sql += "  left join param pol on hbl.hbl_pol_id = pol.param_pkid";
            sql += "  where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += " and h.rec_branch_code = '{BRCODE}'";
            sql += "  and h.jvh_type in ('IN-ES')";
            sql += "  and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += "  group by jvh_docno,jvh_date,jv_gst_rate";
            sql += "  order by jvh_docno";

            sql = sql.Replace("{COMPCODE}", company_code);
            sql = sql.Replace("{BRCODE}", branch_code);
            sql = sql.Replace("{FDATE}", from_date);
            sql = sql.Replace("{EDATE}", to_date);

            Dt_EXP = new DataTable();
            Dt_EXP = Con_Oracle.ExecuteQuery(sql);
            DocInvExpCountDic = new Dictionary<int, string>();
            foreach (DataRow Dr in Dt_EXP.Rows)
            {
                if (!DocInvExpCountDic.ContainsValue(Dr["jvh_docno"].ToString().Trim()))
                    DocInvExpCountDic.Add(DocInvExpCountDic.Count, Dr["jvh_docno"].ToString().Trim().Replace("-",""));
            }

            sql = " select jvh_docno  ";
            sql += "   ,jvh_date ";
            sql += "   ,jvh_gstin,jv_gst_rate, 0 as cess  ";
            sql += "   ,case when jvh_type ='DN' then 'D' else case when jvh_type = 'CN' then 'C' else '' end end  as Doc_type";
            sql += "   ,max(jvh_vrno) as jvh_vrno";
            sql += "   ,max(jvh_org_invno) as jvh_org_invno";
            sql += "   ,max(jvh_org_invdt) as jvh_org_invdt";
            sql += "   ,sum(jv_net_total) as jv_net_total";//inv_amt 
            sql += "   ,sum(jv_taxable_amt) as jv_taxable_amt";
            sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
            sql += "   ,max(party.acc_name) as jvh_party_name ";
            sql += "   from ledgerh h";
            sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id  and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += "   left join param st on h.jvh_state_id = st.param_pkid";
            sql += "   left join acctm party on h.jvh_acc_id = party.acc_pkid";
            sql += "   where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += "   and h.rec_branch_code = '{BRCODE}'";
            sql += "   and h.jvh_type in ('DN','CN') ";
            sql += "   and h.jvh_gstin is not null ";
            sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += "   group by jvh_docno,jvh_date,jvh_type,jv_gst_rate, jvh_gstin,st.param_code, st.param_name";
            sql += "   order by jvh_docno ";

            sql = sql.Replace("{COMPCODE}", company_code);
            sql = sql.Replace("{BRCODE}", branch_code);
            sql = sql.Replace("{FDATE}", from_date);
            sql = sql.Replace("{EDATE}", to_date);

            Dt_CDNR = new DataTable();
            Dt_CDNR = Con_Oracle.ExecuteQuery(sql);

            sql = " select * from (";
            sql += " select '' as ur_type, jvh_docno  ";
            sql += "   ,jvh_date ";
            sql += "   ,jv_gst_rate, 0 as cess  ";
            sql += "   ,case when jvh_type ='DN' then 'D' else case when jvh_type = 'CN' then 'C' else '' end end  as Doc_type";
            sql += "   ,max(jvh_vrno) as jvh_vrno";
            sql += "   ,max(jvh_org_invno) as jvh_org_invno";
            sql += "   ,max(jvh_org_invdt) as jvh_org_invdt";
            sql += "   ,sum(jv_net_total) as jv_net_total";//inv_amt 
            sql += "   ,sum(jv_taxable_amt) as jv_taxable_amt";
            sql += "   ,st.param_code || '-' || initcap(st.param_name) as jvh_state_name ";
            sql += "   from ledgerh h";
            sql += "   inner join ledgert t on h.jvh_pkid = t.jv_parent_id  and nvl(jv_row_type,'JV') not in('HEADER','GST')";
            sql += "   left join param st on h.jvh_state_id = st.param_pkid";
            sql += "   where h.rec_company_code = '{COMPCODE}'";
            if (this.branch_code != "")
                sql += "   and h.rec_branch_code = '{BRCODE}'";
            sql += "   and h.jvh_type in ('DN','CN') ";
            sql += "   and h.jvh_gstin is null ";
            sql += "   and h.jvh_gst_type  = 'INTER-STATE' ";
            sql += "   and h.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
            sql += "   group by jvh_docno,jvh_date,jvh_type,jv_gst_rate,st.param_code, st.param_name";
            sql += "  )a where jv_net_total > 250000";
            sql += "   order by jvh_docno ";

            sql = sql.Replace("{COMPCODE}", company_code);
            sql = sql.Replace("{BRCODE}", branch_code);
            sql = sql.Replace("{FDATE}", from_date);
            sql = sql.Replace("{EDATE}", to_date);

            Dt_CDNUR = new DataTable();
            Dt_CDNUR = Con_Oracle.ExecuteQuery(sql);

            DocCNCountDic = new Dictionary<int, string>();
            DocDNCountDic=new Dictionary<int, string>();
            int sKey = 0;
            foreach (DataRow dr in Dt_CDNR.Rows)
            {
                sKey = Lib.Conv2Integer(dr["jvh_vrno"].ToString());
                if (dr["Doc_type"].ToString() == "C")
                {
                    if (!DocCNCountDic.ContainsKey(sKey))
                        DocCNCountDic.Add(sKey, dr["jvh_docno"].ToString().Trim());
                }
                if (dr["Doc_type"].ToString() == "D")
                {
                    if (!DocDNCountDic.ContainsKey(sKey))
                        DocDNCountDic.Add(sKey, dr["jvh_docno"].ToString().Trim());
                }
            }
            foreach (DataRow dr in Dt_CDNUR.Rows)
            {
                sKey = Lib.Conv2Integer(dr["jvh_vrno"].ToString());
                if (dr["Doc_type"].ToString() == "C")
                {
                    if (!DocCNCountDic.ContainsKey(sKey))
                        DocCNCountDic.Add(sKey, dr["jvh_docno"].ToString().Trim());
                }
                if (dr["Doc_type"].ToString() == "D")
                {
                    if (!DocDNCountDic.ContainsKey(sKey))
                        DocDNCountDic.Add(sKey, dr["jvh_docno"].ToString().Trim());
                }
            }

            Con_Oracle.CloseConnection();

            GenerateMsg = " Files Generated : ";
            printb2bCSV();
            printb2clCSV();
            printb2csCSV();
            printhsnCSV();
            printexpCSV();
            printdocsCSV();
            printcdnrCSV();
            printcdnurCSV();

            Dt_B2CL.Rows.Clear();
            Dt_B2SM.Rows.Clear();
            Dt_HSN.Rows.Clear();
            Dt_EXP.Rows.Clear();
            Dt_CDNR.Rows.Clear();
            Dt_CDNUR.Rows.Clear();
        }

        private void printb2bCSV()
        {
           
            try
            {
                bool bOk = false;
                StringBuilder sb = new StringBuilder();

                sb.Append("GSTIN/UIN of Recipient"); sb.Append(",");
                sb.Append("Invoice Number"); sb.Append(",");
                sb.Append("Invoice date"); sb.Append(",");
                sb.Append("Invoice Value"); sb.Append(",");
                sb.Append("Place Of Supply"); sb.Append(",");
                sb.Append("Reverse Charge"); sb.Append(",");
                sb.Append("Invoice Type"); sb.Append(",");
                sb.Append("E-Commerce GSTIN"); sb.Append(",");
                sb.Append("Rate"); sb.Append(",");
                sb.Append("Taxable Value"); sb.Append(",");
                sb.Append("Cess Amount");
                foreach (GstReport Rec in mList)
                {
                    if (Rec.row_type== "DETAIL" && Rec.jvh_gstin.ToString() != "")
                    {
                        bOk = true;
                        sb.AppendLine();

                        sb.Append(Rec.jvh_gstin); sb.Append(",");
                        sb.Append(Rec.jvh_docno); sb.Append(",");
                        sb.Append(Rec.jvh_date_gstr1); sb.Append(",");
                        sb.Append(Rec.inv_amt.ToString()); sb.Append(",");
                        sb.Append(Rec.jvh_state_name.ToString()); sb.Append(",");
                        sb.Append(Rec.rc.ToString()); sb.Append(",");
                        sb.Append(Rec.jvh_invoice_type.ToString()); sb.Append(",");
                        sb.Append(Rec.ecomgstn.ToString()); sb.Append(",");
                        sb.Append(Rec.jv_gst_rate.ToString()); sb.Append(",");
                        sb.Append(Rec.taxable_amt.ToString()); sb.Append(",");
                        sb.Append(Rec.cess.ToString());
                    }
                }
                WriteCsvFile("b2b", sb, bOk);
            }
            catch (Exception Ex)
            {
                GenerateMsg += " | b2b Error :" + Ex.Message.ToString();
            }
        }
        private void printb2clCSV()
        {
            DateTime bDate;
            try
            {
                bool bOk = false;
                StringBuilder sb = new StringBuilder();
                sb.Append("Invoice Number"); sb.Append(",");
                sb.Append("Invoice date"); sb.Append(",");
                sb.Append("Invoice Value"); sb.Append(",");
                sb.Append("Place Of Supply"); sb.Append(",");
                sb.Append("Rate"); sb.Append(",");
                sb.Append("Taxable Value"); sb.Append(",");
                sb.Append("Cess Amount"); sb.Append(",");
                sb.Append("E-Commerce GSTIN");
                foreach (DataRow dr in Dt_B2CL.Rows)
                {
                    bOk = true;
                    sb.AppendLine();
                    sb.Append(dr["jvh_docno"].ToString()); sb.Append(",");
                    if (!dr["jvh_date"].Equals(DBNull.Value))
                    {
                        bDate = (DateTime)dr["jvh_date"];
                        sb.Append(bDate.ToString("dd-MMM-yyyy")); sb.Append(",");
                    }
                    else
                    {
                        sb.Append(""); sb.Append(",");
                    }

                    sb.Append(dr["inv_amt"].ToString()); sb.Append(",");
                    if (dr["jvh_state_name"].ToString().Trim() == "-")
                        dr["jvh_state_name"] = "";
                    sb.Append(dr["jvh_state_name"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_gst_rate"].ToString()); sb.Append(",");
                    sb.Append(dr["taxable_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["cess"].ToString()); sb.Append(",");
                    sb.Append(dr["ecomgstn"].ToString());

                }
                WriteCsvFile("b2cl", sb, bOk);
            }
            catch (Exception Ex)
            {
                GenerateMsg += " | b2cl Error :" + Ex.Message.ToString();
            }
        }
        private void printb2csCSV()
        {
            try
            {
                bool bOk = false;
                StringBuilder sb = new StringBuilder();

                sb.Append("Type"); sb.Append(",");
                sb.Append("Place Of Supply"); sb.Append(",");
                sb.Append("Rate"); sb.Append(",");
                sb.Append("Taxable Value"); sb.Append(",");
                sb.Append("Cess Amount"); sb.Append(",");
                sb.Append("E-Commerce GSTIN");

                foreach (DataRow dr in Dt_B2SM.Rows)
                {
                    bOk = true;
                    sb.AppendLine();
                    sb.Append(dr["stype"].ToString()); sb.Append(",");
                    if (dr["jvh_state_name"].ToString().Trim() == "-")
                        dr["jvh_state_name"] = "";
                    sb.Append(dr["jvh_state_name"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_gst_rate"].ToString()); sb.Append(",");
                    sb.Append(dr["taxable_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["cess"].ToString()); sb.Append(",");
                    sb.Append(dr["ecomgstin"].ToString());
                }

                WriteCsvFile("b2cs", sb, bOk);
            }
            catch (Exception Ex)
            {
                GenerateMsg += " | b2cs Error :" + Ex.Message.ToString();
            }
        }
        private void printhsnCSV()
        {
            try
            {
                bool bOk = false;
                StringBuilder sb = new StringBuilder();

                sb.Append("HSN"); sb.Append(",");
                sb.Append("Description"); sb.Append(",");
                sb.Append("UQC"); sb.Append(",");
                sb.Append("Total Quantity"); sb.Append(",");
                sb.Append("Total Value"); sb.Append(",");
                sb.Append("Taxable Value"); sb.Append(",");
                sb.Append("Integrated Tax Amount"); sb.Append(",");
                sb.Append("Central Tax Amount"); sb.Append(",");
                sb.Append("State/UT Tax Amount"); sb.Append(",");
                sb.Append("Cess Amount");
                foreach (DataRow dr in Dt_HSN.Rows)
                {
                    bOk = true;
                    sb.AppendLine();
                    sb.Append(dr["jv_sac_code"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_sac_name"].ToString()); sb.Append(",");
                    sb.Append(dr["uqc"].ToString()); sb.Append(",");
                    sb.Append(dr["tot_qty"].ToString()); sb.Append(",");
                    sb.Append(dr["inv_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["taxable_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_igst_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_cgst_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_sgst_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["cess"].ToString());
                }

                WriteCsvFile("hsn", sb, bOk);
            }
            catch (Exception Ex)
            {
                GenerateMsg += " | hsn Error :" + Ex.Message.ToString();
            }
        }

        private void printexpCSV()
        {
            DateTime bDate;
            try
            {
                bool bOk = false;
                StringBuilder sb = new StringBuilder();

                sb.Append("Export Type"); sb.Append(",");
                sb.Append("Invoice Number"); sb.Append(",");
                sb.Append("Invoice date"); sb.Append(",");
                sb.Append("Invoice Value"); sb.Append(",");
                sb.Append("Port Code"); sb.Append(",");
                sb.Append("Shipping Bill Number"); sb.Append(",");
                sb.Append("Shipping Bill Date"); sb.Append(",");
                sb.Append("Rate"); sb.Append(",");
                sb.Append("Applicable % ofTax Rate"); sb.Append(",");
                sb.Append("Taxable Value"); sb.Append(",");
                sb.Append("Cess Amount");
                foreach (DataRow dr in Dt_EXP.Rows)
                {
                    bOk = true;
                    sb.AppendLine();
                    sb.Append(dr["export_type"].ToString()); sb.Append(",");
                    sb.Append(dr["jvh_docno"].ToString().Replace("-","")); sb.Append(",");//Replace Invoice Hypen Here
                    if (!dr["jvh_date"].Equals(DBNull.Value))
                    {
                        bDate = (DateTime)dr["jvh_date"];
                        sb.Append(bDate.ToString("dd-MMM-yyyy")); sb.Append(",");
                    }
                    else
                    {
                        sb.Append(""); sb.Append(",");
                    }
                    sb.Append(dr["inv_amt"].ToString()); sb.Append(",");
                    //   sb.Append(dr["port_code"].ToString()); sb.Append(",");
                    sb.Append(""); sb.Append(",");
                    sb.Append(dr["sb_no"].ToString()); sb.Append(",");
                    if (!dr["sb_date"].Equals(DBNull.Value))
                    {
                        bDate = (DateTime)dr["sb_date"];
                        sb.Append(bDate.ToString("dd-MMM-yyyy")); sb.Append(",");
                    }
                    else
                    {
                        sb.Append(""); sb.Append(",");
                    }
                    sb.Append(Lib.NumericFormat(Lib.Conv2Decimal(dr["gst_rate"].ToString()).ToString(),2)); sb.Append(",");
                    sb.Append(Lib.NumericFormat(Lib.Conv2Decimal(dr["applicable_tax_rate"].ToString()).ToString(),2)); sb.Append(",");
                    sb.Append(dr["taxable_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["cess_amt"].ToString());
                }

                WriteCsvFile("exp", sb, bOk);
            }
            catch (Exception Ex)
            {
                GenerateMsg += " | exp Error :" + Ex.Message.ToString();
            }
        }
        private void printdocsCSV()
        {
            try
            {
                bool bOk = false;
                StringBuilder sb = new StringBuilder();

                sb.Append("Nature  of Document"); sb.Append(",");
                sb.Append("Sr. No. From"); sb.Append(",");
                sb.Append("Sr. No. To"); sb.Append(",");
                sb.Append("Total Number"); sb.Append(",");
                sb.Append("Cancelled");

                if (DocInvCountDic.Count > 0)
                {
                    bOk = true;
                    sb.AppendLine();
                    sb.Append("Invoice for outward supply"); sb.Append(",");
                    sb.Append(DocInvCountDic[0]); sb.Append(",");
                    sb.Append(DocInvCountDic[DocInvCountDic.Count - 1]); sb.Append(",");
                    sb.Append(DocInvCountDic.Count.ToString()); sb.Append(",");
                    sb.Append("");
                }
                if (DocDNCountDic.Count > 0)
                {
                    bOk = true;
                    sb.AppendLine();
                    sb.Append("Debit Note"); sb.Append(",");
                    sb.Append(DocDNCountDic[DocDNCountDic.Keys.Min()]); sb.Append(",");
                    sb.Append(DocDNCountDic[DocDNCountDic.Keys.Max()]); sb.Append(",");
                    sb.Append(DocDNCountDic.Count.ToString()); sb.Append(",");
                    sb.Append("");
                }
                if (DocCNCountDic.Count > 0)
                {
                    bOk = true;
                    sb.AppendLine();
                    sb.Append("Credit Note"); sb.Append(",");
                    sb.Append(DocCNCountDic[DocCNCountDic.Keys.Min()]); sb.Append(",");
                    sb.Append(DocCNCountDic[DocCNCountDic.Keys.Max()]); sb.Append(",");
                    sb.Append(DocCNCountDic.Count.ToString()); sb.Append(",");
                    sb.Append("");
                }
                if (DocInvExpCountDic.Count > 0)
                {
                    bOk = true;
                    sb.AppendLine();
                    sb.Append("Invoice for outward supply"); sb.Append(",");
                    sb.Append(DocInvExpCountDic[0]); sb.Append(",");
                    sb.Append(DocInvExpCountDic[DocInvExpCountDic.Count - 1]); sb.Append(",");
                    sb.Append(DocInvExpCountDic.Count.ToString()); sb.Append(",");
                    sb.Append("");
                }
                WriteCsvFile("docs", sb, bOk);
            }
            catch (Exception Ex)
            {
                GenerateMsg += " | docs Error :" + Ex.Message.ToString();
            }
        }
        private void printcdnrCSV()
        {
            DateTime bDate;
            try
            {
                bool bOk = false;
                StringBuilder sb = new StringBuilder();

                sb.Append("GSTIN/UIN of Recipient"); sb.Append(",");
                sb.Append("Receiver Name"); sb.Append(",");
                sb.Append("Invoice/Advance Receipt Number"); sb.Append(",");
                sb.Append("Invoice/Advance Receipt date"); sb.Append(",");
                sb.Append("Note/Refund Voucher Number"); sb.Append(",");
                sb.Append("Note/Refund Voucher date"); sb.Append(",");
                sb.Append("Document Type"); sb.Append(",");
                sb.Append("Place Of Supply"); sb.Append(",");
                sb.Append("Note/Refund Voucher Value"); sb.Append(",");
                sb.Append("Rate"); sb.Append(",");
                sb.Append("Applicable % of Tax Rate"); sb.Append(",");
                sb.Append("Taxable Value"); sb.Append(",");
                sb.Append("Cess Amount"); sb.Append(",");
                sb.Append("Pre GST"); 

                foreach (DataRow dr in Dt_CDNR.Rows)
                {
                     
                        bOk = true;
                        sb.AppendLine();

                    sb.Append(dr["jvh_gstin"].ToString()); sb.Append(",");
                    sb.Append(dr["jvh_party_name"].ToString()); sb.Append(",");
                    sb.Append(dr["jvh_org_invno"].ToString()); sb.Append(",");
                    if (!dr["jvh_org_invdt"].Equals(DBNull.Value))
                    {
                        bDate = (DateTime)dr["jvh_org_invdt"];
                        sb.Append(bDate.ToString("dd-MMM-yyyy")); sb.Append(",");
                    }
                    else
                    {
                        sb.Append(""); sb.Append(",");
                    }
                    sb.Append(dr["jvh_docno"].ToString()); sb.Append(",");
                    if (!dr["jvh_date"].Equals(DBNull.Value))
                    {
                        bDate = (DateTime)dr["jvh_date"];
                        sb.Append(bDate.ToString("dd-MMM-yyyy")); sb.Append(",");
                    }
                    else
                    {
                        sb.Append(""); sb.Append(",");
                    }
                    
                    sb.Append(dr["doc_type"].ToString()); sb.Append(",");
                    sb.Append(dr["jvh_state_name"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_net_total"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_gst_rate"].ToString()); sb.Append(",");
                    sb.Append(""); sb.Append(",");
                    sb.Append(dr["jv_taxable_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["cess"].ToString()); sb.Append(",");
                    sb.Append("N");  
                }
                WriteCsvFile("cdnr", sb, bOk);
            }
            catch (Exception Ex)
            {
                GenerateMsg += " | cdnr Error :" + Ex.Message.ToString();
            }
        }
        private void printcdnurCSV()
        {
            DateTime bDate;
            try
            {
                bool bOk = false;
                StringBuilder sb = new StringBuilder();

                sb.Append("UR Type"); sb.Append(",");
                sb.Append("Note/Refund Voucher Number"); sb.Append(",");
                sb.Append("Note/Refund Voucher date"); sb.Append(",");
                sb.Append("Document Type"); sb.Append(",");
                sb.Append("Invoice/Advance Receipt Number"); sb.Append(",");
                sb.Append("Invoice/Advance Receipt date"); sb.Append(",");
                sb.Append("Place Of Supply"); sb.Append(",");
                sb.Append("Note/Refund Voucher Value"); sb.Append(",");
                sb.Append("Rate"); sb.Append(",");
                sb.Append("Applicable % of Tax Rate"); sb.Append(",");
                sb.Append("Taxable Value"); sb.Append(",");
                sb.Append("Cess Amount"); sb.Append(",");
                sb.Append("Pre GST");

                foreach (DataRow dr in Dt_CDNUR.Rows)
                {

                    bOk = true;
                    sb.AppendLine();
                    sb.Append(dr["ur_type"].ToString()); sb.Append(",");
                    sb.Append(dr["jvh_docno"].ToString()); sb.Append(",");
                    if (!dr["jvh_date"].Equals(DBNull.Value))
                    {
                        bDate = (DateTime)dr["jvh_date"];
                        sb.Append(bDate.ToString("dd-MMM-yyyy")); sb.Append(",");
                    }
                    else
                    {
                        sb.Append(""); sb.Append(",");
                    }
                    sb.Append(dr["doc_type"].ToString()); sb.Append(",");
                    sb.Append(dr["jvh_org_invno"].ToString()); sb.Append(",");
                    if (!dr["jvh_org_invdt"].Equals(DBNull.Value))
                    {
                        bDate = (DateTime)dr["jvh_org_invdt"];
                        sb.Append(bDate.ToString("dd-MMM-yyyy")); sb.Append(",");
                    }
                    else
                    {
                        sb.Append(""); sb.Append(",");
                    }                   
                    sb.Append(dr["jvh_state_name"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_net_total"].ToString()); sb.Append(",");
                    sb.Append(dr["jv_gst_rate"].ToString()); sb.Append(",");
                    sb.Append(""); sb.Append(",");
                    sb.Append(dr["jv_taxable_amt"].ToString()); sb.Append(",");
                    sb.Append(dr["cess"].ToString()); sb.Append(",");
                    sb.Append("N");
                }
                WriteCsvFile("cdnur", sb, bOk);
            }
            catch (Exception Ex)
            {
                GenerateMsg += " | cdnur Error :" + Ex.Message.ToString();
            }
        }
        private void WriteCsvFile(string sFileName, StringBuilder StrBld, bool CanSave)
        {
            DateTime dt_from = DateTime.Parse(from_date);
            string fName = report_folder + "\\GST\\" + branch_code;
            fName += "\\" + dt_from.ToString("MMMM").ToString().ToUpper();

            if (!System.IO.Directory.Exists(fName))
                System.IO.Directory.CreateDirectory(fName);

            fName += "\\" + sFileName + ".csv";

            if (System.IO.File.Exists(fName))
                System.IO.File.Delete(fName);

            if (CanSave)
            {
                System.IO.File.AppendAllText(fName, StrBld.ToString());
                GenerateMsg += " | " + sFileName.ToUpper();
            }
        }
    }
}

