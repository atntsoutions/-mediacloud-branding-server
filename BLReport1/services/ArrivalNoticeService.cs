using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLReport1
{
    public class ArrivalNoticeService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        DataTable Dt_Tracking = new DataTable(); 

        List<ArrivalNotice> mList = new List<ArrivalNotice>();
        ArrivalNotice mrow;
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string ErrorMessage = "";
        string MailSub = "";
        string MailMsg = "";
        string MailTo_ids = "";
        string shipper_id = "";
        int priordays = 0;
        string sWhere = "";
        long page_count =0;
        long page_current = 0;
        long page_rows =0;
        long page_rowcount =0;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<ArrivalNotice>();
            ErrorMessage = "";
            string pkid = "";
            string type = "";
            try
            {
                type = SearchData["type"].ToString();

                if (type == "MAIL")
                {
                    pkid = SearchData["pkid"].ToString();
                    MailMsg = GetEmail(pkid);
                }
                else
                {
                    page_count = (long)SearchData["page_count"];
                    page_current = (long)SearchData["page_current"];
                    page_rows = (long)SearchData["page_rows"];
                    page_rowcount = (long)SearchData["page_rowcount"];
                    long startrow = 0;
                    long endrow = 0;

                    company_code = SearchData["company_code"].ToString();
                    branch_code = SearchData["branch_code"].ToString();
                    year_code = SearchData["year_code"].ToString();
                    if (SearchData.ContainsKey("shipper_id"))
                        shipper_id = SearchData["shipper_id"].ToString();
                    if (SearchData.ContainsKey("priordays"))
                    {
                        if (SearchData["priordays"] != null)
                            priordays = Lib.Conv2Integer(SearchData["priordays"].ToString());
                    }

                    sWhere = "";
                    sWhere = " where m.rec_company_code = '{COMPCODE}'";
                    sWhere += " and m.rec_branch_code = '{BRCODE}'";
                    sWhere += " and m.hbl_type = 'MBL-SE'";
                    if (shipper_id.Length > 0)
                        sWhere += " and h.hbl_exp_id = '" + shipper_id + "' ";
                    if (priordays != 0)
                        sWhere += " and TRUNC(m.hbl_pod_eta-sysdate) >= " + priordays.ToString();

                    // sWhere += "  and  to_date(to_char(m.hbl_pod_eta,'DD-MON-YYYY'),'DD-MON-YYYY') >= to_char(sysdate + 10,'DD-MON-YYYY')";

                    sWhere = sWhere.Replace("{COMPCODE}", company_code);
                    sWhere = sWhere.Replace("{BRCODE}", branch_code);


                    Con_Oracle = new DBConnection();

                    if (type == "NEW")
                    {
                        sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total ";
                        sql += " from hblm m";
                        sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id ";
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

                    sql = "";
                    sql += " select * from ( ";
                    sql += " select m.hbl_pkid as mbl_pkid,m.hbl_no as mbl_bkno,m.hbl_bl_no as mbl_no, m.hbl_book_no as mbl_book_no , nvl(h.hbl_ar_notice,'N') as hbl_ar_notice ,";
                    sql += " TRUNC(m.hbl_pod_eta-sysdate) as mbl_eta_days,m.hbl_date as mbl_book_date,m.hbl_pol_etd as mbl_pol_etd,m.hbl_pod_eta as mbl_pod_eta,";
                    sql += " h.hbl_pkid,h.hbl_bl_no as hbl_blno,h.hbl_book_cntr as hbl_cntrs,";
                    sql += " pod.param_code as mbl_pod_code,pod.param_name as mbl_pod_name,";
                    sql += " cnge.cust_code as hbl_imp_code,cnge.cust_name as hbl_imp_name,h.hbl_terms,";
                    sql += " h.hbl_pkg as hbl_packages,h.hbl_grwt as hbl_grWeight,h.hbl_cbm as hbl_volume, ";
                    sql += " carr.param_name as mbl_carrier_name,";
                    sql += " row_number() over(order by TRUNC(m.hbl_pod_eta-sysdate) desc) rn";
                    sql += " from hblm m";
                    sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id and h.hbl_type='HBL-SE' ";
                    sql += " inner join param pod on m.hbl_pod_id = pod.param_pkid";
                    sql += " inner join customerm cnge on h.hbl_imp_id = cnge.cust_pkid";
                    sql += " left join param carr on m.hbl_carrier_id = carr.param_pkid";
                    sql += sWhere;
                    sql += ") a where rn between {startrow} and {endrow}";

                    // sql += " order by TRUNC(m.hbl_pod_eta-sysdate) desc ";

                    sql = sql.Replace("{startrow}", startrow.ToString());
                    sql = sql.Replace("{endrow}", endrow.ToString());

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new ArrivalNotice();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.mbl_pkid = Dr["mbl_pkid"].ToString();
                        mrow.mbl_slno = Dr["mbl_bkno"].ToString();
                        mrow.mbl_eta_days = Lib.Conv2Integer(Dr["mbl_eta_days"].ToString());
                        mrow.mbl_book_no = Dr["mbl_book_no"].ToString();
                        mrow.mbl_book_date = Lib.DatetoStringDisplayformat(Dr["mbl_book_date"]);
                        mrow.hbl_blno = Dr["hbl_blno"].ToString();
                        mrow.mbl_pod_code = Dr["mbl_pod_code"].ToString();
                        mrow.mbl_pod_name = Dr["mbl_pod_name"].ToString();
                        mrow.hbl_pkid = Dr["hbl_pkid"].ToString();
                        mrow.hbl_imp_code = Dr["hbl_imp_code"].ToString();
                        mrow.hbl_imp_name = Dr["hbl_imp_name"].ToString();
                        mrow.mbl_carrier_name = Dr["mbl_carrier_name"].ToString();
                        //mrow.hbl_terms = Dr["hbl_terms"].ToString();
                        mrow.hbl_ar_notice = Dr["hbl_ar_notice"].ToString();
                        //mrow.hbl_packages = Lib.Conv2Decimal(Dr["hbl_packages"].ToString());
                        //mrow.hbl_grweight = Lib.Conv2Decimal(Dr["hbl_grweight"].ToString());
                        //mrow.hbl_volume = Lib.Conv2Decimal(Dr["hbl_volume"].ToString());
                        //mrow.mbl_pol_etd = Lib.DatetoStringDisplayformat(Dr["mbl_pol_etd"]);
                        mrow.mbl_pod_eta = Lib.DatetoStringDisplayformat(Dr["mbl_pod_eta"]);
                        mList.Add(mrow);
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

            RetData.Add("list", mList);
            RetData.Add("mailmessage", MailMsg);
            RetData.Add("mailsubject", MailSub);
            RetData.Add("mailto_ids", MailTo_ids.ToLower());
            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            return RetData;
        }

        private string GetEmail(string HBLID)
        {
            string str = "";
            DataRow Dr;
            String sHtml = "";
            try
            {
                sql  = " select m.hbl_pkid as mbl_pkid,m.hbl_no as mbl_bkno,m.hbl_bl_no as mbl_no, m.hbl_book_no as mbl_book_no , nvl(h.hbl_ar_notice,'N') as hbl_ar_notice ,";
                sql += " TRUNC(m.hbl_pod_eta-sysdate) as mbl_eta_days,m.hbl_date as mbl_book_date,m.hbl_pol_etd as mbl_pol_etd,m.hbl_pod_eta as mbl_pod_eta,";
                sql += " h.hbl_pkid,h.hbl_bl_no as hbl_blno,pol.param_code as mbl_pol_code,pol.param_name as mbl_pol_name,h.hbl_book_cntr as hbl_cntrs,";
                sql += " pod.param_code as mbl_pod_code,pod.param_name as mbl_pod_name,";
                sql += " carr.param_name as mbl_carrier_name,";
                sql += " shpr.cust_code as hbl_exp_code,shpr.cust_name as hbl_exp_name,";
                sql += " shpraddr.add_branch_slno as  hbl_exp_br_no,shpraddr.add_line1 as hbl_exp_addr1,";
                sql += " shpraddr.add_line2  as hbl_exp_addr2 ,shpraddr.add_line3  as hbl_exp_addr3,shpraddr.add_email as hbl_exp_email,";
                sql += " cnge.cust_code as hbl_imp_code,cnge.cust_name as hbl_imp_name,cngeaddr.add_branch_slno as hbl_imp_br_no,";
                sql += " cngeaddr.add_line1 as hbl_imp_addr1, cngeaddr.add_line2 as hbl_imp_addr2, cngeaddr.add_line3 as hbl_imp_addr3,";
                sql += " h.hbl_terms,vsl.param_code as mbl_vessel_code,vsl.param_name as mbl_vessel_name,m.hbl_vessel_no as mbl_vessel_no,";
                sql += " comm.param_name as hbl_commodity,h.hbl_pkg as hbl_packages,h.hbl_grwt as hbl_grWeight,h.hbl_cbm as hbl_volume,grunit.param_code as hbl_grwt_unit_code ";
                sql += " from hblm m";
                sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
                sql += " left join param pol on m.hbl_pol_id = pol.param_pkid";
                sql += " left join param pod on m.hbl_pod_id = pod.param_pkid";
                sql += " left join param carr on m.hbl_carrier_id = carr.param_pkid";
                sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";
                sql += " left join addressm shpraddr on h.hbl_exp_br_id = shpraddr.add_pkid";
                sql += " left join customerm cnge on h.hbl_imp_id = cnge.cust_pkid";
                sql += " left join addressm cngeaddr on h.hbl_imp_br_id = cngeaddr.add_pkid";
                sql += " left join param vsl on m.hbl_vessel_id = vsl.param_pkid";
                sql += " left join param comm on h.hbl_commodity_id = comm.param_pkid";
                sql += " left join param grunit on h.hbl_grwt_unit_id = grunit.param_pkid ";
                sql += " where h.hbl_pkid = '" + HBLID + "' ";
                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                if (Dt_List.Rows.Count > 0)
                {
                    Dr = Dt_List.Rows[0];

                    GetTrackingDetais(Dr["mbl_pkid"].ToString());

                    MailTo_ids = Dr["hbl_exp_email"].ToString();
                    MailSub = "PRIOR ARRIVAL INTIMATION OF CNTR - "+ Dr["hbl_cntrs"].ToString();
                    if (Dr["hbl_blno"].ToString() != "")
                        MailSub += ", HBL - " + Dr["hbl_blno"].ToString();

                    sHtml = " <html>";
                    sHtml += "<body style='font-family=Calibri;'>";
                    sHtml += "<br> ";
                    sHtml += " Dear Sir/ Madam,<br><br />";
                    sHtml += "Please find below shipment expect to arrive final destination<br><br />";

                    sHtml += "<table border=1 cellspacing=0   borderColor=grey  width=800px> ";
                    sHtml += " 	<tr>";
                    sHtml += " 	<td align='center'>";
                    str = "ARRIVAL NOTICE";// REPORT_CAPTION;
                    sHtml += str;
                    sHtml += " </td>";
                    sHtml += " </tr>";

                    sHtml += " 	<tr>";
                    sHtml += " 	<td align='center'>";
                    sHtml += " 	<table  width='800px' BORDER-COLLAPSE: collapse; HEIGHT: 0px; COLOR: black;'  border=0 cellSpacing=0 borderColor=grey>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100' >MBL#</td>  ";
                    sHtml += " 				<td width='700' colspan='3' >:" + Dr["mbl_no"].ToString() + "</td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100' >BOOKING#</td>  ";
                    sHtml += " 				<td width='700' colspan='3' >:" + Dr["mbl_book_no"].ToString() + "</td>  ";
                    sHtml += " 			</tr>";

                    sHtml += "          <tr>";
                    sHtml += " 	            <td width='100' >DATE</td>  ";
                    sHtml += " 				<td width='700' colspan='3' >: " +Lib.DatetoStringDisplayformat(Dr["mbl_book_date"]) + " </td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 	<tr>";
                    sHtml += " 	            <td width='100' >HBL#</td>  ";
                    sHtml += " 				<td width='700' colspan='3' >: " + Dr["hbl_blno"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100' >SHIPPER</td>  ";
                    sHtml += " 				<td width='700' colspan='3' >: " + Dr["hbl_exp_name"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100' >ADDRESS</td>  ";
                    sHtml += " 				<td width='700' colspan='3' >: " + Dr["hbl_exp_addr1"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";

                    if (Dr["hbl_exp_addr2"].ToString() != "")
                    {
                        sHtml += " 			<tr>";
                        sHtml += " 				<td width='100' ></td>  ";
                        sHtml += " 				<td width='700' colspan='3' >: " + Dr["hbl_exp_addr2"].ToString() + " </td>  ";
                        sHtml += " 			</tr>";
                    }

                    if (Dr["hbl_exp_addr3"].ToString() != "")
                    {
                        sHtml += " 			<tr>";
                        sHtml += " 				<td width='100' ></td>  ";
                        sHtml += " 				<td width='700' colspan='3' >: " + Dr["hbl_exp_addr3"].ToString() + " </td>  ";
                        sHtml += " 			</tr>";
                    }

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100' >CONSIGNEE</td>  ";
                    sHtml += " 				<td width='700' colspan='3' >: " + Dr["hbl_imp_name"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100'>ADDRESS</td>  ";
                    sHtml += " 				<td width='700' colspan='3'  >: " + Dr["hbl_imp_addr1"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";

                    if (Dr["hbl_imp_addr2"].ToString() != "")
                    {
                        sHtml += " 			<tr>";
                        sHtml += " 				<td width='100'></td>  ";
                        sHtml += " 				<td width='700' colspan='3'  >: " + Dr["hbl_imp_addr2"].ToString() + " </td>  ";
                        sHtml += " 			</tr>";
                    }

                    if (Dr["hbl_imp_addr3"].ToString() != "")
                    {
                        sHtml += " 			<tr>";
                        sHtml += " 				<td width='100'></td>  ";
                        sHtml += " 				<td width='700' colspan='3'  >: " + Dr["hbl_imp_addr3"].ToString() + " </td>  ";
                        sHtml += " 			</tr>";
                    }

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100'>COMM.INVOICE</td>  ";
                    sHtml += " 				<td width='700' colspan='3'  >: " + GetJobInvoiceDetails(Dr["hbl_pkid"].ToString()) + " </td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100'>CONTAINER</td>  ";
                    sHtml += " 				<td width='700' colspan='3'  >: " + Dr["hbl_cntrs"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100'>COMMODITY</td>  ";
                    sHtml += " 				<td width='700' colspan='3'  >: " + Dr["hbl_commodity"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";


                    sHtml += " 		</table>";

                    sHtml += " 	</br>";

                    sHtml += " 		<table  width='800px' BORDER-COLLAPSE: collapse; HEIGHT: 0px; COLOR: black;'  border=0 cellSpacing=0 borderColor=black >";

                    //sHtml += " 			<tr>";
                    //sHtml += " 				<td width='150' >VESSEL NAME</td>  ";
                    //sHtml += " 				<td width='400' >: " + Dr["mbl_vessel_name"].ToString() + " </td>  ";
                    //sHtml += " 				<td width='100' >VOYAGE NO</td>  ";
                    //sHtml += " 				<td width='150' >: " + Dr["mbl_vessel_voyage"].ToString() + " </td>  ";
                    //sHtml += " 			</tr>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100' >TERMS</td>  ";
                    sHtml += " 				<td width='400' >: " + Dr["hbl_terms"].ToString() + " </td>  ";
                    sHtml += " 				<td width='100' >PKGS</td>  ";
                    sHtml += " 				<td width='200' >: " + Dr["hbl_packages"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 			<tr>";
                    sHtml += " 				<td width='100' >GR.WT</td>  ";
                    sHtml += " 				<td width='400' >: " + Dr["hbl_grweight"].ToString()+" "+Dr["hbl_grwt_unit_code"].ToString() + " </td>  ";
                    sHtml += " 				<td width='100' >CBM</td>  ";
                    sHtml += " 				<td width='200' >: " + Dr["hbl_volume"].ToString() + " </td>  ";
                    sHtml += " 			</tr>";

                    sHtml += " 		</table>";
                    sHtml += " 	</br>";

                    sHtml += " 	<table  WIDTH: 700px; BORDER-COLLAPSE: collapse; HEIGHT: 0px; border=1 cellSpacing=0 borderColor=grey >";
                    sHtml += " 	<tr>";
                    sHtml += " 	<td align='center'  colspan='5' >";
                    sHtml += "  CONNECTION DETAIL";
                    sHtml += " </td>";
                    sHtml += " </tr>";
                    sHtml += "          <tr>";
                    sHtml += " 				<td width='150'>POL</td>  ";
                    sHtml += " 				<td width='100'>ETD</td>  ";
                    sHtml += " 				<td width='150'>POD</td>  ";
                    sHtml += " 				<td width='100'>ETA</td>  ";
                    sHtml += " 				<td width='200'>VESSEL & VOYAGE</td>  ";
                    sHtml += " 			</tr>";
                    foreach (DataRow Drr in Dt_Tracking.Rows)
                    {
                        sHtml += " 			<tr>";
                        sHtml += " 				<td width='150' >" + Drr["POL_NAME"].ToString() + "</td>  ";
                        str = Lib.DatetoStringDisplayformat(Drr["ETD_DATE"]);
                        sHtml += " 				<td width='100' >" + str + "</td>  ";
                        sHtml += " 				<td width='150' >" + Drr["POD_NAME"].ToString() + "</td>  ";
                        str = Lib.DatetoStringDisplayformat(Drr["ETA_DATE"]);
                        sHtml += " 				<td width='100' >" + str + "</td>  ";
                        sHtml += " 				<td width='200' >" + Drr["VESSEL"].ToString() + " - " + Drr["VOYAGE"].ToString() + "</td>  ";
                        sHtml += " 			</tr>";
                    }

                    sHtml += " 		</table>";
                    sHtml += " 	<br /> <br />";
                    sHtml += " 		</td>";
                    sHtml += " </table>";
                    sHtml += " 	<br /> <br />";
                    sHtml += " </body>";
                    sHtml += " </html>";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return sHtml;
        }

        private void GetTrackingDetais(string MBLID)
        {
            sql = " select vsl.param_name as vessel";
            sql += " ,a.trk_voyage as voyage";
            sql += " ,a.trk_pol_etd as etd_date";
            sql += " ,a.trk_pol_etd_confirm";
            sql += " ,pol.param_name as pol_name";
            sql += " ,a.trk_pod_eta as eta_date";
            sql += " ,a.trk_pod_eta_confirm";
            sql += " ,pod.param_name as pod_name";
            sql += " ,a.trk_order as trk_order ";
            sql += " from trackingm a";
            sql += " inner join param  vsl on (a.trk_vsl_id=vsl.param_pkid)";
            sql += " inner join param pol on(a.trk_pol_id=pol.param_pkid)";
            sql += " inner join param pod on(a.trk_pod_id=pod.param_pkid)";
            sql += " where a.trk_parent_id ='" + MBLID + "'";
            sql += " order by trk_order";
            Con_Oracle = new DBConnection();
            Dt_Tracking = new DataTable();
            Dt_Tracking = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
        }

        private string GetJobInvoiceDetails(string HBL_ID)
        {
            bool differentInvDate = false;
            DataTable Dt_Temp = new DataTable();

            string InvNo = "";
            string InvDate = "";

            sql = " select distinct jexp_invoice_no,jexp_invoice_date from jobexpm a";
            sql += "  inner join jobm b on a.jexp_job_id = b.job_pkid ";
            sql += "  where b.jobs_hbl_id ='" + HBL_ID + "'";
            sql += "  order by jexp_invoice_no";
            Con_Oracle = new DBConnection();
            Dt_Temp = new DataTable();
            Dt_Temp = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            if (Dt_Temp.Rows.Count > 0)
            {
                DataTable DistinctINVDT = Dt_Temp.DefaultView.ToTable(true, "jexp_invoice_date");
                if (DistinctINVDT.Rows.Count > 1)
                    differentInvDate = true;
                foreach (DataRow dr in Dt_Temp.Rows)
                {
                    if (!dr["jexp_invoice_date"].Equals(DBNull.Value))
                        InvDate = ((DateTime)dr["jexp_invoice_date"]).ToString("dd/MM/yyyy");

                    if (InvNo.Trim() != "")
                        InvNo += ",";
                    InvNo += dr["jexp_invoice_no"].ToString();
                    if (differentInvDate)
                        InvNo += " DT: " + InvDate;
                }
            }
            if (!differentInvDate && InvNo.Trim() != "" && InvDate.Trim() != "")
                InvNo += " DT:" + InvDate;

            return InvNo;
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
    }
}
