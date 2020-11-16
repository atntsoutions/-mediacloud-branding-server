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
    public class TeuClrReportService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<TeuClrReport> mList = new List<TeuClrReport>();
        TeuClrReport mrow;
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
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string ErrorMessage = "";
        string shipper_id = "";
        string consginee_id = "";
        string agent_id = "";
        string carrier_id = "";
        string pol_id = "";
        string pod_id = "";
        Boolean all = false;

        decimal tot_teu = 0;
        decimal tot_pcs = 0;
        decimal tot_ntwt = 0;
        decimal tot_grwt = 0;
        decimal tot_cbm = 0;
        decimal tot_20 = 0;
        decimal tot_40 = 0;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<TeuClrReport>();
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

                all = (Boolean)SearchData["all"];

                if (SearchData.ContainsKey("shipper_id"))
                    shipper_id = SearchData["shipper_id"].ToString();

                if (SearchData.ContainsKey("consignee_id"))
                    consginee_id = SearchData["consignee_id"].ToString();
               
                if (SearchData.ContainsKey("agent_id"))
                    agent_id = SearchData["agent_id"].ToString();

                if (SearchData.ContainsKey("carrier_id"))
                    carrier_id = SearchData["carrier_id"].ToString();

                if (SearchData.ContainsKey("pol_id"))
                    pol_id = SearchData["pol_id"].ToString();

                if (SearchData.ContainsKey("pod_id"))
                    pod_id = SearchData["pod_id"].ToString();


                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);


                if (from_date == "NULL" || to_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");

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

                

                sql = " select job_date,branch,cntr_no,cntr_pcs,cntr_ntwt,cntr_grwt,cntr_cbm,cntr_clearing,cntr_csealno,cntr_asealno";
                sql += "   ,cntr_stuffed_on ,ctype.param_code as cntr_type_code,Cntr_teu";
                sql += " ,case when instr(ctype.param_name,'20') > 0 then 1 else 0 end as Cntr_20_tot";
                sql += "  ,case when (instr(ctype.param_name,'40') > 0 or instr(ctype.param_name,'45') > 0) then 1 else 0 end as Cntr_40_tot";
                sql += "   ,decode(job_type_tot,1,job_type, 'VARIOUS')       as job_type";
                sql += "  ,decode(job_nature_tot,1,job_nature, 'VARIOUS')   as job_nature";
                sql += "  ,decode(Exporter_tot,1,cust_name, 'VARIOUS')       as shipper_name";
                sql += "  ,decode(Exporter_tot,1,cust_code, 'VARIOUS')       as shipper_code";
                sql += "  ,decode(consignee_tot,1,cons.cust_name, 'VARIOUS') as consignee_name";
                sql += "  ,decode(consignee_tot,1,cons.cust_code, 'VARIOUS') as consignee_code";
                sql += "  ,decode(Nom_tot,1,Nomination, 'VARIOUS') 	  as Nomination";
                sql += "  ,decode(Agent_tot,1,Agent.cust_Name, 'VARIOUS')    as Agent_name";
                sql += "  ,decode(Agent_tot,1,Agent.cust_Code, 'VARIOUS')    as Agent_code";
                sql += "  ,decode(Liner_tot,1,Liner.param_Name, 'VARIOUS')    as Liner_name";
                sql += "  ,decode(Liner_tot,1,Liner.param_Code, 'VARIOUS')    as Liner_code";
                sql += "  ,decode(pol_tot,1,pol.param_Name, 'VARIOUS')       as Pol_name";
                sql += "  ,decode(pod_tot,1,pod.param_Name, 'VARIOUS')       as Pod_name";
                sql += "  ,decode(pod_tot,1,pod.param_Code, 'VARIOUS')       as Pod_code";
                sql += "  ,decode(pofd_tot,1,pofd.param_Name, 'VARIOUS')       as Pofd_name";
                sql += "  ,decode(pofd_tot,1,pofd.param_Code, 'VARIOUS')       as Pofd_code";
                sql += "  ,decode(Country_tot,1,Country.Param_Name, 'VARIOUS') as Country_name";
                sql += "  ,decode(Country_tot,1,Country.Param_Code, 'VARIOUS') as Country_code";
                sql += "  ,cntr_stuffed_at,a.cntr_pkid";
                sql += "  ,consignee_Id, Consignee_Tot,Liner_Id,Liner_tot,Agent_Id,branch_code";
                sql += "  ,Exporter_Id, Exporter_Tot";
                sql += "  ,Agent_tot,Country_Id,Country_Tot";
                sql += "  ,nvl(sman2.param_name, sman.param_name) as salesman";
                sql += "  ,decode(Exporter_tot,1,shpraddr.add_city, 'VARIOUS')   as shpr_Location";
                sql += "  ,a.rec_branch_code as br_location";
                sql += "  from ( ";
                sql += "  select cntr_pkid,d.rec_branch_code as branch,";
                sql += "  j.rec_branch_code,";
                sql += "  max(j.job_date) as job_date,";
                sql += "  max(j.job_nature)  as job_nature ,count(distinct j.job_nature)  as job_nature_Tot,";
                sql += "  max(j.job_exp_id)  as Exporter_Id ,count(distinct j.job_exp_id)  as Exporter_Tot,";
                sql += "  max(j.rec_branch_code) as branch_code,";
                sql += "  max(j.job_imp_id) as Consignee_Id,count(distinct j.job_imp_id) as Consignee_Tot, ";
                sql += "  max(j.job_carrier_id)     as Liner_Id,    count(distinct j.job_carrier_id)     as Liner_Tot,";
                sql += "  max(j.job_Agent_id)     as Agent_Id,    count(distinct j.job_Agent_id)     as Agent_Tot,";
                sql += "  max(j.job_nomination)  as Nomination,  count(distinct j.job_nomination)  as Nom_Tot,";
                sql += "  max(j.job_pol_id)      as Pol_Id,      count(distinct j.job_pol_id)      as Pol_Tot,";
                sql += "  max(j.job_pod_id)      as Pod_Id,      count(distinct j.job_pod_id)      as Pod_Tot,";
                sql += "  max(j.job_pofd_id)      as Pofd_Id,      count(distinct j.job_pofd_id)      as Pofd_Tot,";
                sql += "  max(job_pod_country_id)  as Country_Id,  count(distinct job_pod_country_id)  as Country_Tot,";
                sql += "  max(j.job_type)  as job_type ,count(distinct j.job_type)  as job_type_Tot";
                sql += "  from   containerm d";
                sql += "  inner  join packingm  a  on (d.cntr_pkid = a.pack_cntr_id )";
                sql += "  inner  join jobm        j  on (a.pack_job_id = j.job_pkid)";
                sql += "  where cntr_teu > 0 ";
                sql += "  and d.rec_company_code = '{COMPCODE}'";

                if(!all)
                {
                    sql += "  and d.rec_branch_code = '{BRCODE}'";
                }
               
                sql += "  and d.cntr_year = {YEARCODE} ";

                if (shipper_id.Length > 0)
                    sql += " and j.job_exp_id = '" + shipper_id + "' ";

                if (consginee_id.Length > 0)
                    sql += " and j.job_imp_id = '" + consginee_id + "' ";

                if (agent_id.Length > 0)
                    sql += " and j.job_Agent_id = '" + agent_id + "' ";
                if (carrier_id.Length > 0)
                    sql += " and j.job_carrier_id = '" + carrier_id + "'";
                if (pol_id.Length > 0)
                    sql += " and j.job_pol_id = '" + pol_id + "'";
                if (pod_id.Length > 0)
                    sql += " and j.job_pod_id = '" + pod_id + "'";

                sql += "  and j.job_type in ('CLEARING','BOTH')";

                sql += "  and j.job_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                sql += "  group by d.rec_branch_code,cntr_pkid,j.rec_branch_code";
                sql += "  )  a";
                sql += "  left join containerm cont on (a.cntr_pkid = cont.cntr_pkid)";
                sql += "  left join customerm cons on (a.consignee_id = cons.cust_pkid)";
                sql += "  left join param Liner on (a.Liner_id = Liner.param_pkid)";
                sql += "  left join customerm Agent on (a.Agent_id = Agent.cust_pkid)";
                sql += "  left join customerm 	 shipper on (a.exporter_id =shipper.cust_pkid)";
                sql += "  left join custdet  cd on branch_code = cd.det_branch_code and Exporter_Id = cd.det_cust_id ";
                sql += "  left join param sman on shipper.cust_sman_id = sman.param_pkid";
                sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                sql += "  left join param 	 country on (a.country_Id = country.param_pkid)";
                sql += "  left join param 	 pol on (a.pol_Id = pol.param_pkid)";
                sql += "  left join param 	 pod on (a.pod_Id = pod.param_pkid)";
                sql += "  left join param 	 pofd on (a.pofd_Id = pofd.param_pkid)";
                sql += "  left join param 	 ctype on (cont.cntr_type_Id = ctype.param_pkid) ";
                sql += "  left join addressm shpraddr on (a.exporter_id = shpraddr.add_pkid and shpraddr.add_branch_slno=0)";
                sql += "  order by branch,job_date";



                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                tot_teu = 0;
                tot_pcs = 0;
                tot_ntwt = 0;
                tot_grwt = 0;
                tot_cbm = 0;
                tot_20 = 0;
                tot_40 = 0;
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new TeuClrReport();
                    mrow.row_type = "DETAIL";
                    mrow.row_colour = "BLACK";
                   // mrow.cntr_pkid = Dr["cntr_pkid"].ToString();
                   mrow.job_date = Lib.DatetoStringDisplayformat(Dr["job_date"]);
                    mrow.cntr_no = Dr["cntr_no"].ToString();

                    mrow.cntr_pcs = Lib.Conv2Decimal(Dr["cntr_pcs"].ToString());
                    mrow.cntr_ntwt = Lib.Conv2Decimal(Dr["cntr_ntwt"].ToString());
                    mrow.cntr_grwt = Lib.Conv2Decimal(Dr["cntr_grwt"].ToString());
                    mrow.cntr_cbm = Lib.Conv2Decimal(Dr["cntr_cbm"].ToString());
                    mrow.cntr_clearing = Dr["cntr_clearing"].ToString();
                    mrow.cntr_csealno = Dr["cntr_csealno"].ToString();
                    mrow.cntr_asealno = Dr["cntr_asealno"].ToString();
                    mrow.cntr_stuffed_on = Lib.DatetoStringDisplayformat(Dr["cntr_stuffed_on"]);
                    mrow.cntr_type_code = Dr["cntr_type_code"].ToString();
                    mrow.cntr_teu = Lib.Conv2Decimal(Dr["Cntr_teu"].ToString());
                    mrow.cntr_20_tot = Lib.Conv2Decimal(Dr["Cntr_20_tot"].ToString());
                    mrow.cntr_40_tot = Lib.Conv2Decimal(Dr["Cntr_40_tot"].ToString());
                    mrow.job_type = Dr["job_type"].ToString();
                    mrow.job_nature = Dr["job_nature"].ToString();

                    mrow.hbl_exp_name = Dr["shipper_name"].ToString();
                    mrow.hbl_imp_name = Dr["consignee_name"].ToString();
                    mrow.job_nomination = Dr["Nomination"].ToString();
                    mrow.hbl_agent_name = Dr["Agent_name"].ToString();
                    mrow.hbl_carrier_name = Dr["Liner_name"].ToString();
                    mrow.hbl_pol_name = Dr["Pol_name"].ToString();
                    mrow.hbl_pod_name = Dr["Pod_name"].ToString();
                    mrow.hbl_pofd_name = Dr["Pofd_name"].ToString();
                    mrow.cntr_stuffed_at = Dr["cntr_stuffed_at"].ToString();
                    mrow.sman_name = Dr["salesman"].ToString();
                  
                    mrow.branch = Dr["branch"].ToString();
                    mList.Add(mrow);

                    tot_teu += Lib.Conv2Decimal(mrow.cntr_teu.ToString());
                    tot_pcs += Lib.Conv2Decimal(mrow.cntr_pcs.ToString());
                    tot_ntwt += Lib.Conv2Decimal(mrow.cntr_ntwt.ToString());
                    tot_grwt += Lib.Conv2Decimal(mrow.cntr_grwt.ToString());
                    tot_cbm += Lib.Conv2Decimal(mrow.cntr_cbm.ToString());
                    tot_20 += Lib.Conv2Decimal(mrow.cntr_20_tot.ToString());
                    tot_40 += Lib.Conv2Decimal(mrow.cntr_40_tot.ToString());
                }
                if (mList.Count > 1)
                {
                    mrow = new TeuClrReport();
                    mrow.row_type = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.cntr_no = "TOTAL";
                    mrow.cntr_teu = Lib.Conv2Decimal(Lib.NumericFormat(tot_teu.ToString(), 2));
                 //   mrow.cntr_pcs = Lib.Conv2Decimal(Lib.NumericFormat(tot_pcs.ToString(), 0));
                    mrow.cntr_ntwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_ntwt.ToString(), 3));
                    mrow.cntr_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_grwt.ToString(), 3));
                    mrow.cntr_cbm = Lib.Conv2Decimal(Lib.NumericFormat(tot_cbm.ToString(), 3));
                    mrow.cntr_20_tot = Lib.Conv2Decimal(Lib.NumericFormat(tot_20.ToString(), 3));
                    mrow.cntr_40_tot = Lib.Conv2Decimal(Lib.NumericFormat(tot_40.ToString(), 3));
                    mList.Add(mrow);
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintTeuReport();
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
            return RetData;
        }

        private void PrintTeuReport()
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
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }
                
                File_Display_Name = "TeuClr.xls";
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
                WS.Columns[5].Width = 256 * 20;
                WS.Columns[6].Width = 256 * 20;
                WS.Columns[7].Width = 256 * 20;
                WS.Columns[8].Width = 256 * 20;
                WS.Columns[9].Width = 256 * 20;
                WS.Columns[10].Width = 256 * 20;
                WS.Columns[11].Width = 256 * 20;
                WS.Columns[12].Width = 256 * 20;
                WS.Columns[13].Width = 256 * 20;
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
                WS.Columns[26].Width = 256 * 15;
                WS.Columns[27].Width = 256 * 15;
                WS.Columns[28].Width = 256 * 15;
                iRow = 0; iCol = 1;
                if(all)
                {
                   
                   
                    WS.Columns[21].Style.NumberFormat = "#0.00";
                    WS.Columns[22].Style.NumberFormat = "#0.00";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[24].Style.NumberFormat = "#0.000";
                    WS.Columns[25].Style.NumberFormat = "#0.000";
                    WS.Columns[26].Style.NumberFormat = "#0.000";
                }
                else
                {
                   
                    WS.Columns[20].Style.NumberFormat = "#0.00";
                    WS.Columns[21].Style.NumberFormat = "#0.00";
                    WS.Columns[22].Style.NumberFormat = "#0.000";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[24].Style.NumberFormat = "#0.000";
                    WS.Columns[25].Style.NumberFormat = "#0.000";

                }
                
                iRow++;
                _Size = 14;
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
                Lib.WriteData(WS, iRow, 1, "CLEARING REPORT", _Color, true, "", "L", "", 15, false, 325, "", true);
               
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "JOB-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTAINER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "A.SEALNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "C.SEALNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STUFFED AT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STUFFED ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NATURE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CLEARING", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CNTR.20", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CNTR.40", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PCS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                //iRow++;
                //iCol = 1;

                //str = mrow.cntr_booking_no != null ? mrow.cntr_booking_no : ""; 
                //Lib.WriteData(WS, iRow, iCol++,str, _Color, true, "", "C", "", _Size, false, 325, "", true);
                List<TeuClrReport> MList = new List<TeuClrReport>();
                foreach (TeuClrReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_type_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_asealno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_csealno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_exp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_imp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_carrier_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pol_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pod_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pofd_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_stuffed_at, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.cntr_stuffed_on, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_nature, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_clearing, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_20_tot, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_40_tot, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_pcs, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_no, _Color, true, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++,"", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++,"" , _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++,"" , _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_20_tot, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_40_tot, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_ntwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_grwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_cbm, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

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
    }
}

