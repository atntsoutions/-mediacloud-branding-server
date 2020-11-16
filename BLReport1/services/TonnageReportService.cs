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
    public class TonnageReportService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<TonnageReport> mList = new List<TonnageReport>();
        TonnageReport mrow;
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
        string consignee_id = "";
        string agent_id = "";
        string carrier_id = "";
        string pol_id = "";
        string pod_id = "";
        string rec_category = "";
        Boolean all = false;

        decimal tot_mbl_grwt = 0;
        decimal tot_mbl_chwt = 0;
        decimal tot_hbl_grwt = 0;
        decimal tot_hbl_chwt = 0;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<TonnageReport>();
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
                rec_category = SearchData["rec_category"].ToString();

                all = (Boolean)SearchData["all"];

                if (SearchData.ContainsKey("shipper_id"))
                    shipper_id = SearchData["shipper_id"].ToString();

                if (SearchData.ContainsKey("consignee_id"))
                    consignee_id = SearchData["consignee_id"].ToString();

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

                if (rec_category == "AIR IMPORT")
                {
                    sql = "  SELECT  MBL_PKID,MBL_DATE,CREATED_DATE,MBL_NO, MBL_GRWT,MBL_CHWT,HBL_GRWT,HBL_CHWT,BRANCH";
                    sql += "  ,Agent.cust_name as MBL_AGENT_NAME,liner.param_name as MBL_AIRLINE_NAME,";
                    sql += "  pod.param_name as MBL_POD_NAME,pofd.param_name as MBL_POFD_NAME";
                    sql += "  ,DECODE(Exporter_tot,1,shipper.cust_name, 'VARIOUS') AS HBL_SHIPPER_NAME";
                    sql += "  ,DECODE(consignee_tot,1,cons.cust_name, 'VARIOUS') AS HBL_CONSIGNEE_NAME";
                    sql += "  ,shipper.cust_nomination as HBL_NOMINATION";

                    sql += "  ,DECODE(pol_tot,1,pol.param_Name, 'VARIOUS')   AS HBL_POL_NAME";
                    sql += "  ,MBL_STATUS_NAME,Exporter_Id,branch_code,nvl(sman2.param_name,sman.param_name) as SMAN_NAME ";
                    sql += "  FROM ( ";
                    sql += "  SELECT";
                    sql += "  m.hbl_pkid as MBL_PKID ,m.hbl_bl_no as MBL_NO,m.hbl_terms as MBL_STATUS_NAME,m.rec_branch_code as BRANCH,";
                    sql += "  MAX(m.hbl_grwt) AS MBL_GRWT,";
                    sql += "  MAX(m.hbl_chwt) AS MBL_CHWT,";
                    sql += "  SUM(h.hbl_grwt) AS HBL_GRWT ,";
                    sql += "  SUM(h.hbl_chwt) AS HBL_CHWT,";
                    sql += "  MAX(m.hbl_date) AS MBL_DATE,";
                    sql += "  MAX(m.rec_created_date) AS CREATED_DATE,";
                    sql += "  MAX(h.hbl_exp_id)  AS Exporter_Id ,COUNT(DISTINCT h.hbl_exp_id)  AS Exporter_Tot,";
                    sql += "  MAX(h.rec_branch_code) as branch_code,";
                    sql += "  MAX(h.hbl_imp_id) AS Consignee_Id,COUNT(DISTINCT h.hbl_imp_id) AS Consignee_Tot,";
                    sql += "  MAX(m.hbl_carrier_id)     AS Liner_Id,    COUNT(DISTINCT m.hbl_carrier_id)     AS Liner_Tot,";
                    sql += "  MAX(m.hbl_Agent_id)     AS Agent_Id,    COUNT(DISTINCT m.hbl_Agent_id)     AS Agent_Tot,";
                    sql += "  MAX(m.hbl_pod_id)      AS Pod_Id,      COUNT(DISTINCT m.hbl_pod_id)      AS Pod_Tot,";
                    sql += "  MAX(m.hbl_pofd_id)     AS Pofd_Id,     COUNT(DISTINCT m.hbl_pofd_id)     AS Pofd_Tot,";
                    sql += "  MAX(h.hbl_pol_id)    As pol_Id,   COUNT(DISTINCT h.hbl_pol_id)   AS pol_Tot";
                    sql += "  FROM  hblm m";
                    sql += "  inner join hblm h on (m.hbl_pkid = h.hbl_mbl_id and m.hbl_type='MBL-AI') ";
                    sql += "  where m.hbl_jobtype in ('FORWARDING','BOTH')";
                    sql += "  and m.rec_company_code = '{COMPCODE}'";
                    if (!all)
                    {
                        sql += "  and m.rec_branch_code = '{BRCODE}'";
                    }
                    sql += "  and m.hbl_year = {YEARCODE} ";

                    if (type_date == "MAWB DATE")
                        sql += "  and m.hbl_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    else
                        sql += "  and to_char(m.rec_created_date,'DD-MON-YYYY') between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                    if (consignee_id.Length > 0)
                        sql += " and h.hbl_exp_id = '" + consignee_id + "' ";
                    if (shipper_id.Length > 0)
                        sql += " and h.hbl_imp_id  = '" + shipper_id + "' ";
                    if (agent_id.Length > 0)
                        sql += " and m.hbl_Agent_id = '" + agent_id + "' ";
                    if (carrier_id.Length > 0)
                        sql += " and m.hbl_carrier_id = '" + carrier_id + "'";
                    if (pol_id.Length > 0)
                        sql += " and h.hbl_pol_id = '" + pol_id + "'";
                    if (pod_id.Length > 0)
                        sql += " and m.hbl_pod_id = '" + pod_id + "'";

                    sql += "  GROUP BY m.rec_branch_code,m.hbl_pkid,m.hbl_bl_no,m.hbl_terms";
                    sql += "  )  a";
                    sql += "  left join customerm cons ON (a.consignee_Id = cons.cust_pkid)";
                    sql += "  left join param Liner ON (a.Liner_Id = Liner.param_pkid)";
                    sql += "  left join customerm Agent ON (a.Agent_Id = Agent.cust_pkid)";
                    sql += "  left join customerm 	 shipper ON (a.Exporter_Id =shipper.cust_pkid)";
                    sql += "  left join custdet  cd on branch_code = cd.det_branch_code and consignee_Id  = cd.det_cust_id ";
                    sql += "  left join param sman on cons.cust_sman_id = sman.param_pkid";
                    sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += "  left join param 	 pod ON (a.Pod_Id = pod.param_pkid)";
                    sql += "  left join param 	 pofd ON (a.Pofd_Id = pofd.param_pkid)";
                    sql += "  left join param    pol ON (a.pol_Id = pol.param_pkid)";

                    if (type_date == "MAWB DATE")
                        sql += "  order by BRANCH,MBL_DATE";
                    else
                        sql += "  order by BRANCH,CREATED_DATE";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    tot_mbl_grwt = 0;
                    tot_mbl_chwt = 0;
                    tot_hbl_grwt = 0;
                    tot_hbl_chwt = 0;

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TonnageReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.mbl_pkid = Dr["mbl_pkid"].ToString();
                        mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.mbl_grwt = Lib.Conv2Decimal(Dr["mbl_grwt"].ToString());
                        mrow.mbl_chwt = Lib.Conv2Decimal(Dr["mbl_chwt"].ToString());
                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                        mrow.hbl_chwt = Lib.Conv2Decimal(Dr["hbl_chwt"].ToString());
                        mrow.hbl_shipper_name = Dr["hbl_shipper_name"].ToString();
                        mrow.hbl_consignee_name = Dr["hbl_consignee_name"].ToString();
                        mrow.hbl_nomination = Dr["hbl_nomination"].ToString();
                        mrow.mbl_agent_name = Dr["mbl_agent_name"].ToString();
                        mrow.mbl_airline_name = Dr["mbl_airline_name"].ToString();
                        mrow.hbl_pod_name = Dr["mbl_pod_name"].ToString();
                        mrow.hbl_pofd_name = Dr["mbl_pofd_name"].ToString();
                        mrow.mbl_status_name = Dr["mbl_status_name"].ToString();
                        mrow.sman_name = Dr["sman_name"].ToString();

                        mrow.hbl_pol_name = Dr["hbl_pol_name"].ToString();
                        mrow.branch = Dr["BRANCH"].ToString();

                        mList.Add(mrow);

                        tot_mbl_grwt += Lib.Conv2Decimal(Dr["mbl_grwt"].ToString());
                        tot_mbl_chwt += Lib.Conv2Decimal(Dr["mbl_chwt"].ToString());
                        tot_hbl_grwt += Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                        tot_hbl_chwt += Lib.Conv2Decimal(Dr["hbl_chwt"].ToString());

                    }
                    if (mList.Count > 1)
                    {
                        mrow = new TonnageReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.mbl_no = "TOTAL";
                        mrow.mbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_mbl_grwt.ToString(), 3));
                        mrow.mbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_mbl_chwt.ToString(), 3));
                        mrow.hbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_hbl_grwt.ToString(), 3));
                        mrow.hbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_hbl_chwt.ToString(), 3));
                        mList.Add(mrow);
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintTonnageReport();
                    }
                    Dt_List.Rows.Clear();

                }
                else
                {

                    sql = " SELECT  MBL_PKID,MBL_DATE,CREATED_DATE,MBL_NO, MBL_GRWT,MBL_CHWT,HBL_GRWT,HBL_CHWT,BRANCH";
                    sql += " ,DECODE(Exporter_tot,1,shipper.cust_name, 'VARIOUS') AS HBL_SHIPPER_NAME";
                    sql += " ,DECODE(consignee_tot,1,cons.cust_name, 'VARIOUS') AS HBL_CONSIGNEE_NAME";
                   // sql += " ,cons.cust_nomination as HBL_NOMINATION";
                    sql += " ,DECODE(consignee_tot,1,nvl(nomination,cons.cust_nomination), 'VARIOUS') as  HBL_NOMINATION";
                    // sql += " ,DECODE(Nom_tot,1,Nomination, 'VARIOUS') 	  AS HBL_NOMINATION";
                    sql += " ,DECODE(Agent_tot,1,Agent.cust_Name, 'VARIOUS')    AS MBL_AGENT_NAME";
                    sql += " ,DECODE(Liner_tot,1,Liner.param_Name, 'VARIOUS')    AS MBL_AIRLINE_NAME";
                    sql += " ,DECODE(pod_tot,1,pod.param_Name, 'VARIOUS')       AS HBL_POD_NAME";
                    sql += " ,DECODE(pofd_tot,1,pofd.param_Name, 'VARIOUS')       AS HBL_POFD_NAME";

                    sql += " ,DECODE(pol_tot,1,pol.param_Name, 'VARIOUS')   AS HBL_POL_NAME";

                    sql += " ,MBL_STATUS_NAME,Exporter_Id,branch_code,";
                    //sql += " nvl(sman2.param_name,sman.param_name) as SMAN_NAME ";//sman.param_name as SMAN_NAME
                    sql += " nvl(sman3.param_name,nvl(sman2.param_name, sman.param_name)) as SMAN_NAME ";

                    sql += " FROM ( ";
                    sql += " SELECT";
                    sql += " m.hbl_pkid as MBL_PKID ,m.hbl_bl_no as MBL_NO,m.hbl_terms as MBL_STATUS_NAME,m.rec_branch_code as BRANCH,";
                    sql += " MAX(m.hbl_grwt) AS MBL_GRWT,";
                    sql += " MAX(m.hbl_chwt) AS MBL_CHWT,";
                    sql += " SUM(h.hbl_grwt) AS HBL_GRWT ,";
                    sql += " SUM(h.hbl_chwt) AS HBL_CHWT,";
                    sql += " MAX(m.hbl_date) AS MBL_DATE,";
                    sql += " MAX(m.rec_created_date) AS CREATED_DATE,";
                    sql += " MAX(h.hbl_exp_id)  AS Exporter_Id ,COUNT(DISTINCT h.hbl_exp_id)  AS Exporter_Tot,";
                    sql += " MAX(h.rec_branch_code) as branch_code,";
                    sql += " MAX(h.hbl_imp_id) AS Consignee_Id,COUNT(DISTINCT h.hbl_imp_id) AS Consignee_Tot,";
                    sql += " MAX(m.hbl_carrier_id)     AS Liner_Id,    COUNT(DISTINCT m.hbl_carrier_id)     AS Liner_Tot,";
                    sql += " MAX(m.hbl_Agent_id)     AS Agent_Id,    COUNT(DISTINCT m.hbl_Agent_id)     AS Agent_Tot,";
                    //  sql += " MAX(h.hbl_nomination)  AS Nomination,  COUNT(DISTINCT h.hbl_nomination)  AS Nom_Tot,";
                    sql += " max(h.hbl_nomination)  as Nomination,  count(distinct h.hbl_nomination)  as Nom_Tot,";
                    sql += " max(h.hbl_salesman_id)  as hbl_salesman_id,  ";
                    sql += " MAX(h.hbl_pod_id)      AS Pod_Id,      COUNT(DISTINCT h.hbl_pod_id)      AS Pod_Tot,";
                    sql += " MAX(h.hbl_pofd_id)     AS Pofd_Id,     COUNT(DISTINCT h.hbl_pofd_id)     AS Pofd_Tot,";

                    sql += " MAX(h.hbl_pol_id)    As pol_Id,   COUNT(DISTINCT h.hbl_pol_id)   AS pol_Tot";

                    sql += " FROM  hblm m";
                    sql += " inner join hblm h on (m.hbl_pkid = h.hbl_mbl_id and m.hbl_type='MBL-AE') ";
                    sql += " where m.rec_company_code = '{COMPCODE}'";
                    if (!all)
                    {
                        sql += " and m.rec_branch_code = '{BRCODE}'";
                    }
                    sql += " and m.hbl_year = {YEARCODE} ";
                    if (type_date == "MAWB DATE")
                        sql += "  and m.hbl_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    else
                        sql += "  and to_char(m.rec_created_date,'DD-MON-YYYY') between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                    if (shipper_id.Length > 0)
                        sql += " and h.hbl_exp_id = '" + shipper_id + "' ";
                    if (consignee_id.Length > 0)
                        sql += " and h.hbl_imp_id  = '" + consignee_id + "' ";
                    if (agent_id.Length > 0)
                        sql += " and m.hbl_Agent_id = '" + agent_id + "' ";
                    if (carrier_id.Length > 0)
                        sql += " and m.hbl_carrier_id = '" + carrier_id + "'";
                    if (pol_id.Length > 0)
                        sql += " and h.hbl_pol_id = '" + pol_id + "'";
                    if (pod_id.Length > 0)
                        sql += " and h.hbl_pod_id = '" + pod_id + "'";


                    sql += " GROUP BY m.rec_branch_code,m.hbl_pkid,m.hbl_bl_no,m.hbl_terms";
                    sql += " )  a";
                    sql += " left join customerm cons ON (a.consignee_Id = cons.cust_pkid)";
                    sql += " left join param Liner ON (a.Liner_Id = Liner.param_pkid)";
                    sql += " left join customerm Agent ON (a.Agent_Id = Agent.cust_pkid)";

                    sql += " left join customerm 	 shipper ON (a.Exporter_Id =shipper.cust_pkid)";
                    sql += "  left join custdet  cd on branch_code = cd.det_branch_code and Exporter_Id = cd.det_cust_id ";

                    sql += " left join param sman on shipper.cust_sman_id = sman.param_pkid";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                    sql += " left join param sman3 on hbl_salesman_id = sman3.param_pkid";

                    sql += " left join param 	 pod ON (a.Pod_Id = pod.param_pkid)";
                    sql += " left join param 	 pofd ON (a.Pofd_Id = pofd.param_pkid)";
                    sql += " left join param    pol ON (a.pol_Id = pol.param_pkid)";


                    if (type_date == "MAWB DATE")
                        sql += " order by BRANCH,MBL_DATE";
                    else
                        sql += " order by BRANCH,CREATED_DATE";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    tot_mbl_grwt = 0;
                    tot_mbl_chwt = 0;
                    tot_hbl_grwt = 0;
                    tot_hbl_chwt = 0;

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TonnageReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        mrow.mbl_pkid = Dr["mbl_pkid"].ToString();
                        mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.mbl_grwt = Lib.Conv2Decimal(Dr["mbl_grwt"].ToString());
                        mrow.mbl_chwt = Lib.Conv2Decimal(Dr["mbl_chwt"].ToString());
                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                        mrow.hbl_chwt = Lib.Conv2Decimal(Dr["hbl_chwt"].ToString());
                        mrow.hbl_shipper_name = Dr["hbl_shipper_name"].ToString();
                        mrow.hbl_consignee_name = Dr["hbl_consignee_name"].ToString();
                        mrow.hbl_nomination = Dr["hbl_nomination"].ToString();
                        mrow.mbl_agent_name = Dr["mbl_agent_name"].ToString();
                        mrow.mbl_airline_name = Dr["mbl_airline_name"].ToString();
                        mrow.hbl_pod_name = Dr["hbl_pod_name"].ToString();
                        mrow.hbl_pofd_name = Dr["hbl_pofd_name"].ToString();
                        mrow.mbl_status_name = Dr["mbl_status_name"].ToString();
                        mrow.sman_name = Dr["sman_name"].ToString();

                        mrow.hbl_pol_name = Dr["hbl_pol_name"].ToString();
                        mrow.branch = Dr["BRANCH"].ToString();

                        mList.Add(mrow);

                        tot_mbl_grwt += Lib.Conv2Decimal(mrow.mbl_grwt.ToString());
                        tot_mbl_chwt += Lib.Conv2Decimal(mrow.mbl_chwt.ToString());
                        tot_hbl_grwt += Lib.Conv2Decimal(mrow.hbl_grwt.ToString());
                        tot_hbl_chwt += Lib.Conv2Decimal(mrow.hbl_chwt.ToString());

                    }
                    if (mList.Count > 1)
                    {
                        mrow = new TonnageReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.mbl_no = "TOTAL";
                        mrow.mbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_mbl_grwt.ToString(), 3));
                        mrow.mbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_mbl_chwt.ToString(), 3));
                        mrow.hbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_hbl_grwt.ToString(), 3));
                        mrow.hbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_hbl_chwt.ToString(), 3));
                        mList.Add(mrow);
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintTonnageReport();
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

        private void PrintTonnageReport()
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
                    mSearchData.Add("branch_code","HOCPL");
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



                File_Display_Name = "Tonnage.xls";
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
                WS.Columns[3].Width = 256 * 30;
                WS.Columns[4].Width = 256 * 30;
                WS.Columns[5].Width = 256 * 30;
                WS.Columns[6].Width = 256 * 20;
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


                if(all)
                {
                    WS.Columns[14].Style.NumberFormat = "#0.000";
                    WS.Columns[15].Style.NumberFormat = "#0.000";
                    WS.Columns[16].Style.NumberFormat = "#0.000";
                    WS.Columns[17].Style.NumberFormat = "#0.000";
                }
                else
                {
                    WS.Columns[13].Style.NumberFormat = "#0.000";
                    WS.Columns[14].Style.NumberFormat = "#0.000";
                    WS.Columns[15].Style.NumberFormat = "#0.000";
                    WS.Columns[16].Style.NumberFormat = "#0.000";
                }
                
                iRow = 0; iCol = 1;
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
                if(rec_category == "TON-EXP")
                {
                    Lib.WriteData(WS, iRow, 1, "TONNAGE REPORT-EXPORT", _Color, true, "", "L", "", 15, false, 325, "", true);
                }

                if (rec_category == "AIR IMPORT")
                {
                    Lib.WriteData(WS, iRow, 1, "TONNAGE REPORT-IMPORT", _Color, true, "", "L", "", 15, false, 325, "", true);
                }


                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MAWB.NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AIRLINE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DESTINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HAWB.GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HAWB.CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MAWB.GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MAWB.CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                foreach (TonnageReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_shipper_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_consignee_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_airline_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_status_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pol_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pod_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pofd_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, true, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    }
                }




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

