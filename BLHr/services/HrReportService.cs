using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLHr
{
    public class HrReportService : BL_Base
    {
        ExcelFile file;
        ExcelWorksheet ws = null;
        ExcelWorksheet ws2 = null;
        CellRange myCell;
        List<HrReport> mList = null;
        HrReport mRow;
        DataTable Dt_EMP = new DataTable();
        string type = "";
        string searchstring = "";
        string branch_code = "";
        string company_code = "";
        string branch_region = "";
        string empstatus = "";
        string reporttype = "";
        string report_folder = "";
        string folderid = "";
        string File_Name = "";
        string File_Type = "";
        string File_Display_Name = "myreport.pdf";
        int salmonth = 0;
        int salyear = 0;
        int iCol = 0;
        string comp_name = "", comp_add1 = "", comp_add2 = "", comp_add3 = "";
        bool IsConsol = false;
        private int CopyFromRow = 0;
        private int CopyToRow = 0;
        private int CopyDelhiTotRow = 0;
        Dictionary<string, decimal> myDict = null;
        private string ImagePath = "";
        private string uploadfileid = "";
        int wRow = 0;
        int wCol = 0;
        Boolean IsTotalMatch = true;
        int ws_active_sheetno = 0;
        private string print_date = "";
        private string effective_date = "";
        private string sError = "";
        private DateTime DTprint = DateTime.Now;
        private DateTime DTEffective = DateTime.Now;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            sError = "";

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            mList = new List<HrReport>();
            
            type = SearchData["type"].ToString();
            searchstring = SearchData["searchstring"].ToString().ToUpper();
            branch_code = SearchData["branch_code"].ToString();
            company_code = SearchData["company_code"].ToString();
            empstatus = SearchData["empstatus"].ToString();
            reporttype = SearchData["reporttype"].ToString();
            salyear = Lib.Conv2Integer(SearchData["salyear"].ToString());
            salmonth = 0;
            if (SearchData.ContainsKey("salmonth"))
                salmonth = Lib.Conv2Integer(SearchData["salmonth"].ToString());
            branch_region = "";
            if (SearchData.ContainsKey("branch_region"))
                branch_region = SearchData["branch_region"].ToString();
            report_folder = "";
            if (SearchData.ContainsKey("report_folder"))
                report_folder = SearchData["report_folder"].ToString();
            folderid = "";
            if (SearchData.ContainsKey("folderid"))
                folderid = SearchData["folderid"].ToString();
            try
            {
                IsConsol = false;

                if (reporttype == "EPF")
                    ProcessEPF(branch_code, true);
                if (reporttype == "EPF-SOUTH")
                    ProcessConsolEPF("SOUTH");
                if (reporttype == "EPF-NORTH")
                    ProcessConsolEPF("NORTH");
                if (reporttype == "ESI")
                    ProcessESI(branch_code, true);
                if (reporttype == "ESI-SOUTH")
                    ProcessConsolESI("SOUTH");
                if (reporttype == "ESI-NORTH")
                    ProcessConsolESI("NORTH");

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", mList);

            return RetData;
        }


        private void ProcessEPF(string Br_Code, bool CanPrint)
        {
            decimal AdminPercent = 0;
            decimal EdliPercent = 0;
            decimal AdminAmount = 0;
            decimal EdliAmount = 0;
            string EdliBasedOn = "T";
            string sWhere = "";

            sWhere = " where a.rec_company_code = '" + company_code + "'";
            sWhere += " and a.rec_branch_code = '" + Br_Code + "'";
            sWhere += " and a.sal_month = " + salmonth.ToString();
            sWhere += " and a.sal_year = " + salyear.ToString();
            if (empstatus != "BOTH")
                sWhere += " and a.rec_category = '" + empstatus.ToString() + "'";
            if (searchstring != "")
            {
                sWhere += " and (";
                sWhere += "  upper(b.emp_name) like '%" + searchstring.ToUpper() + "%'";
                sWhere += " or ";
                sWhere += "  b.emp_no like '%" + searchstring.ToUpper() + "%'";
                sWhere += " ) ";
            }

            sql = "   select ";
            sql += "  sal_pkid,emp_name,emp_no,emp_pfno,sal_pf_mon_year,sal_pf_bal,";
            sql += "  sal_pf_base  as basic_da,";
            sql += "  sal_pf_emplr_pension  as pension, ";// pension=if(basic+da)>6500 then 541 (max limit) else (basic+da)* 8.33/100
            sql += "  sal_pf_base  as lmt_basic_da,";//north basic+da <=6500 for edli chrgs
            sql += "  sal_admin_based_on,sal_admin_per,sal_admin_amt,sal_edli_based_on,sal_edli_per,sal_edli_amt, ";
            sql += "  d01- nvl(sal_pf_bal,0) as deductn ,";
            sql += "  d14 as vpf ,";
            sql += "  a.rec_branch_code as  branch, ";
            sql += "  sal_pf_eps_amt  as eps_amt ";
            sql += "  from  salarym a";
            sql += "  inner join  empm b on a.sal_emp_id = b.emp_pkid ";
            sql += "  " + sWhere;
            sql += "  order by EMP_NO ";

            DataTable Dt_List = new DataTable();
            Dt_List = Con_Oracle.ExecuteQuery(sql);
            Dt_List.Columns.Add("EMPR_SHARE", typeof(System.Decimal));
            Dt_List.Columns.Add("ADMIN_CHRG", typeof(System.Decimal));
            Dt_List.Columns.Add("EDLI_CHRG", typeof(System.Decimal));
            Dt_List.Columns.Add("TOTAL", typeof(System.Decimal));

            decimal nTot = 0; EdliBasedOn = "T";
            foreach (DataRow dr in Dt_List.Rows)
            {
                nTot = Lib.Convert2Decimal(dr["DEDUCTN"].ToString());
                nTot -= Lib.Convert2Decimal(dr["PENSION"].ToString());
                dr["EMPR_SHARE"] = Lib.NumericFormat(nTot.ToString(), 0);

                EdliBasedOn = dr["SAL_EDLI_BASED_ON"].ToString().Trim();
                //dr["EDLI_CHRG"] = 0;
                if (EdliBasedOn == "E")
                    dr["EDLI_CHRG"] = Lib.NumericFormat(dr["SAL_EDLI_AMT"].ToString(), 2);

                mRow = new HrReport();
                mRow.row_type = "DETAIL";
                mRow.row_colour = "BLACK";
                mRow.branch = dr["branch"].ToString();
                mRow.emp_no = dr["emp_no"].ToString();
                mRow.emp_name = dr["emp_name"].ToString();
                mRow.emp_pfno = dr["emp_pfno"].ToString();
                mRow.pf_base_salary = Lib.Conv2Decimal(dr["basic_da"].ToString());
                mRow.pf_deduction = Lib.Conv2Decimal(dr["deductn"].ToString());
                mRow.emplyr_share = Lib.Conv2Decimal(dr["empr_share"].ToString());
                mRow.pension = Lib.Conv2Decimal(dr["pension"].ToString());
                mRow.vpf = Lib.Conv2Decimal(dr["vpf"].ToString());
                //mRow.admin_chrg = 0;
                if (EdliBasedOn == "E")
                    mRow.edli_chrg = Lib.Conv2Decimal(dr["edli_chrg"].ToString());
                //mRow.total_chrg =0;
                mRow.eps_amt = Lib.Conv2Decimal(dr["eps_amt"].ToString());

                if (EdliBasedOn == "T")
                    mRow.edli_based_on = "TOTAL";
                else if (EdliBasedOn == "E")
                    mRow.edli_based_on = "EMPLOYEE";
                else
                    mRow.edli_based_on = "FIXED";
                mList.Add(mRow);
            }

            object TotBasicDa = null;
            object TotEdliAmt = null;
            object TotEPN = null;
            object TotEPF = null;
            object TotVPF = null;
            object TotDEDUCT = null;

            if (Dt_List.Rows.Count > 0)
            {
                AdminPercent = Lib.Convert2Decimal(Dt_List.Rows[0]["SAL_ADMIN_PER"].ToString());
                EdliPercent = Lib.Convert2Decimal(Dt_List.Rows[0]["SAL_EDLI_PER"].ToString());
                if (Dt_List.Rows[0]["SAL_ADMIN_BASED_ON"].ToString().Trim() == "F")
                    AdminAmount = Lib.Convert2Decimal(Dt_List.Compute("max(SAL_ADMIN_AMT)", "1=1").ToString());
                else
                    AdminAmount = 0;
                if (Dt_List.Rows[0]["SAL_EDLI_BASED_ON"].ToString().Trim() == "F")
                    EdliAmount = Lib.Convert2Decimal(Dt_List.Compute("max(SAL_EDLI_AMT)", "1=1").ToString());//CPL 54,TWL 0, CPL TPR 0 since all under Cpl edli fixed on 54 for sep TWL compy it will be 0
                else
                    EdliAmount = 0;

                TotBasicDa = Dt_List.Compute("sum(BASIC_DA)", "1=1");
                TotEdliAmt = Dt_List.Compute("sum(EDLI_CHRG)", "1=1");
                TotEPN = Dt_List.Compute("sum(EMPR_SHARE)", "1=1");
                TotEPF = Dt_List.Compute("sum(PENSION)", "1=1");
                TotVPF = Dt_List.Compute("sum(VPF)", "1=1");
                TotDEDUCT = Dt_List.Compute("sum(DEDUCTN)", "1=1");

                decimal AdminChrg = 0, EDLIChrg = 0, nTotChrg = 0;

                if (AdminAmount > 0)
                    AdminChrg = AdminAmount;
                else
                    AdminChrg = Lib.Convert2Decimal(TotBasicDa.ToString()) * (AdminPercent / 100); //Common.Convert2Decimal("0.011");
                AdminChrg = Lib.Convert2Decimal(Lib.NumericFormat(AdminChrg.ToString(), 2));

                if (EdliAmount > 0 || EdliBasedOn == "F")
                    EDLIChrg = EdliAmount;
                else
                {
                    if (EdliBasedOn == "E")
                        EDLIChrg = Lib.Convert2Decimal(TotEdliAmt.ToString());
                    else
                    {
                        EDLIChrg = Lib.Convert2Decimal(TotBasicDa.ToString()) * (EdliPercent / 100);//Common.Convert2Decimal("0.00005");
                        EDLIChrg = Lib.Convert2Decimal(Lib.NumericFormat(EDLIChrg.ToString(), 2));
                    }
                }

                nTotChrg = Math.Round(Lib.Convert2Decimal(TotEPN.ToString()));
                nTotChrg += Math.Round(Lib.Convert2Decimal(TotEPF.ToString()));
                nTotChrg += Math.Round(Lib.Convert2Decimal(TotDEDUCT.ToString()));
                // nTotChrg += Math.Round(AdminChrg);
                nTotChrg += AdminChrg;
                nTotChrg += EDLIChrg;

                mRow = new HrReport();
                mRow.row_type = "TOTAL";
                mRow.row_colour = "RED";
                mRow.emp_name = "TOTAL";
                mRow.branch = "";
                mRow.pf_base_salary = Math.Round(Lib.Convert2Decimal(TotBasicDa.ToString()));
                mRow.pf_deduction = Math.Round(Lib.Convert2Decimal(TotDEDUCT.ToString()));
                mRow.emplyr_share = Math.Round(Lib.Convert2Decimal(TotEPN.ToString()));
                mRow.pension = Math.Round(Lib.Convert2Decimal(TotEPF.ToString()));
                mRow.vpf = Math.Round(Lib.Convert2Decimal(TotVPF.ToString()));
                mRow.admin_chrg = AdminChrg;
                mRow.edli_chrg = EDLIChrg;
                mRow.total_chrg = nTotChrg;
                mList.Add(mRow);
            }

            if (type == "EXCEL" && CanPrint)
            {
                PrintEPF();
            }
        }
        
        private void ProcessConsolEPF(string Br_Region)
        {
            IsConsol = true;
           // int TotNoOfEmp = 0;
            Dt_EMP = new DataTable();
            Dt_EMP.Columns.Add("BRANCH", typeof(System.String));
            Dt_EMP.Columns.Add("EMP_NAME", typeof(System.String));
            Dt_EMP.Columns.Add("EMP_NO", typeof(System.String));
            Dt_EMP.Columns.Add("EMP_STATUS", typeof(System.String));

            sql = "select rec_branch_code from payroll_setting where rec_company_code='" + company_code + "' and ps_pf_br_region ='" + Br_Region + "' order by case when rec_branch_code='HOCPL' then 1 else 2 end,  rec_branch_code";
            DataTable Dt_Branch = new DataTable();
            Dt_Branch = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow drb in Dt_Branch.Rows)
            {
                LoadNewEmp(drb["rec_branch_code"].ToString());
                ProcessEPF(drb["rec_branch_code"].ToString(), false);
            }
            if (type == "EXCEL")
            {
                PrintEPF();
            }
        }
        private void ProcessConsolESI(string Br_Region)
        {
            IsConsol = true;
            myDict = new Dictionary<string, decimal>();

            sql = "select rec_branch_code from payroll_setting where rec_company_code='" + company_code + "' and ps_pf_br_region ='" + Br_Region + "' order by case when rec_branch_code='HOCPL' then 1 else 2 end, rec_branch_code";
            DataTable Dt_Branch = new DataTable();
            Dt_Branch = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow drb in Dt_Branch.Rows)
            {
                ProcessESI(drb["rec_branch_code"].ToString(), false);
            }
            if (type == "EXCEL")
            {
                PrintESI();
            }
        }
        private void AddToList(string sKey, decimal nTot)
        {
            if (myDict.ContainsKey(sKey))
            {
                nTot = nTot + myDict[sKey];
                myDict[sKey] = nTot;
            }
            else
            {
                myDict.Add(sKey, nTot);
            }
        }
        private void ProcessESI(string Br_Code, bool CanPrint)
        {
          
            string sWhere = "";

            sWhere = " where a.rec_company_code = '" + company_code + "'";
            sWhere += " and a.rec_branch_code = '" + Br_Code + "'";
            sWhere += " and a.sal_month = " + salmonth.ToString();
            sWhere += " and a.sal_year = " + salyear.ToString();
            sWhere += " and nvl(sal_is_esi,'N') = 'Y'";
            if (empstatus != "BOTH")
                sWhere += " and a.rec_category = '" + empstatus.ToString() + "'";
            if (searchstring != "")
            {
                sWhere += " and (";
                sWhere += "  upper(b.emp_name) like '%" + searchstring.ToUpper() + "%'";
                sWhere += " or ";
                sWhere += "  b.emp_no like '%" + searchstring.ToUpper() + "%'";
                sWhere += " ) ";
            }
           

            sql = " select sal_pkid,emp_name,emp_no,sal_gross_earn";
            sql += " ,emp_esino ";
            sql += " ,d02 as emply_esi";
            sql += " ,null as emplr_esi";
            sql += " ,a.rec_branch_code as branch,sal_esi_emplr_per ";
            sql += " from salarym a";
            sql += " inner join empm b on (a.sal_emp_id = b.emp_pkid)";
            sql += "  " + sWhere;
            sql += "  order by emp_esino desc, emp_no ";

            DataTable Dt_List = new DataTable();
            Dt_List = Con_Oracle.ExecuteQuery(sql);
           // Dt_List.Columns.Add("TOTAL", typeof(System.Decimal));

            foreach (DataRow dr in Dt_List.Rows)
            {
                mRow = new HrReport();
                mRow.row_type = "DETAIL";
                mRow.row_colour = "BLACK";
                mRow.emp_no = dr["emp_no"].ToString();
                mRow.emp_name = dr["emp_name"].ToString();
                mRow.emp_esino = dr["emp_esino"].ToString();
                mRow.sal_gross_earn = Lib.Conv2Decimal(dr["sal_gross_earn"].ToString());
                mRow.emply_esi = Lib.Conv2Decimal(dr["emply_esi"].ToString());
                mRow.branch = dr["branch"].ToString();
                mList.Add(mRow);
            }

            object TotSal = null;
            object TotEMPYE = null;
            object TotEMPYR = null;
            object TotTOT = null;
           
            if (Dt_List.Rows.Count > 0)
            {
                TotSal = Dt_List.Compute("sum(SAL_GROSS_EARN)", "1=1");
                TotEMPYE = Dt_List.Compute("sum(EMPLY_ESI)", "1=1");
    
                if (IsConsol)
                {
                    double EmpyrPer = (double)(Lib.Convert2Decimal(Dt_List.Rows[0]["SAL_ESI_EMPLR_PER"].ToString()) / 100);
                    if (Br_Code=="CHNSF" || Br_Code == "CHNAF" || Br_Code == "COKSF" || Br_Code == "COKAF" || Br_Code == "HOCPL" || Br_Code == "SEZSF"|| Br_Code == "COKPR" ||
                        Br_Code == "MBISF" || Br_Code == "MBYAF"|| Br_Code == "DELSF" || Br_Code == "DELAF")
                    {
                        TotEMPYR = Lib.Conv2Decimal(Lib.NumericFormat((Lib.Convert2Decimal(TotSal.ToString()) * Lib.Convert2Decimal(EmpyrPer.ToString())).ToString(),2));
                    }
                    else
                        TotEMPYR = Math.Ceiling((Lib.Convert2Decimal(TotSal.ToString()) * Lib.Convert2Decimal(EmpyrPer.ToString())));
                    TotTOT = Lib.Convert2Decimal(TotEMPYE.ToString()) + Lib.Convert2Decimal(TotEMPYR.ToString());

                    if (Br_Code == "HOCPL" || Br_Code == "COKPR")
                    {
                        AddToList("HOSAL", Lib.Convert2Decimal(TotSal.ToString()));
                        AddToList("HOER", Lib.Convert2Decimal(TotEMPYR.ToString()));
                        AddToList("HOEY", Lib.Convert2Decimal(TotEMPYE.ToString()));
                        AddToList("HOTOT", Lib.Convert2Decimal(TotTOT.ToString()));
                    }
                    if (Br_Code == "CHNSF" || Br_Code == "CHNAF")
                    {
                        AddToList("CHNSAL", Lib.Convert2Decimal(TotSal.ToString()));
                        AddToList("CHNER", Lib.Convert2Decimal(TotEMPYR.ToString()));
                        AddToList("CHNEY", Lib.Convert2Decimal(TotEMPYE.ToString()));
                        AddToList("CHNTOT", Lib.Convert2Decimal(TotTOT.ToString()));
                    }
                    if (Br_Code == "DELSF" || Br_Code == "DELAF")
                    {
                        AddToList("DELSAL", Lib.Convert2Decimal(TotSal.ToString()));
                        AddToList("DELER", Lib.Convert2Decimal(TotEMPYR.ToString()));
                        AddToList("DELEY", Lib.Convert2Decimal(TotEMPYE.ToString()));
                        AddToList("DELTOT", Lib.Convert2Decimal(TotTOT.ToString()));
                    }
                    if (Br_Code == "MBISF" || Br_Code == "MBYAF")
                    {
                        AddToList("MUMSAL", Lib.Convert2Decimal(TotSal.ToString()));
                        AddToList("MUMER", Lib.Convert2Decimal(TotEMPYR.ToString()));
                        AddToList("MUMEY", Lib.Convert2Decimal(TotEMPYE.ToString()));
                        AddToList("MUMTOT", Lib.Convert2Decimal(TotTOT.ToString()));
                    }
                    if (Br_Code == "COKSF" || Br_Code == "COKAF"||Br_Code=="SEZSF")
                    {
                        AddToList("COKSAL", Lib.Convert2Decimal(TotSal.ToString()));
                        AddToList("COKER", Lib.Convert2Decimal(TotEMPYR.ToString()));
                        AddToList("COKEY", Lib.Convert2Decimal(TotEMPYE.ToString()));
                        AddToList("COKTOT", Lib.Convert2Decimal(TotTOT.ToString()));
                    }
                }
                else
                {
                    TotEMPYR = Math.Ceiling((Lib.Convert2Decimal(TotSal.ToString()) * (Lib.Convert2Decimal(Dt_List.Rows[0]["SAL_ESI_EMPLR_PER"].ToString()) / 100)));
                    TotTOT = Lib.Convert2Decimal(TotEMPYE.ToString()) + Lib.Convert2Decimal(TotEMPYR.ToString());
                }

                mRow = new HrReport();
                mRow.row_type = "TOTAL";
                mRow.row_colour = "RED";
                mRow.emp_name = "TOTAL";
                mRow.sal_gross_earn = Lib.Convert2Decimal(TotSal.ToString());
                mRow.emply_esi = Lib.Convert2Decimal(TotEMPYE.ToString());
                mRow.emplr_esi = Lib.Convert2Decimal(TotEMPYR.ToString());
                mRow.total = Lib.Convert2Decimal(Lib.NumericFormat(TotTOT.ToString(), 2));
                mRow.branch = "";
                mList.Add(mRow);
            }

            if (type == "EXCEL" && CanPrint)
            {
                PrintESI();
            }
        }
        public string AllValid(Salarym Record)
        {
            string str = "";
            DateTime tdate = DateTime.Now;
            try
            {
                
                //Boolean bRet = true;
                //DataTable Dt_locked = new DataTable();
                //string sql = "";
                //sql = " select rec_locked from salarym where ";
                //sql += " sal_pkid='" + drow["SAL_PKID"].ToString() + "'";
                //Dt_locked = orConnection.RunSql(sql);
                //if (Dt_locked.Rows.Count > 0)
                //    if (Dt_locked.Rows[0]["REC_LOCKED"].ToString().Trim() == "Y")
                //    {
                //        bRet = false;
                //        MessageBox.Show("Details Closed, Can't Edit", "Payroll");
                //        return bRet;
                //    }
                //return bRet;



                //if (Record.sal_code.Trim().Length <= 0)
                //    Lib.AddError(ref str, " | Code Cannot Be Empty");

                //if (Record.sal_code.Trim().Length > 0)
                //{

                //    sql = "select sal_pkid from (";
                //    sql += "select sal_pkid  from salaryheadm a where (a.sal_code = '{CODE}')  ";
                //    sql += ") a where sal_pkid <> '{PKID}'";

                //    sql = sql.Replace("{CODE}", Record.sal_code);
                //    sql = sql.Replace("{PKID}", Record.sal_pkid);

                //    if (Con_Oracle.IsRowExists(sql))
                //        Lib.AddError(ref str, " | Code Exists");
                //}

                //if (Record.sal_desc.Trim().Length > 0)
                //{

                //    sql = "select sal_pkid from (";
                //    sql += "select sal_pkid  from salaryheadm a where (a.sal_desc = '{NAME}')  ";
                //    sql += ") a where sal_pkid <> '{PKID}'";

                //    sql = sql.Replace("{NAME}", Record.sal_desc);
                //    sql = sql.Replace("{PKID}", Record.sal_pkid);


                //    if (Con_Oracle.IsRowExists(sql))

                //        Lib.AddError(ref str, " | Description Exists");
                //}

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        private void LoadNewEmp(string Br_Code)
        {
            DataTable Dt_Temp;
            DataRow Dr_Target = null;

            string sql = "";
            string Emp_Ids = "";
            int cMonth, pMonth;
            int cYear, pYear;
            cMonth = salmonth;
            cYear = salyear;

            if (cMonth == 1)
            {
                pMonth = 12;
                pYear = cYear - 1;
            }
            else
            {
                pMonth = cMonth - 1;
                pYear = cYear;
            }

            Emp_Ids = GetEmpIds(Br_Code, pMonth, pYear);
            if (Emp_Ids != "")
            {
                Emp_Ids = Emp_Ids.Replace(",", "','");

                sql = " select emp_name,emp_no from empm where emp_pkid in ( ";
                sql += " select sal_emp_id from salarym ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + Br_Code + "'";
                sql += " and sal_month = " + cMonth.ToString();
                sql += " and sal_year = " + cYear.ToString();
                if (empstatus != "BOTH")
                    sql += " and rec_category = '" + empstatus.ToString() + "'";
                sql += " and sal_emp_id not in ('" + Emp_Ids + "') )";
                
                Dt_Temp = new DataTable();
                Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow dr in Dt_Temp.Rows)
                {
                    Dr_Target = Dt_EMP.NewRow();
                    Dr_Target["BRANCH"] = Br_Code;
                    Dr_Target["EMP_NAME"] = dr["EMP_NAME"];
                    Dr_Target["EMP_NO"] = dr["EMP_NO"];
                    Dr_Target["EMP_STATUS"] = "JOINING";
                    Dt_EMP.Rows.Add(Dr_Target);
                }
            }

            Emp_Ids = GetEmpIds(Br_Code, cMonth, cYear);
            if (Emp_Ids != "")
            {
                Emp_Ids = Emp_Ids.Replace(",", "','");

                sql = " select emp_name,emp_no from empm where emp_pkid in ( ";
                sql += " select sal_emp_id from salarym ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + Br_Code + "'";
                sql += " and sal_month = " + pMonth.ToString();
                sql += " and sal_year = " + pYear.ToString();
                if (empstatus != "BOTH")
                    sql += " and rec_category = '" + empstatus.ToString() + "'";
                sql += " and sal_emp_id not in ('" + Emp_Ids + "') )";

                Dt_Temp = new DataTable();
                Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow dr in Dt_Temp.Rows)
                {
                    Dr_Target = Dt_EMP.NewRow();
                    Dr_Target["BRANCH"] = Br_Code;
                    Dr_Target["EMP_NAME"] = dr["EMP_NAME"];
                    Dr_Target["EMP_NO"] = dr["EMP_NO"];
                    Dr_Target["EMP_STATUS"] = "LEAVING";
                    Dt_EMP.Rows.Add(Dr_Target);
                }
            }
        }

        private string GetEmpIds(string Br_code, int sMonth, int sYear)
        {
            string str = "";
            sql = " select sal_emp_id from salarym ";
            sql += " where rec_company_code ='" + company_code + "' ";
            sql += " and rec_branch_code ='" + Br_code + "' ";
            sql += " and sal_month =" + sMonth;
            sql += " and sal_year =" + sYear;

            DataTable Dt_Sal = new DataTable();
            Dt_Sal = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow dr in Dt_Sal.Rows)
            {
                if (str != "")
                    str += ",";
                str += dr["sal_emp_id"].ToString();
            }

            return str;
        }
        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;
            LovService lovservice = new LovService();

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //parameter.Add("comp_code", comp_code);
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }

        private void ReadCompanyDetails()
        {
            comp_name = ""; comp_add1 = ""; comp_add2 = ""; comp_add3 = "";
            Dictionary<string, object> mSearchData = new Dictionary<string, object>();
            LovService mService = new LovService();
            mSearchData.Add("table", "ADDRESS");
            mSearchData.Add("branch_code", branch_code);
            DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
            if (Dt_CompAddress != null)
            {
                foreach (DataRow Dr in Dt_CompAddress.Rows)
                {
                    comp_name = Dr["COMP_NAME"].ToString();
                    comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                    comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                    comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                    //comp_tel = Dr["COMP_TEL"].ToString();
                    //comp_fax = Dr["COMP_FAX"].ToString();
                    //comp_web = Dr["COMP_WEB"].ToString();
                    //comp_email = Dr["COMP_EMAIL"].ToString();
                    //comp_cinno = Dr["COMP_CINNO"].ToString();
                    //comp_gstin = Dr["COMP_GSTIN"].ToString();
                    break;
                }
            }
        }


        private void PrintEPF()
        {
            string fname = "myreport";
            fname = "EPF-" + branch_code + "-" + new DateTime(salyear, salmonth, 01).ToString("MMMM").ToUpper() + "-" + salyear.ToString();
            if (fname.Length > 30)
                fname = fname.Substring(0, 30);
            File_Display_Name = Lib.ProperFileName(fname) + ".xls";
            File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
            File_Type = "xls";
            // ImagePath = report_folder + "\\Images";

            OpenFile();
            SetColumns();
            WriteHeadingPF();
            FillDataPF();
            file.SaveXls(File_Name);
        }

        private void PrintESI()
        {
            string fname = "myreport";
            fname = "ESI-" + branch_code + "-" + new DateTime(salyear, salmonth, 01).ToString("MMMM").ToUpper() + "-" + salyear.ToString();
            if (fname.Length > 30)
                fname = fname.Substring(0, 30);
            File_Display_Name = Lib.ProperFileName(fname) + ".xls";
            File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
            File_Type = "xls";
          //  ImagePath = report_folder + "\\Images\\Logo.gif";

            OpenFile();
            SetColumns();  
            WriteHeadingESI();
            FillDataESI();

            if (CopyFromRow > 0 && CopyToRow > 0 && reporttype == "ESI-NORTH")
            {
                ws.Cells.GetSubrangeRelative(CopyFromRow, 0, 12, CopyToRow - CopyFromRow).CopyTo(ws2, 8, 0);
                ws2.Rows[8 + (CopyToRow - CopyFromRow) + 2].InsertCopy(1, ws.Rows[CopyDelhiTotRow]);
                for (int c = 0, sl = 0; c < CopyToRow - CopyFromRow; c++)
                {
                    if (ws2.Cells[c + 10, 1].Value == null)
                        continue;
                    if (ws2.Cells[c + 10, 1].Value.ToString().Trim().Length <= 0)
                        continue;
                    if (Char.IsNumber(ws2.Cells[c + 10, 1].Value.ToString(), 0))
                    {
                        sl++;
                        ws2.Cells[c + 10, 1].Value = sl;
                    }
                }

                try
                {
                    Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                    LovService mService = new LovService();
                    mSearchData.Add("table", "ADDRESS");
                    mSearchData.Add("branch_code", "DELSF");
                    DataTable Dt_BrAddr = mService.Search2Datatable(mSearchData);//Del Separate sheet
                    if (Dt_BrAddr != null)
                        if (Dt_BrAddr.Rows.Count > 0)
                        {
                            ws2.Cells[0, 1].Value = Dt_BrAddr.Rows[0]["COMP_NAME"].ToString();
                            ws2.Cells[1, 1].Value = Dt_BrAddr.Rows[0]["COMP_ADDRESS1"].ToString();
                            ws2.Cells[2, 1].Value = Dt_BrAddr.Rows[0]["COMP_ADDRESS2"].ToString() + ", " + Dt_BrAddr.Rows[0]["COMP_ADDRESS3"].ToString();
                        }
                }
                catch (Exception)
                {
                }
            }
            file.SaveXls(File_Name);
        }
        private void OpenFile()
        {
            file = new ExcelFile();
            file.Worksheets.Add("Report");
            ws = file.Worksheets["Report"];
            ws.PrintOptions.FitWorksheetWidthToPages = 1;
        }

        private void SetColumns()
        {

            iCol = 2;
            ws.Columns[0].Width = 255 * 0;
            ws.Columns[1].Width = 255 * 4;
            ws.Columns[iCol + 0].Width = 255 * 5;
            ws.Columns[iCol + 1].Width = 255 * 12;
            ws.Columns[iCol + 2].Width = 255 * 20;
            ws.Columns[iCol + 3].Width = 255 * 7;//11
            ws.Columns[iCol + 4].Width = 255 * 7;//9
            ws.Columns[iCol + 5].Width = 255 * 9;//9,11
            ws.Columns[iCol + 6].Width = 255 * 8;//11
            ws.Columns[iCol + 7].Width = 255 * 5;//11
            ws.Columns[iCol + 8].Width = 255 * 7;
            ws.Columns[iCol + 9].Width = 255 * 6;
            ws.Columns[iCol + 10].Width = 255 * 10;//9
            for (int s = 0; s < 13; s++)
            {
                ws.Columns[s].Style.Font.Name = "Arial";
                ws.Columns[s].Style.Font.Size = 8 * 20;
            }
            //ws.Columns[iCol - 1].Style.Font.Size = 9 * 20;
            ws.Columns[iCol + 1].Style.Font.Size = 7 * 20;
            ws.Columns[iCol + 2].Style.Font.Size = 7 * 20;
            ws.Columns[iCol + 5].Style.NumberFormat = "#0.00";
            ws.Columns[iCol + 6].Style.NumberFormat = "#0.00";
            ws.Columns[iCol + 8].Style.NumberFormat = "#0.00";
            ws.Columns[iCol + 9].Style.NumberFormat = "#0.00";
            ws.Columns[iCol + 10].Style.NumberFormat = "#0.00";
        }
        private void WriteHeadingPF()
        {
            string str = "";
            ReadCompanyDetails();
            WriteData(0, iCol - 1, comp_name, true);
            WriteData(1, iCol - 1, comp_add1, true);
            str = comp_add2;
            if (str.Trim() != "" && comp_add3.Trim() != "")
                str += ",";
            str += comp_add3;
            WriteData(2, iCol - 1, str, true);
            Merge_Cell(4, iCol - 1, 12, 1);

            str = " P.F. DEDUCTIONS ";
            //if (CmbType.Text == "ARREARS_EPF_SOUTH" || CmbType.Text == "ARREARS_EPF_NORTH")
            //    str += "ON ARREARS ";
            //if (CmbType.Text == "CONSOL_EPF_SOUTH" || CmbType.Text == "CONSOL_EPF_NORTH" || CmbType.Text == "ARREARS_EPF_SOUTH" || CmbType.Text == "ARREARS_EPF_NORTH")
            //{
            //    str += "FROM " + Dt_From.Value.ToString("MMMM").ToUpper() + " " + Dt_From.Value.Year.ToString();
            //    str += " TO " + Dt_To.Value.ToString("MMMM").ToUpper() + " " + Dt_To.Value.Year.ToString();
            //}
            //else
            //{
                if (salmonth > 0 && salyear > 0)
                    str +=  new DateTime(salyear,salmonth,01).ToString("MMMM").ToUpper() + ", " + salyear.ToString();
            //}
            
            WriteData(4, iCol - 1, str, true);
            ws.Cells[4, iCol - 1].Style.Font.Size = 9 * 20;
            ws.Cells[4, iCol - 1].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
            Merge_Cell(5, iCol - 1, 1, 3);
            WriteData(5, iCol - 1, "SI#", true, "ALL");
            Merge_Cell(5, iCol + 0, 1, 3);
            WriteData(5, iCol + 0, "CODE", true, "ALL");
            Merge_Cell(5, iCol + 1, 1, 3);
            WriteData(5, iCol + 1, "PF A/C NO", true, "ALL");
            Merge_Cell(5, iCol + 2, 1, 3);
            WriteData(5, iCol + 2, "NAME", true, "ALL");
            Merge_Cell(5, iCol + 3, 1, 3);
            WriteData(5, iCol + 3, "BASE SALARY", true, "ALL");
            Merge_Cell(5, iCol + 4, 1, 3);
            WriteData(5, iCol + 4, "EPF DEDUCTS", true, "ALL");
            Merge_Cell(5, iCol + 5, 1, 3);
            WriteData(5, iCol + 5, "EMPLOYER'S SHARE", true, "ALL");
            Merge_Cell(5, iCol + 6, 1, 3);
            WriteData(5, iCol + 6, "PENSION FUND", true, "ALL");
            Merge_Cell(5, iCol + 7, 1, 3);
            WriteData(5, iCol + 7, "VPF", true, "ALL");
            Merge_Cell(5, iCol + 8, 1, 3);
            WriteData(5, iCol + 8, "ADMN. CHRGS", true, "ALL");
            Merge_Cell(5, iCol + 9, 1, 3);
            WriteData(5, iCol + 9, "EDLI. CONTRIB.", true, "ALL"); 
            Merge_Cell(5, iCol + 10, 1, 3);
            WriteData(5, iCol + 10, "TOTAL", true, "ALL");
        }
        private void FillDataPF()
        {
            int dRow = 8;
            int SiNo = 0;
            string PreBranch = "";
            decimal TotBaseSal = 0; decimal TotEpfdedut = 0; decimal TotEmplrShare = 0; decimal TotVpfund = 0;
            decimal TotPenfund = 0; decimal TotAdminchrg = 0; decimal TotEdlichrg = 0; decimal TotTotal = 0;
            foreach (HrReport Rec in mList)
            {
                if (IsConsol)
                    if (PreBranch != Rec.branch.ToString().Trim())
                    {
                        PreBranch = Rec.branch.ToString().Trim();
                        if (PreBranch != "")
                        {
                            Merge_Cell(dRow, iCol - 1, 12, 2);
                            ws.Cells[dRow, iCol - 1].Style.Font.Size = 9 * 20;
                            WriteData(dRow, iCol - 1, Rec.branch, true, "ALL");
                            dRow = dRow + 2;
                            WriteData(dRow, iCol + 8, "", false, "R_FORMAT");
                            WriteData(dRow, iCol + 9, "", false, "R_FORMAT");
                            WriteData(dRow, iCol + 10, "", false, "R_FORMAT");
                        }
                    }

                if (Rec.emp_name != "TOTAL")
                {
                    SiNo++;
                    ws.Cells[dRow, iCol - 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    WriteData(dRow, iCol - 1, SiNo, false, "ALL");
                }
                else
                    WriteData(dRow, iCol - 1, "", false, "ALL");
                WriteData(dRow, iCol, Rec.emp_no, false, "ALL");
                WriteData(dRow, iCol + 1, Rec.emp_pfno, false, "ALL");
                WriteData(dRow, iCol + 2, Rec.emp_name, false, "ALL");
                WriteData(dRow, iCol + 3, Rec.pf_base_salary, false, "ALL");
                WriteData(dRow, iCol + 4, Rec.pf_deduction, false, "ALL");
                WriteData(dRow, iCol + 5, Rec.emplyr_share, false, "ALL");
                WriteData(dRow, iCol + 6, Rec.pension, false, "ALL");
                WriteData(dRow, iCol + 7, Rec.vpf, false, "ALL");
                if (Rec.emp_name == "TOTAL")
                {
                    //TotBaseSal += Common.Convert2Decimal(dr["BASIC_DA"].ToString());
                    //TotEpfdedut += Common.Convert2Decimal(dr["DEDUCTN"].ToString());
                    //TotEmplrShare += Common.Convert2Decimal(dr["EMPR_SHARE"].ToString());
                    //TotPenfund += Common.Convert2Decimal(dr["PENSION"].ToString());
                    //TotVpfund += Common.Convert2Decimal(dr["VPF"].ToString());
                    //TotAdminchrg += Common.Convert2Decimal(dr["ADMIN_CHRG"].ToString());
                    //TotEdlichrg += Common.Convert2Decimal(dr["EDLI_CHRG"].ToString());
                    //TotTotal += Common.Convert2Decimal(dr["TOTAL"].ToString());

                    TotBaseSal += Lib.Convert2Decimal(Rec.pf_base_salary.ToString());
                    TotEpfdedut += Lib.Convert2Decimal(Rec.pf_deduction.ToString());
                    TotEmplrShare += Lib.Convert2Decimal(Rec.emplyr_share.ToString());
                    TotPenfund += Lib.Convert2Decimal(Rec.pension.ToString());
                    TotVpfund += Lib.Convert2Decimal(Rec.vpf.ToString());
                    TotAdminchrg += Lib.Convert2Decimal(Rec.admin_chrg.ToString());
                    TotEdlichrg += Lib.Convert2Decimal(Rec.edli_chrg.ToString());
                    TotTotal += Lib.Convert2Decimal(Rec.total_chrg.ToString());

                    WriteData(dRow, iCol + 8, Rec.admin_chrg, false, "ALL");
                    WriteData(dRow, iCol + 9, Rec.edli_chrg, false, "ALL");
                    WriteData(dRow, iCol + 10, Rec.total_chrg, false, "ALL");
                    ws.Rows[dRow].Style.Font.Weight = ExcelFont.BoldWeight;
                }
                else
                {
                    if (Rec.edli_based_on == "EMPLOYEE")//PF NORTH old cndtn  branch_region == "NORTH" || reporttype.Contains("NORTH")
                    {
                        WriteData(dRow, iCol + 8, "", false, "ALL");
                        WriteData(dRow, iCol + 9, Rec.edli_chrg, false, "ALL");
                        WriteData(dRow, iCol + 10, "", false, "ALL");
                    }
                    else
                    {
                        WriteData(dRow, iCol + 8, "", false, "R_FORMAT");
                        WriteData(dRow, iCol + 9, "", false, "R_FORMAT");
                        WriteData(dRow, iCol + 10, "", false, "R_FORMAT");
                    }
                }
                //if (CanPrintEPS)
                //    WriteData(dRow, iCol + 11, dr["EPS_AMT"], false, "ALL");
                dRow++;
            }
            if (IsConsol)
            {
                dRow++;
                WriteData(dRow, iCol + 1, "SUMMARY", true, "");
                dRow++;
                WriteData(dRow, iCol + 0, "", true, "ALL");
                WriteData(dRow, iCol + 1, "", true, "ALL");
                WriteData(dRow, iCol + 2, "GRAND TOTAL", true, "ALL");
                WriteData(dRow, iCol + 3, TotBaseSal, true, "ALL");
                WriteData(dRow, iCol + 4, TotEpfdedut, true, "ALL");
                WriteData(dRow, iCol + 5, TotEmplrShare, true, "ALL");
                WriteData(dRow, iCol + 6, TotPenfund, true, "ALL");
                WriteData(dRow, iCol + 7, TotVpfund, true, "ALL");


                decimal EdliTolerance = 0;
                decimal AdminTolerance = 0;

                AdminTolerance = (TotAdminchrg > 0 && TotAdminchrg < 500) ? 500 - TotAdminchrg : 0;
                // EdliTolerance = (TotEdlichrg > 0 && TotEdlichrg < 200) ? 200 - TotEdlichrg : 0;

                EdliTolerance = 0;

                TotAdminchrg = TotAdminchrg + AdminTolerance;
                TotEdlichrg = TotEdlichrg + EdliTolerance;
                TotTotal = TotTotal + EdliTolerance + AdminTolerance;

                TotAdminchrg = Lib.Convert2Decimal(Lib.NumericFormat(TotAdminchrg.ToString(), 2));
                //ws.Cells[dRow, iCol + 8].Style.NumberFormat = "#0";
                WriteData(dRow, iCol + 8, TotAdminchrg, true, "ALL");
                TotEdlichrg = Lib.Convert2Decimal(Lib.NumericFormat(TotEdlichrg.ToString(), 2));
                //ws.Cells[dRow, iCol + 9].Style.NumberFormat = "#0";
                WriteData(dRow, iCol + 9, TotEdlichrg, true, "ALL");
                TotTotal = Lib.Convert2Decimal(Lib.NumericFormat(TotTotal.ToString(), 2));
                //ws.Cells[dRow, iCol + 10].Style.NumberFormat = "#0";
                WriteData(dRow, iCol + 10, TotTotal, true, "ALL");

                dRow += 3;
                WriteData(dRow, iCol + 2, "NAME", true, "ALL");
                WriteData(dRow, iCol + 3, "STATUS", true, "ALL");
                WriteData(dRow++, iCol + 4, "BRANCH", true, "ALL");
                foreach (DataRow Dr in Dt_EMP.Select("1=1", "EMP_NO"))
                {
                    WriteData(dRow, iCol + 2, Dr["EMP_NAME"], false, "ALL");
                    WriteData(dRow, iCol + 3, Dr["EMP_STATUS"], false, "ALL");
                    WriteData(dRow++, iCol + 4, Dr["BRANCH"], false, "ALL");
                }
            }
            //CanPrintEPS = false;
        }

        private void WriteHeadingESI()
        {
            string str = "";
            ReadCompanyDetails();
            WriteData(0, iCol - 1, comp_name, true);
            WriteData(1, iCol - 1, comp_add1, true);
            str = comp_add2;
            if (str.Trim() != "" && comp_add3.Trim() != "")
                str += ",";
            str += comp_add3;
            WriteData(2, iCol - 1, str, true);
            Merge_Cell(4, iCol - 1, 8, 1);

            str = " ESI REMITTANCE ";
            //if (CmbType.Text == "ARREARS_ESI")
            //{
            //    str += "ON ARREARS ";
            //    str += "FROM " + Dt_From.Value.ToString("MMMM").ToUpper() + " " + Dt_From.Value.Year.ToString();
            //    str += " TO " + Dt_To.Value.ToString("MMMM").ToUpper() + " " + Dt_To.Value.Year.ToString();
            //}
            //if (Common.Convert2Decimal(Txt_Month.Text) > 0 && Common.Convert2Decimal(Txt_Year.Text) > 0)
            //    str += Convert.ToDateTime("01/" + Txt_Month.Text + "/" + Txt_Year.Text).ToString("MMMM").ToUpper() + ", " + Txt_Year.Text;

            if (salmonth > 0 && salyear > 0)
                str += new DateTime(salyear, salmonth, 01).ToString("MMMM").ToUpper() + ", " + salyear.ToString();

            WriteData(4, iCol - 1, str, true);
            ws.Cells[4, iCol - 1].Style.Font.Size = 9 * 20;
            ws.Cells[4, iCol - 1].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
            Merge_Cell(5, iCol - 1, 1, 3);
            WriteData(5, iCol - 1, "SI#", true, "ALL");
            Merge_Cell(5, iCol + 0, 1, 3);
            WriteData(5, iCol + 0, "CODE", true, "ALL");
            Merge_Cell(5, iCol + 1, 1, 3);
            WriteData(5, iCol + 1, "INSURANCE NO.", true, "ALL");
            Merge_Cell(5, iCol + 2, 1, 3);
            WriteData(5, iCol + 2, "NAME", true, "ALL");
            Merge_Cell(5, iCol + 3, 1, 3);
            WriteData(5, iCol + 3, "SALARY", true, "ALL");
            Merge_Cell(5, iCol + 4, 1, 3);
            WriteData(5, iCol + 4, "EMPLOYEE'S SHARE", true, "ALL");
            Merge_Cell(5, iCol + 5, 1, 3);
            WriteData(5, iCol + 5, "EMPLOYER'S SHARE", true, "ALL");
            Merge_Cell(5, iCol + 6, 1, 3);
            WriteData(5, iCol + 6, "TOTAL", true, "ALL");
            if (reporttype == "ESI-NORTH")
                ws2 = file.Worksheets.AddCopy("Report2", ws);
        }
        private void FillDataESI()
        {
            int dRow = 8;
            int SiNo = 0;
            decimal TotSal = 0, TotEmplyeShare = 0, TotEmplrShare = 0, TotTotal = 0;
            string PreBranch = "", Br_Name = "";
            CopyFromRow = 0; CopyToRow = 0;
            foreach (HrReport Rec in mList)
            {
                if (IsConsol)
                    if (PreBranch != Rec.branch.ToString().Trim())
                    {
                        PreBranch = Rec.branch.ToString().Trim();
                        if (PreBranch != "")
                        {
                            if (PreBranch.StartsWith("DEL") && CopyFromRow <= 0)
                                CopyFromRow = dRow;

                            if (!PreBranch.StartsWith("DEL") && CopyFromRow > 0 && CopyToRow <= 0)
                                CopyToRow = dRow;

                            Br_Name = PreBranch;
                            Merge_Cell(dRow, iCol - 1, 8, 2);
                            ws.Cells[dRow, iCol - 1].Style.Font.Size = 9 * 20;
                            WriteData(dRow, iCol - 1, Rec.branch, true, "ALL");
                            dRow = dRow + 2;
                        }
                    }

                if (Rec.emp_name != "TOTAL")
                {
                    SiNo++;
                    ws.Cells[dRow, iCol - 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    WriteData(dRow, iCol - 1, SiNo, false, "ALL");
                }
                else
                    WriteData(dRow, iCol - 1, "", false, "ALL");
                WriteData(dRow, iCol + 0, Rec.emp_no, false, "ALL");
                WriteData(dRow, iCol + 1, Rec.emp_esino, false, "ALL");
                WriteData(dRow, iCol + 2, Rec.emp_name, false, "ALL");
                WriteData(dRow, iCol + 3, Rec.sal_gross_earn, false, "ALL");
                WriteData(dRow, iCol + 4, Rec.emply_esi, false, "ALL");
                //if (IsBeforeDec2010 || dr["EMPLR_ESI"].ToString().Trim()!="")
                //    WriteData(dRow, iCol + 5, dr["EMPLR_ESI"], false, "ALL");

                if (Rec.emplr_esi != null)
                    WriteData(dRow, iCol + 5, Rec.emplr_esi, false, "ALL");
                else
                    WriteData(dRow, iCol + 5, Rec.emplr_esi, false, "R_FORMAT");

                WriteData(dRow, iCol + 6, "", false, "R_FORMAT");

                if (Rec.emp_name == "TOTAL")
                {
                    TotSal += Lib.Convert2Decimal(Rec.sal_gross_earn.ToString());
                    TotEmplyeShare += Lib.Convert2Decimal(Rec.emply_esi.ToString());
                    if (Br_Name == "CHNSF" || Br_Name == "CHNAF" || Br_Name == "COKSF" || Br_Name == "COKAF" || Br_Name == "HOCPL" || Br_Name == "SEZSF" || Br_Name == "COKPR")
                    {
                        TotEmplrShare += 0;
                        TotTotal += 0;
                    }
                    else
                    {
                        TotEmplrShare += Lib.Convert2Decimal(Rec.emplr_esi.ToString());
                        TotTotal += Lib.Convert2Decimal(Rec.total.ToString());
                    }
                    WriteData(dRow, iCol + 6, Rec.total, false, "ALL");
                    ws.Rows[dRow].Style.Font.Weight = ExcelFont.BoldWeight;
                }
                dRow++;
            }
            if (IsConsol)
            {
                dRow++;
                WriteData(dRow, iCol + 1, "SUMMARY", true, "");
                if (reporttype == "ESI-SOUTH")
                {
                    dRow++;
                    WriteData(dRow, iCol + 0, "", true, "ALL");
                    WriteData(dRow, iCol + 1, "CHENNAI", true, "ALL");
                    WriteData(dRow, iCol + 2, "TOTAL", true, "ALL");
                    if (myDict.ContainsKey("CHNSAL"))
                        WriteData(dRow, iCol + 3, myDict["CHNSAL"], true, "ALL");
                    if (myDict.ContainsKey("CHNEY"))
                        WriteData(dRow, iCol + 4, myDict["CHNEY"], true, "ALL");//employee
                    if (myDict.ContainsKey("CHNER"))
                        WriteData(dRow, iCol + 5, Math.Ceiling(myDict["CHNER"]), true, "ALL");//employer
                    if (myDict.ContainsKey("CHNTOT"))
                        WriteData(dRow, iCol + 6, Math.Ceiling(myDict["CHNTOT"]), true, "ALL");

                    dRow++;
                    WriteData(dRow, iCol + 0, "", true, "ALL");
                    WriteData(dRow, iCol + 1, "KOCHI", true, "ALL");
                    WriteData(dRow, iCol + 2, "TOTAL", true, "ALL");
                    if (myDict.ContainsKey("COKSAL"))
                        WriteData(dRow, iCol + 3, myDict["COKSAL"], true, "ALL");
                    if (myDict.ContainsKey("COKEY"))
                        WriteData(dRow, iCol + 4, myDict["COKEY"], true, "ALL");
                    if (myDict.ContainsKey("COKER"))
                        WriteData(dRow, iCol + 5, Math.Ceiling(myDict["COKER"]), true, "ALL");
                    if (myDict.ContainsKey("COKTOT"))
                        WriteData(dRow, iCol + 6, Math.Ceiling(myDict["COKTOT"]), true, "ALL");

                    dRow++;
                    WriteData(dRow, iCol + 0, "", true, "ALL");
                    WriteData(dRow, iCol + 1, "HO", true, "ALL");
                    WriteData(dRow, iCol + 2, "TOTAL", true, "ALL");
                    if (myDict.ContainsKey("HOSAL"))
                        WriteData(dRow, iCol + 3, myDict["HOSAL"], true, "ALL");
                    if (myDict.ContainsKey("HOEY"))
                        WriteData(dRow, iCol + 4, myDict["HOEY"], true, "ALL");
                    if (myDict.ContainsKey("HOER"))
                        WriteData(dRow, iCol + 5, Math.Ceiling(myDict["HOER"]), true, "ALL");
                    if (myDict.ContainsKey("HOTOT"))
                        WriteData(dRow, iCol + 6, Math.Ceiling(myDict["HOTOT"]), true, "ALL");

                }
                if (reporttype == "ESI-NORTH")
                {
                    dRow++;
                    CopyDelhiTotRow = dRow;
                    WriteData(dRow, iCol + 0, "", true, "ALL");
                    WriteData(dRow, iCol + 1, "DELHI", true, "ALL");
                    WriteData(dRow, iCol + 2, "TOTAL", true, "ALL");
                    if (myDict.ContainsKey("DELSAL"))
                        WriteData(dRow, iCol + 3, myDict["DELSAL"], true, "ALL");
                    if (myDict.ContainsKey("DELEY"))
                        WriteData(dRow, iCol + 4, myDict["DELEY"], true, "ALL");//employee
                    if (myDict.ContainsKey("DELER"))
                        WriteData(dRow, iCol + 5, Math.Ceiling(myDict["DELER"]), true, "ALL");//employer
                    if (myDict.ContainsKey("DELTOT"))
                        WriteData(dRow, iCol + 6, Math.Ceiling(myDict["DELTOT"]), true, "ALL");

                    dRow++;
                    WriteData(dRow, iCol + 0, "", true, "ALL");
                    WriteData(dRow, iCol + 1, "MUMBAI", true, "ALL");
                    WriteData(dRow, iCol + 2, "TOTAL", true, "ALL");
                    if (myDict.ContainsKey("MUMSAL"))
                        WriteData(dRow, iCol + 3, myDict["MUMSAL"], true, "ALL");
                    if (myDict.ContainsKey("MUMEY"))
                        WriteData(dRow, iCol + 4, myDict["MUMEY"], true, "ALL");//employee
                    if (myDict.ContainsKey("MUMER"))
                        WriteData(dRow, iCol + 5, Math.Ceiling(myDict["MUMER"]), true, "ALL");//employer
                    if (myDict.ContainsKey("MUMTOT"))
                        WriteData(dRow, iCol + 6, Math.Ceiling(myDict["MUMTOT"]), true, "ALL");
                }
                dRow++;
                WriteData(dRow, iCol + 0, "", true, "ALL");
                WriteData(dRow, iCol + 1, "ALL", true, "ALL");
                WriteData(dRow, iCol + 2, "GRAND TOTAL", true, "ALL");
                WriteData(dRow, iCol + 3, TotSal, true, "ALL");
                WriteData(dRow, iCol + 4, TotEmplyeShare, true, "ALL");
                if (myDict.ContainsKey("CHNER"))
                    TotEmplrShare += Math.Ceiling(myDict["CHNER"]);
                if (myDict.ContainsKey("COKER"))
                    TotEmplrShare += Math.Ceiling(myDict["COKER"]);
                if (myDict.ContainsKey("HOER"))
                    TotEmplrShare += Math.Ceiling(myDict["HOER"]);
                WriteData(dRow, iCol + 5, TotEmplrShare, true, "ALL");
                if (myDict.ContainsKey("CHNTOT"))
                    TotTotal += Math.Ceiling(myDict["CHNTOT"]);
                if (myDict.ContainsKey("COKTOT"))
                    TotTotal += Math.Ceiling(myDict["COKTOT"]);
                if (myDict.ContainsKey("HOTOT"))
                    TotTotal += Math.Ceiling(myDict["HOTOT"]);
                WriteData(dRow, iCol + 6, TotTotal, true, "ALL");
            }
        }
        private void WriteData(int _Row, int _Col, Object sData)
        {
            WriteData(_Row, _Col, sData, false, System.Drawing.Color.Black, "");
        }
        private void WriteData(int _Row, int _Col, Object sData, string BORDERS)
        {
            WriteData(_Row, _Col, sData, false, System.Drawing.Color.Black, BORDERS);
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold)
        {
            WriteData(_Row, _Col, sData, bBold, System.Drawing.Color.Black, "");
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold, string BORDERS, string Alignment = "LEFT")
        {
            WriteData(_Row, _Col, sData, bBold, System.Drawing.Color.Black, BORDERS, Alignment);
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold, System.Drawing.Color c, string BORDERS, string Alignment = "LEFT")
        {
            if (ws_active_sheetno == 2)
            {
                ws2.Cells[_Row, _Col].Value = sData;
                if (bBold)
                    ws2.Cells[_Row, _Col].Style.Font.Weight = ExcelFont.BoldWeight;
                ws2.Cells[_Row, _Col].Style.Font.Color = c;
                if (BORDERS == "ALL")
                    ws2.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
                else if (BORDERS == "NFORMAT")
                    ws2.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
                else if (BORDERS == "R_FORMAT")
                {
                    ws2.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Right, System.Drawing.Color.Black, LineStyle.Thin);
                }
                else if (BORDERS == "L_FORMAT")
                {
                    ws2.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Left, System.Drawing.Color.Black, LineStyle.Thin);
                }
                else if (BORDERS == "ALL_NFORMAT")
                {
                    ws2.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
                    ws2.Cells[_Row, _Col].Style.NumberFormat = "#,##0.00";
                }

                if (Alignment == "RIGHT")
                    ws2.Cells[_Row, _Col].Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
            }
            else
            {
                ws.Cells[_Row, _Col].Value = sData;
                if (bBold)
                    ws.Cells[_Row, _Col].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[_Row, _Col].Style.Font.Color = c;
                if (BORDERS == "ALL")
                    ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
                else if (BORDERS == "NFORMAT")
                    ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
                else if (BORDERS == "R_FORMAT")
                {
                    ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Right, System.Drawing.Color.Black, LineStyle.Thin);
                }
                else if (BORDERS == "L_FORMAT")
                {
                    ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Left, System.Drawing.Color.Black, LineStyle.Thin);
                }
                else if (BORDERS == "ALL_NFORMAT")
                {
                    ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
                    ws.Cells[_Row, _Col].Style.NumberFormat = "#,##0.00";
                }

                if (Alignment == "RIGHT")
                    ws.Cells[_Row, _Col].Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
            }
        }
        private void Merge_Cell(int _Row, int _Col, int _Width, int _Height)
        {
            myCell = ws.Cells.GetSubrangeRelative(_Row, _Col, _Width, _Height);
            myCell.Merged = true;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            myCell.Style.WrapText = true;
            myCell.Style.Font.Name = "Arial";
            myCell.Style.Font.Size = 8 * 20;
        }

        public IDictionary<string, object> ProcessLetter(Dictionary<string, object> SearchData)
        {
            sError = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            mList = new List<HrReport>();
            

            type = SearchData["type"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_region = "";
            if (SearchData.ContainsKey("branch_region"))
                branch_region = SearchData["branch_region"].ToString();
            report_folder = "";
            if (SearchData.ContainsKey("report_folder"))
                report_folder = SearchData["report_folder"].ToString();
            folderid = "";
            if (SearchData.ContainsKey("folderid"))
                folderid = SearchData["folderid"].ToString();

            uploadfileid = "";
            if (SearchData.ContainsKey("uploadfileid"))
                uploadfileid = SearchData["uploadfileid"].ToString();

            print_date = "";
            if (SearchData.ContainsKey("print_date"))
                print_date = SearchData["print_date"].ToString();

            effective_date = "";
            if (SearchData.ContainsKey("effective_date"))
                effective_date = SearchData["effective_date"].ToString();

            try
            {
     
                if (type == "DOWNLOAD-TEMPLATE")
                    printIncrementFormat();

                if (type == "PROCESS-LETTER")
                    ProcessIncrementLetter();

             
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", mList);
            RetData.Add("serror", sError);
            return RetData;
        }

        private void printIncrementFormat()
        {
            ws_active_sheetno = 0;
            string fName = "";
            string Branch_Location = "";
            try
            {
                sql = "select comp_location from companym where rec_company_code ='" + company_code + "' and rec_branch_code = '" + branch_code + "'";
                Object sVal = Con_Oracle.ExecuteScalar(sql);
                Branch_Location = sVal.ToString();

                sql = " select EMP_NO,EMP_NAME,EMP_GENDER,c.comp_name as EMP_COMPANY,SAL_GROSS_EARN,d.param_name as EMP_DESIGNATION ";
                sql += "   from salarym a";
                sql += "   inner join  empm b on a.sal_emp_id = b.emp_pkid ";
                sql += "   inner join companym c on a.rec_company_code=c.rec_company_code and c.comp_type='C'";
                sql += "   left join param d on b.emp_designation_id = d.param_pkid ";
                sql += "   where a.rec_company_code = '" + company_code + "'";
                sql += "   and a.rec_branch_code = '" + branch_code + "'";
                sql += "   and sal_month = sal_year ";
                sql += "   and nvl(emp_in_payroll,'N') = 'Y' ";
                sql += "   order by emp_no ";
     
                DataTable Dt_Parent = new DataTable();
                Dt_Parent = Con_Oracle.ExecuteQuery(sql);

                ExcelFile file = new ExcelFile();
                ExcelWorksheet ws = file.Worksheets.Add("Sheet1");
                int iRow = 0;
                iCol = 0;
                ws.Columns[0].Width = 255 * 6;
                ws.Columns[1].Width = 255 * 10;
                ws.Columns[2].Width = 255 * 30;
                ws.Columns[3].Width = 255 * 10;
                ws.Columns[4].Width = 255 * 20;
                ws.Columns[5].Width = 255 * 12;
                ws.Columns[6].Width = 255 * 10;
                ws.Columns[7].Width = 255 * 12;
                ws.Columns[8].Width = 255 * 12;
                ws.Columns[9].Width = 255 * 20;
                ws.Columns[10].Width = 255 * 20;
                ws.Columns[11].Width = 255 * 12;
                ws.Columns[12].Width = 255 * 12;
                ws.Columns[13].Width = 255 * 12;
                ws.Columns[14].Width = 255 * 12;
                ws.Columns[15].Width = 255 * 12;
                ws.Columns[16].Width = 255 * 12;
                ws.Columns[17].Width = 255 * 12;
                ws.Columns[18].Width = 255 * 12;
                ws.Columns[19].Width = 255 * 12;
                ws.Columns[20].Width = 255 * 12;
                ws.Columns[21].Width = 255 * 12;
                ws.Columns[22].Width = 255 * 12;

                ws.Cells[iRow, 0].Value = "SL.NO"; MergeCell(ws, iRow, 0, 1, 3);
                ws.Cells[iRow, 1].Value = "EMP CODE"; MergeCell(ws, iRow, 1, 1, 3);
                ws.Cells[iRow, 2].Value = "EMP NAME"; MergeCell(ws, iRow, 2, 1, 3);
                ws.Cells[iRow, 3].Value = "GENDER"; MergeCell(ws, iRow, 3, 1, 3);
                ws.Cells[iRow, 4].Value = "EMP COMPANY"; MergeCell(ws, iRow, 4, 1, 3);
                ws.Cells[iRow, 5].Value = "EMP LOCATION"; MergeCell(ws, iRow, 5, 1, 3);
                ws.Cells[iRow, 6].Value = "PRESENT SALARY"; MergeCell(ws, iRow, 6, 1, 3);
                ws.Cells[iRow, 7].Value = "INCREMENT"; MergeCell(ws, iRow, 7, 1, 3);
                ws.Cells[iRow, 8].Value = "SALARY AFTER INCREMENT"; MergeCell(ws, iRow, 8, 1, 3);
                ws.Cells[iRow, 9].Value = "DESIGNATION"; MergeCell(ws, iRow, 9, 1, 3);
                ws.Cells[iRow, 10].Value = "PROMOTION"; MergeCell(ws, iRow, 10, 1, 3);
                ws.Cells[iRow, 11].Value = "EMP DEPT."; MergeCell(ws, iRow, 11, 1, 3);
                ws.Cells[iRow, 12].Value = "Consolidated Basic"; MergeCell(ws, iRow, 12, 1, 3);
                ws.Cells[iRow, 13].Value = "Basic"; MergeCell(ws, iRow, 13, 1, 3);
                ws.Cells[iRow, 14].Value = "HRA"; MergeCell(ws, iRow, 14, 1, 3);
                ws.Cells[iRow, 15].Value = "CCA"; MergeCell(ws, iRow, 15, 1, 3);
                ws.Cells[iRow, 16].Value = "Conveyance"; MergeCell(ws, iRow, 16, 1, 3);
                ws.Cells[iRow, 17].Value = "Uniform & Washing Allow:"; MergeCell(ws, iRow, 17, 1, 3);
                ws.Cells[iRow, 18].Value = "Medical Allowance"; MergeCell(ws, iRow, 18, 1, 3);
                ws.Cells[iRow, 19].Value = "Education Allowance"; MergeCell(ws, iRow, 19, 1, 3);
                ws.Cells[iRow, 20].Value = "Entertainment Allowance"; MergeCell(ws, iRow, 20, 1, 3);
                ws.Cells[iRow, 21].Value = "Other Allowances"; MergeCell(ws, iRow, 21, 1, 3);
                ws.Cells[iRow, 22].Value = "Total"; MergeCell(ws, iRow, 22, 1, 3);

                for (int i = 0; i < 23; i++)
                    ws.Cells[iRow, i].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
                iRow++; iRow++;
                foreach (DataRow row in Dt_Parent.Rows)
                {
                    iRow++;
                    ws.Cells[iRow, 0].Value = iRow - 2;
                    ws.Cells[iRow, 1].Value = row["EMP_NO"];
                    ws.Cells[iRow, 2].Value = row["EMP_NAME"];
                    ws.Cells[iRow, 3].Value = (row["EMP_GENDER"].ToString() == "M") ? "MALE" : "FEMALE";
                    ws.Cells[iRow, 4].Value = row["EMP_COMPANY"];
                    ws.Cells[iRow, 5].Value = Branch_Location;
                    ws.Cells[iRow, 6].Value = row["SAL_GROSS_EARN"];
                    ws.Cells[iRow, 8].Formula = "=SUM(G" + (iRow + 1).ToString() + ", H" + (iRow + 1).ToString() + ")";
                    ws.Cells[iRow, 9].Value = row["EMP_DESIGNATION"];
                    ws.Cells[iRow, 11].Value = null;
                    ws.Cells[iRow, 22].Formula = "=SUM(M" + (iRow + 1).ToString() + ":V" + (iRow + 1).ToString() + ")";
                    ws.Rows[iRow].Style.Font.Size = 9 * 20;
                    for (int i = 0; i < 23; i++)
                        ws.Cells[iRow, i].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
                }
             
                fName = "EMPINCMNTRPT-" + branch_code;
                if (fName.Length > 30)
                    fName = fName.Substring(0, 30);
                File_Display_Name = Lib.ProperFileName(fName) + ".xls";
                File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
                File_Type = "xls";

                file.SaveXls(File_Name);

                Dt_Parent.Clear();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        private void MergeCell(ExcelWorksheet ws, int cRow, int cCol, int cWidth, int cHeight)
        {
            myCell = ws.Cells.GetSubrangeRelative(cRow, cCol, cWidth, cHeight);
            myCell.Merged = true;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            myCell.Style.WrapText = true;
            myCell.Style.Font.Name = "Arial";
            myCell.Style.Font.Size = 9 * 20;
        }


        private void ProcessIncrementLetter()
        {

            if (uploadfileid.Trim() == "")
                return;

            ws_active_sheetno = 2;

            string[] ArrDate = null;
            if (print_date.Trim() != "")
            {
                ArrDate = print_date.Split('-');
                DTprint = new DateTime(Lib.Conv2Integer(ArrDate[0]), Lib.Conv2Integer(ArrDate[1]), Lib.Conv2Integer(ArrDate[2]));
            }
            if (effective_date.Trim() != "")
            {
                ArrDate = effective_date.Split('-');
                DTEffective = new DateTime(Lib.Conv2Integer(ArrDate[0]), Lib.Conv2Integer(ArrDate[1]), Lib.Conv2Integer(ArrDate[2]));
            }

            sError = "";
            string eName = "";
            IsTotalMatch = true;
            file = new ExcelFile();
            file.LoadXls(uploadfileid);
            ws = file.Worksheets["Sheet1"];
            ws2 = file.Worksheets.Add("Sheet_" + file.Worksheets.Count.ToString());

            ws2.Columns[0].Width = 255 * 10;
            ws2.Columns[1].Width = 255 * 30;
            ws2.Columns[2].Width = 255 * 40;
            ws2.Columns[3].Width = 255 * 10;

            wRow = 0; wCol = 0;
            int rRow = 0;
            while (1 == 1)
            {
                rRow++;
                eName = GetData(rRow, 2);
                if (eName.Trim().Length <= 0 || rRow > 200)
                    break;

                if (Lib.Conv2Integer(GetData(rRow, 7)) > 0)
                    WriteIncrementLetter(rRow, 12);
                if (!IsTotalMatch)
                    return;
            }

            string fName = "EMPINCMNTLTR-" + branch_code;
            if (fName.Length > 30)
                fName = fName.Substring(0, 30);
            File_Display_Name = Lib.ProperFileName(fName) + ".xls";
            File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
            File_Type = "xls";
            file.SaveXls(File_Name);
        }

        private string GetData(int cRow, int cCol)
        {
            string str = "";
            if (ws.Cells[cRow, cCol].Value != null)
                str = ws.Cells[cRow, cCol].Value.ToString();
            return str;
        }


        private void WriteIncrementLetter(int datRow, int datCol)
        {
            string sData = "";
            string PromoteDesig = "";
            string Gender = "";
            int sTot = 0;
            wRow++;
            wRow++;

            wRow++;
            wRow++;
            wRow++;
            wRow++;
            wRow++;
            wRow++;
            wRow++;

            sData = DTprint.Day.ToString() + " " + DTprint.ToString("MMMM") + " " + DTprint.Year.ToString();
            WriteData(wRow, 0, sData);
            wRow++;
            wRow++;
            Gender = GetData(datRow, 3);
            WriteData(wRow, 0, ((Gender.ToUpper() == "MALE") ? "Mr. " : "Ms. ") + GetProcessedName(GetData(datRow, 2)), true);
            wRow++;
            WriteData(wRow, 0, GetProcessedName(GetData(datRow, 4)), true);
            wRow++;
            WriteData(wRow, 0, GetProcessedName(GetData(datRow, 5)), true);
            wRow++;
            wRow++;
            if (Gender.ToUpper() == "MALE")
                WriteData(wRow, 0, "Dear Sir,");
            else
                WriteData(wRow, 0, "Dear Madam,");
            wRow++;
            wRow++;
            sData = DTEffective.Day.ToString() + " " + DTEffective.ToString("MMMM") + " " + DTEffective.Year.ToString();
           
            PromoteDesig = GetData(datRow, 10);
            if (PromoteDesig.Trim().Length > 0)
            {
                WriteData(wRow, 0, "Management is pleased to promote you to " + GetProcessedName(PromoteDesig) + " Cadre and revise your salary ");
                wRow++;
                WriteData(wRow, 0, "with effect from " + sData + " as shown below:- ");
            }
            else
                WriteData(wRow, 0, "Management is pleased to revise your salary with effect from " + sData + " as shown below:-");

            wRow++;
            wRow++;

            wRow++;
            WriteData(wRow, 1, "Effective Date", false, "ALL");
           // WriteData(wRow, 2, Common.getDateOnly(Dt_EffectiveDate.Value.ToString()), true, "ALL");
            WriteData(wRow, 2, DTEffective.ToString(Lib.FRONT_END_DATE_DISPLAY_FORMAT), true, "ALL");
            wRow++;
            WriteData(wRow, 1, "Grade", false, "ALL");
            if (PromoteDesig.Trim().Length > 0)
                WriteData(wRow, 2, GetProcessedName(PromoteDesig), true, "ALL");
            else
                WriteData(wRow, 2, GetProcessedName(GetData(datRow, 9)), true, "ALL");
            wRow++;
            WriteData(wRow, 1, "Basic Structure", false, "ALL");
            sData = GetData(datRow, 12);
            if (sData.Trim().Length <= 0)
                sData = GetData(datRow, 13);

            WriteData(wRow, 2, sData + "/-", true, "ALL");

            wRow++;

            WriteData(wRow, 1, "Designation", false, "ALL");
            if (PromoteDesig.Trim().Length > 0)
                sData = GetProcessedName(PromoteDesig);
            else
                sData = GetProcessedName(GetData(datRow, 9));

            if (sData.Trim() != "")
                sData += " - " + GetProcessedName(GetData(datRow, 11));

            WriteData(wRow, 2, sData, true, "ALL");
            wRow++;
            wRow++;

            wRow++;
            WriteData(wRow, 1, "Monthly Salary Structure", true);
            wRow++;
            wRow++;
            WriteData(wRow, 1, "Particulars", true, "ALL");
            WriteData(wRow, 2, "Amount (Rs.)", true, "ALL", "RIGHT");

            sTot = 0;
            while (1 == 1)
            {
                sData = GetData(datRow, datCol);
                if (Lib.Conv2Integer(sData) > 0)
                {
                    wRow++;
                    WriteData(wRow, 1, GetData(0, datCol), false, "ALL");
                    WriteData(wRow, 2, Lib.Conv2Integer(sData), true, "ALL_NFORMAT", "RIGHT");
                    sTot += Lib.Conv2Integer(sData);
                }
                datCol++;

                if (GetData(0, datCol).Trim().ToUpper() == "TOTAL" || datCol > 50)
                    break;
            }

            if (sTot > 0)
            {
                if (sTot != Lib.Conv2Integer(GetData(datRow, datCol)))
                {
                    IsTotalMatch = false;
                    //      MessageBox.Show("Total Mismatch for " + GetData(datRow, 2), "Error");
                    sError += " | Total Mismatch for " + GetData(datRow, 2);
                    return;
                }
                wRow++;
                WriteData(wRow, 1, "Total (Rs)", true, "ALL");
                WriteData(wRow, 2, sTot, true, "ALL_NFORMAT", "RIGHT");
            }
            wRow++;
            wRow++;
            WriteData(wRow, 0, "All other terms and conditions will remain same.");
            wRow++;
            wRow++;
            WriteData(wRow, 0, "Please sign and return the duplicate copy of this letter, as token of your acceptance.");
            wRow++;
            wRow++;
            WriteData(wRow, 0, "Yours truly,");
            wRow++;
            WriteData(wRow, 0, "For  " + GetProcessedName(GetData(datRow, 4)), true);
            wRow++;
            wRow++;
            wRow++;
            wRow++;
            //if (GetData(datRow, 4).ToUpper() == "CARGOMAR" || GetData(datRow, 4).ToUpper() == "SREEGAYATRI")
            //    WriteData(wRow, 0, "Partner", true);
            //else
            //    WriteData(wRow, 0, "Managing Director", true);

            if (company_code == "CMR" || company_code== "SGT")
                WriteData(wRow, 0, "Partner", true);
            else
                WriteData(wRow, 0, "Managing Director", true);

            wRow++;

            ws2.HorizontalPageBreaks.Add(wRow);
        }

        private string GetProcessedName(string sName)
        {
            string str = "";
            string[] Swrd = sName.Split(' ');
            foreach (string S in Swrd)
            {
                if (str.Trim() != "")
                    str += " ";
                if (S.Length > 1)
                {
                    if (S.Contains("."))
                        str += GetProcessedDotName(S);
                    else
                        str += S.Substring(0, 1).ToUpper() + S.Substring(1).ToLower();
                }
                else
                    str += S.ToUpper();
            }

            return str;

        }

        private string GetProcessedDotName(string sName)
        {
            string str = "";
            string[] Swrd = sName.Split('.');
            foreach (string S in Swrd)
            {
                if (str.Trim() != "")
                    str += ".";
                if (S.Length > 1)
                    str += S.Substring(0, 1).ToUpper() + S.Substring(1).ToLower();
                else
                    str += S.ToUpper();
            }
            return str;
        }
    }
}
