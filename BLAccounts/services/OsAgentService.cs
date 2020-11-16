using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

using XL.XSheet;

namespace BLAccounts
{
    public class OsAgentService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;

        int iRow = 0;
        int iCol = 0;

        string type = "";
        string report_folder = "";
        string File_Name = "";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string to_date = "";
        string agent_id = "";
        string curr_id = "";
        string curr_code = "";
        string category = "";
        string category_type = "";

        decimal tot_G1 = 0;
        decimal tot_G2 = 0;
        decimal tot_G3 = 0;
        decimal tot_G4 = 0;
        decimal tot_G5 = 0;
        decimal tot_balance = 0;
        decimal tot_advance = 0;


        decimal tot_G1_inr = 0;
        decimal tot_G2_inr = 0;
        decimal tot_G3_inr = 0;
        decimal tot_G4_inr = 0;
        decimal tot_G5_inr = 0;


        decimal tot_balance_inr = 0;
        decimal tot_advance_inr = 0;

        Boolean all = false;
        Boolean IsOverDue = false;
        Hashtable HT = new Hashtable();
        List<OsAgentReport> mList = null;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();



            Con_Oracle = new DBConnection();
            mList = new List<OsAgentReport>();

            OsAgentReport mrow = new OsAgentReport();

            type = SearchData["type"].ToString();

            report_folder = SearchData["report_folder"].ToString();
            PKID = SearchData["pkid"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();
            string edate = SearchData["to_date"].ToString();
            to_date = Lib.StringToDate(edate).ToUpper();
            category_type = SearchData["category_type"].ToString();
            if (SearchData.ContainsKey("agent_id"))
                agent_id = SearchData["agent_id"].ToString();

            if (SearchData.ContainsKey("curr_id"))
                curr_id = SearchData["curr_id"].ToString();

            if (SearchData.ContainsKey("curr_code"))
                curr_code = SearchData["curr_code"].ToString();

            if (SearchData.ContainsKey("category"))
                category = SearchData["category"].ToString();


            IsOverDue = (Boolean)SearchData["isoverdue"];
            all = (Boolean)SearchData["all"];

            string searchstring = SearchData["searchstring"].ToString().ToUpper();

            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            File_Name = System.IO.Path.Combine(report_folder, PKID);

            try
            {

                sql = " select 'A' as type,agereport.rec_category as catg,acc_code,acc_name,curr.param_code as curr_code,    ";
                sql += " sum(case when days  <=30               then DR - CR else 0 end) as G1,";
                sql += " sum(case when days  >30 and days <=60  then DR - CR else 0 end) as G2,";
                sql += " sum(case when days  >60 and days <=90  then DR - CR else 0 end) as G3,";
                sql += " sum(case when days  >90 and days <=180 then DR - CR else 0 end) as G4,";
                sql += " sum(case when days  >180               then DR - CR else 0 end) as G5,";
                sql += " sum( nvl(DR,0) - nvl(CR,0) ) as balance , 0 as Advance,  ";


                sql += " sum(case when days  <=30               then DR_INR - CR_INR else 0 end) as G1_INR,";
                sql += " sum(case when days  >30 and days <=60  then DR_INR - CR_INR else 0 end) as G2_INR,";
                sql += " sum(case when days  >60 and days <=90  then DR_INR - CR_INR else 0 end) as G3_INR,";
                sql += " sum(case when days  >90 and days <=180 then DR_INR - CR_INR else 0 end) as G4_INR,";
                sql += " sum(case when days  >180               then DR_INR - CR_INR else 0 end) as G5_INR,";

                sql += " sum( nvl(DR_INR,0) - nvl(CR_INR,0) ) as balance_inr,  0 as advance_inr from ( ";
                
                sql += " 	 select    jv_acc_id,   jvh_date,   jv_curr_id,   round(sysdate -jvh_date,0) as days,     ";
                sql += " 	 acc_code,acc_name, substr(h.rec_category,1,3) as rec_category, jvh_reference,       ";
                sql += " 	 jv_ftotal as Amount,   nvl(xref_Amt,0) as Allocation,     ";
                sql += " 	 jv_ftotal - nvl(xref_Amt,0) as Balance,     ";
                sql += " 	 case when jv_debit <>0 then jv_ftotal - nvl(xref_Amt,0) else 0 end as DR,    ";
                sql += " 	 case when jv_credit <>0 then jv_ftotal - nvl(xref_Amt,0) else 0 end as CR,   ";
                sql += "     case when jv_debit<>0 then jv_debit -nvl(xref_Amt_inr, 0) else 0 end as DR_INR,  ";
                sql += "     case when jv_credit<>0 then jv_credit -nvl(xref_Amt_inr, 0) else 0 end as CR_INR ";
                sql += " 	 from ledgerh h inner join ledgert a on jvh_pkid = jv_parent_id    ";
                sql += " 	 inner join acctm on jv_acc_id = acc_pkid   ";
                sql += " 	 inner join acgroupm on (acc_group_id = acgrp_pkid and acgrp_pkid = '84919031-18A5-0BFD-1899-5CDB06FCBF4A') ";
                sql += " 	 left join    (     ";
                sql += " 	 	  select 	     std_jv_pkid,     std_currencyid,     sum(std_amt) as xref_Amt,sum(std_amt_inr) as xref_Amt_inr ";
                sql += " 		  from	 stmtm inner join stmtd on stm_pkid = std_parentid   where stm_date  <= '{EDATE}' ";
                sql += " 		  group by std_jv_pkid,std_currencyid     ";
                sql += "      )  b on (a.jv_pkid = b.std_jv_pkid and a.jv_curr_id = b.std_currencyid)     ";
                sql += " 	 where      h.jvh_type not in ('OP','OB', 'OI', 'OS')     and h.rec_deleted = 'N'     and nvl(h.jvh_reference,'A') <> 'COSTING ADJUSTMENT' ";
                sql += " 	 and h.rec_company_code= '{COMPCODE}' and h.rec_branch_code ='{BRCODE}' and jvh_type in('HO','OC','IN-ES') and   (jv_ftotal - nvl(xref_Amt,0)) <>0  ";
                sql += " 	 and h.jvh_date <= '{EDATE}'";

                if (curr_id.Trim().Length > 0)
                {

                    sql += " and  a.jv_curr_id = '" + curr_id + "' ";
                }


                if (agent_id.Trim().Length > 0)
                {
                    sql += " and  jv_acc_id = '" + agent_id + "'";
                }

                if (category != "ALL")
                {
                    sql += " and  substr(h.rec_category,1,3) = '{CATEGORY}'";
                }

                sql += " )  AgeReport   ";
                sql += " left join param curr on ( jv_curr_id = curr.param_pkid)  ";
                sql += " group by agereport.rec_category,curr.param_code,acc_code,acc_name   ";

                sql += " union all   ";

                sql += " select 'B' as type, AgeReport.rec_category as catg,acc_code,acc_name,curr.param_code as curr_code,0 as G1, 0 as G2, 0 as G3, 0 as G4,0 as G5, ";
                sql += " sum( nvl(DR,0) - nvl(CR,0) ) as balance,   sum( nvl(DR,0) - nvl(CR,0) ) as Advance,  ";
                sql += " 0 as G1_INR, 0 as G2_INR, 0 as G3_INR, 0 as G4_INR,0 as G5_INR,";
                sql += " sum(nvl(DR_inr, 0) - nvl(CR_inr, 0)) as balance_inr, sum(nvl(DR_inr, 0) - nvl(CR_inr, 0)) as Advance_inr  from ( ";

                sql += " select     jv_acc_id,    jv_curr_id,    ";
                sql += " acc_code,acc_name, substr(h.rec_category,1,3) as rec_category,    ";
                sql += " jv_ftotal - nvl(xref_Amt,0) as Balance,      ";
                sql += " case when jv_debit <>0 then jv_ftotal - nvl(xref_Amt,0) else 0 end as DR,     ";
                sql += " case when jv_credit <>0 then jv_ftotal - nvl(xref_Amt,0) else 0 end as CR,    ";
                sql += " case when jv_debit<>0 then jv_debit -nvl(xref_Amt_inr, 0) else 0 end as DR_inr, ";
                sql += " case when jv_credit<>0 then jv_credit -nvl(xref_Amt_inr, 0) else 0 end as CR_inr ";

                sql += " from ledgerh h inner join ledgert a on jvh_pkid = jv_parent_id     ";
                sql += " inner join acctm on jv_acc_id = acc_pkid     ";
                sql += " inner join acgroupm on (acc_group_id = acgrp_pkid and acgrp_pkid = '84919031-18A5-0BFD-1899-5CDB06FCBF4A')    ";
                sql += " left join     (      ";
                sql += " select std_jv_pkid,     std_currencyid,     sum(std_amt) as xref_Amt,sum(std_amt_inr) as xref_Amt_inr ";
                sql += "  from	 stmtm inner join stmtd on stm_pkid = std_parentid   where stm_date  <= '{EDATE}'   ";
                sql += "  group by std_jv_pkid,std_currencyid  ";
                sql += "   )  b on (a.jv_pkid = b.std_jv_pkid and a.jv_curr_id = b.std_currencyid) ";
                sql += "  where  h.jvh_type not in ('OP','OB','OI','OS','HO','OC','IN-ES' )  ";
                sql += "  and nvl(h.jvh_reference,'A') <> 'COSTING ADJUSTMENT' ";
                sql += "  and h.rec_company_code = '{COMPCODE}' and h.rec_branch_code= '{BRCODE}' and   (jv_ftotal - nvl(xref_Amt,0)) <> 0  ";
                sql += "  and h.jvh_date <= '{EDATE}'  ";

                if (curr_id.Trim().Length > 0)
                {

                    sql += " and  a.jv_curr_id = '" + curr_id + "' ";
                }

                if (agent_id.Trim().Length > 0)
                {
                    sql += " and  jv_acc_id = '" + agent_id + "'";
                }
                if (category != "ALL")
                {
                    sql += " and  substr(h.rec_category,1,3) = '{CATEGORY}'";
                }

                sql += " )  AgeReport ";
                sql += " left join param curr on ( jv_curr_id = curr.param_pkid ) ";
                sql += " group by AgeReport.rec_category,curr.param_code,acc_code,acc_name   ";
                if (category_type == "DETAIL")
                {
                    sql += " order by acc_name,type,catg,curr_code";
                }

                if (category_type == "SUMMERY")
                {
                    sql += " order by acc_name,curr_code,catg,type ";
                }


                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{EDATE}", to_date);
                sql = sql.Replace("{CATEGORY}", category);


                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                if (category_type == "DETAIL")
                {
                    tot_balance = 0;
                    tot_advance = 0;
                    tot_G1 = 0;
                    tot_G2 = 0;
                    tot_G3 = 0;
                    tot_G4 = 0;
                    tot_G5 = 0;

                    tot_balance_inr = 0;
                    tot_advance_inr = 0;
                    tot_G1_inr = 0;
                    tot_G2_inr = 0;
                    tot_G3_inr = 0;
                    tot_G4_inr = 0;
                    tot_G5_inr = 0;

                    foreach (DataRow Dr in Dt_List.Rows)
                    {

                        mrow = new OsAgentReport();
                        mrow.rowtype = "ROW";
                        mrow.rowcolor = "BLACK";
                        mrow.reccategory = Dr["catg"].ToString();
                        mrow.acc_code = Dr["acc_code"].ToString();
                        mrow.acc_name = Dr["acc_name"].ToString();
                        mrow.curr_code = Dr["curr_code"].ToString();

                        mrow.G1 = Lib.Conv2Decimal(Dr["G1"].ToString());
                        mrow.G2 = Lib.Conv2Decimal(Dr["G2"].ToString());
                        mrow.G3 = Lib.Conv2Decimal(Dr["G3"].ToString());
                        mrow.G4 = Lib.Conv2Decimal(Dr["G4"].ToString());
                        mrow.G5 = Lib.Conv2Decimal(Dr["G5"].ToString());

                        mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                        mrow.advance = Lib.Conv2Decimal(Dr["Advance"].ToString());


                        mrow.G1_INR = Lib.Conv2Decimal(Dr["G1_INR"].ToString());
                        mrow.G2_INR = Lib.Conv2Decimal(Dr["G2_INR"].ToString());
                        mrow.G3_INR = Lib.Conv2Decimal(Dr["G3_INR"].ToString());
                        mrow.G4_INR = Lib.Conv2Decimal(Dr["G4_INR"].ToString());
                        mrow.G5_INR = Lib.Conv2Decimal(Dr["G5_INR"].ToString());


                        mrow.balance_inr = Lib.Conv2Decimal(Dr["balance_inr"].ToString());
                        mrow.advance_inr = Lib.Conv2Decimal(Dr["Advance_inr"].ToString());

                        mList.Add(mrow);

                        tot_G1 += Lib.Conv2Decimal(Dr["G1"].ToString());
                        tot_G2 += Lib.Conv2Decimal(Dr["G2"].ToString());
                        tot_G3 += Lib.Conv2Decimal(Dr["G3"].ToString());
                        tot_G4 += Lib.Conv2Decimal(Dr["G4"].ToString());
                        tot_G5 += Lib.Conv2Decimal(Dr["G5"].ToString());
                        tot_balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                        tot_advance += Lib.Conv2Decimal(Dr["Advance"].ToString());


                        tot_G1_inr += Lib.Conv2Decimal(Dr["G1_INR"].ToString());
                        tot_G2_inr += Lib.Conv2Decimal(Dr["G2_INR"].ToString());
                        tot_G3_inr += Lib.Conv2Decimal(Dr["G3_INR"].ToString());
                        tot_G4_inr += Lib.Conv2Decimal(Dr["G4_INR"].ToString());
                        tot_G5_inr += Lib.Conv2Decimal(Dr["G5_INR"].ToString());
                        tot_balance_inr += Lib.Conv2Decimal(Dr["balance_inr"].ToString());
                        tot_advance_inr += Lib.Conv2Decimal(Dr["Advance_inr"].ToString());

                    }

                    if (mList.Count > 1 && curr_id.Trim().Length > 0)
                    {
                        mrow = new OsAgentReport();
                        mrow.rowtype = "TOTAL";
                        mrow.rowcolor = "RED";
                        mrow.acc_code = "TOTAL";
                        mrow.G1 = Lib.Conv2Decimal(Lib.NumericFormat(tot_G1.ToString(), 2));
                        mrow.G2 = Lib.Conv2Decimal(Lib.NumericFormat(tot_G2.ToString(), 2));
                        mrow.G3 = Lib.Conv2Decimal(Lib.NumericFormat(tot_G3.ToString(), 2));
                        mrow.G4 = Lib.Conv2Decimal(Lib.NumericFormat(tot_G4.ToString(), 2));
                        mrow.G5 = Lib.Conv2Decimal(Lib.NumericFormat(tot_G5.ToString(), 2));
                        mrow.balance = Lib.Conv2Decimal(Lib.NumericFormat(tot_balance.ToString(), 2));
                        mrow.advance = Lib.Conv2Decimal(Lib.NumericFormat(tot_advance.ToString(), 2));


                        mrow.G1_INR = Lib.Conv2Decimal(Lib.NumericFormat(tot_G1_inr.ToString(), 2));
                        mrow.G2_INR = Lib.Conv2Decimal(Lib.NumericFormat(tot_G2_inr.ToString(), 2));
                        mrow.G3_INR = Lib.Conv2Decimal(Lib.NumericFormat(tot_G3_inr.ToString(), 2));
                        mrow.G4_INR = Lib.Conv2Decimal(Lib.NumericFormat(tot_G4_inr.ToString(), 2));
                        mrow.G5_INR = Lib.Conv2Decimal(Lib.NumericFormat(tot_G5_inr.ToString(), 2));
                        mrow.balance_inr = Lib.Conv2Decimal(Lib.NumericFormat(tot_balance_inr.ToString(), 2));
                        mrow.advance_inr = Lib.Conv2Decimal(Lib.NumericFormat(tot_advance_inr.ToString(), 2));

                        mList.Add(mrow);
                    }

                }
                if (category_type == "SUMMERY")
                {
                    HT = new Hashtable();
                    decimal nSea = 0, nAir = 0, nOth = 0, nAdj = 0, nBal = 0;
                    string sCode = "";
                    string sData = "";
                    string CURCODE = "";

                    foreach (DataRow Drow in Dt_List.Rows)
                    {
                        sData = Drow["acc_name"].ToString() + Drow["curr_code"].ToString();
                        if (sCode != sData)
                        {

                            if (sCode != "")
                            {

                                nBal = 0;
                                nBal += (nSea > 0) ? nSea : 0;
                                nBal -= (nSea < 0) ? Math.Abs(nSea) : 0;
                                nBal += (nAir > 0) ? nAir : 0;
                                nBal -= (nAir < 0) ? Math.Abs(nAir) : 0;
                                nBal += (nOth > 0) ? nOth : 0;
                                nBal -= (nOth < 0) ? Math.Abs(nOth) : 0;
                                nBal += (nAdj > 0) ? nAdj : 0;
                                nBal -= (nAdj < 0) ? Math.Abs(nAdj) : 0;

                                mrow.sea = Lib.Conv2Decimal(nSea.ToString());
                                mrow.air = Lib.Conv2Decimal(nAir.ToString());
                                mrow.oth = Lib.Conv2Decimal(nOth.ToString());
                                mrow.adj = Lib.Conv2Decimal(nAdj.ToString());
                                mrow.bal = Lib.Conv2Decimal(nBal.ToString());

                                mList.Add(mrow);
                                UpdateTotal(CURCODE, nBal);
                            }


                            sCode = Drow["acc_name"].ToString() + Drow["curr_code"].ToString();
                            CURCODE = Drow["curr_code"].ToString();
                            nSea = 0; nAir = 0; nOth = 0; nAdj = 0; nBal = 0;
                            mrow = new OsAgentReport();
                            mrow.rowtype = "SUMMERY";
                            mrow.curr_code = Drow["curr_code"].ToString();
                            mrow.acc_name = Drow["acc_name"].ToString();
                            mrow.acc_code = Drow["acc_code"].ToString();

                        }
                        if (Drow["catg"].ToString() == "SEA")
                            nSea += Lib.Convert2Decimal(Drow["balance"].ToString());
                        if (Drow["catg"].ToString() == "AIR")
                            nAir += Lib.Convert2Decimal(Drow["balance"].ToString());
                        if (Drow["catg"].ToString() == "OTH")
                            nOth += Lib.Convert2Decimal(Drow["balance"].ToString());
                        if (Drow["catg"].ToString() == "")
                            nAdj += Lib.Convert2Decimal(Drow["balance"].ToString());


                    }


                    nBal = 0;
                    nBal += (nSea > 0) ? nSea : 0;
                    nBal -= (nSea < 0) ? Math.Abs(nSea) : 0;
                    nBal += (nAir > 0) ? nAir : 0;
                    nBal -= (nAir < 0) ? Math.Abs(nAir) : 0;
                    nBal += (nOth > 0) ? nOth : 0;
                    nBal -= (nOth < 0) ? Math.Abs(nOth) : 0;
                    nBal += (nAdj > 0) ? nAdj : 0;
                    nBal -= (nAdj < 0) ? Math.Abs(nAdj) : 0;
                    if (nBal != 0)
                    {

                        mrow.sea = Lib.Conv2Decimal(nSea.ToString());
                        mrow.air = Lib.Conv2Decimal(nAir.ToString());
                        mrow.oth = Lib.Conv2Decimal(nOth.ToString());
                        mrow.adj = Lib.Conv2Decimal(nAdj.ToString());
                        mrow.bal = Lib.Conv2Decimal(nBal.ToString());

                        mList.Add(mrow);
                        UpdateTotal(CURCODE, nBal);
                    }

                    foreach (DictionaryEntry entry in HT)
                    {
                        //Console.WriteLine("{0}, {1}", entry.Key, entry.Value);
                        mrow = new OsAgentReport();
                        mrow.acc_code = "TOTAL";
                        mrow.rowtype = "TOTAL";
                        mrow.rowcolor = "RED";
                        mrow.curr_code = entry.Key.ToString();
                        mrow.bal = Lib.Conv2Decimal(entry.Value.ToString());
                        mList.Add(mrow);
                    }

                }


                if (type == "EXCEL")
                {
                    if (Lib.CreateFolder(report_folder))
                        ProcessExcelFile();
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
            RetData.Add("reportfile", File_Name);
            if (type != "EXCEL")
                RetData.Add("list", mList);

            return RetData;
        }
        private void UpdateTotal(string CURCODE, decimal nBal)
        {
            if (HT.ContainsKey(CURCODE))
            {
                HT[CURCODE] = Lib.Conv2Decimal(HT[CURCODE].ToString()) + nBal;
            }
            else
            {
                HT.Add(CURCODE, nBal);
            }
        }

        private void ProcessExcelFile()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 0;

            string sTitle = "";

            string sName = "Report";
            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];
            WS.ViewOptions.ShowGridLines = false;
            WS.PrintOptions.Portrait = false;
            WS.PrintOptions.FitWorksheetWidthToPages = 1;

            if (category_type == "DETAIL")
            {
                WS.Columns[0].Width = 256;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 20;
                WS.Columns[4].Width = 256 * 12;
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



                WS.Columns[5].Style.NumberFormat = "#,0.00";
                WS.Columns[6].Style.NumberFormat = "#,0.00";
                WS.Columns[7].Style.NumberFormat = "#,0.00";
                WS.Columns[8].Style.NumberFormat = "#,0.00";
                WS.Columns[9].Style.NumberFormat = "#,0.00";
                WS.Columns[10].Style.NumberFormat = "#,0.00";
                WS.Columns[11].Style.NumberFormat = "#,0.00";

                iRow = 1; iCol = 1;

                iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

                sTitle = "AGENT OUTSTANDING REPORT AS ON " + Lib.getFrontEndDate(to_date);

                Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

                iCol = 1;
                _Color = Color.DarkBlue;
                _Border = "TB";
                _Size = 10;

                Lib.WriteData(WS, iRow, iCol++, "A/C CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "A/C NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CURR-CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "0-30", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "31-60", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "61-90", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "91-180", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ADVANCE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "0-30-INR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "31-60-INR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "61-90-INR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "91-180-INR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BALANCE-INR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ADVANCE-INR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);


                foreach (OsAgentReport Dr in mList)
                {
                    iRow++; iCol = 1;
                    _Border = "";
                    _Bold = false;
                    _Color = Color.Black;

                    if (Dr.rowtype.ToString() == "TOTAL")
                    {
                        _Border = "TB";

                        _Bold = true;
                        _Color = Color.DarkBlue;
                    }

                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_name, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.reccategory, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.curr_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G1, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G2, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G3, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G4, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G5, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.balance, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.advance, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G1_INR, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G2_INR, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G3_INR, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G4_INR, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.G5_INR, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.balance_inr, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.advance_inr, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                }
                WB.SaveXls(File_Name + ".xls");
            }
            if (category_type == "SUMMERY")
            {

                WS.Columns[0].Width = 256;
                WS.Columns[1].Width = 256 * 20;
                WS.Columns[2].Width = 256 * 30;
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

                WS.Columns[4].Style.NumberFormat = "#,0.00";
                WS.Columns[5].Style.NumberFormat = "#,0.00";
                WS.Columns[6].Style.NumberFormat = "#,0.00";
                WS.Columns[7].Style.NumberFormat = "#,0.00";
                WS.Columns[8].Style.NumberFormat = "#,0.00";



                iRow = 1; iCol = 1;

                iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

                sTitle = "Agewise/International Debtors Report As On " + Lib.getFrontEndDate(to_date);

                Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

                iCol = 1;
                _Color = Color.DarkBlue;
                _Border = "TB";
                _Size = 10;

                Lib.WriteData(WS, iRow, iCol++, "A/C CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "A/C NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "CURR-CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SEA", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, "AIR", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, "OTH", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, "ADJ", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, "BAL", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                foreach (OsAgentReport Dr in mList)
                {
                    iRow++; iCol = 1;
                    _Border = "";
                    _Bold = false;
                    _Color = Color.Black;

                    if (Dr.rowtype.ToString() == "TOTAL")
                    {
                        _Border = "TB";

                        _Bold = true;
                        _Color = Color.DarkBlue;
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.curr_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.bal, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }

                    if (Dr.rowtype.ToString() == "SUMMERY")
                    {
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_name, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.curr_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.sea, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.air, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.oth, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.adj, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.bal, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }


                }
                WB.SaveXls(File_Name + ".xls");
            }

        }

        public object nvl(object svalue, object sret)
        {
            if (svalue == null)
                return sret;
            else
                return svalue;
        }

    }

}

