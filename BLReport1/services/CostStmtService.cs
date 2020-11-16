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
    public class CostStmtService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<CostStmtReport> mList = new List<CostStmtReport>();
        CostStmtReport mrow;
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
        
        string ErrorMessage = "";
        string agent_id = "";
        string agent_name = "";
        string curr_id = "";
        string curr_code = "";
        string from_date = "";
        string to_date = "";
        decimal tot_debit = 0;
        decimal tot_credit = 0;
        decimal tot_diff = 0;


        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<CostStmtReport>();
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
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();
            
               

                if (SearchData.ContainsKey("agent_id"))
                {
                    agent_id = SearchData["agent_id"].ToString();
                    agent_name = SearchData["agent_name"].ToString();
                }
                    

                if (SearchData.ContainsKey("curr_id"))
                {
                    curr_id = SearchData["curr_id"].ToString();
                    curr_code = SearchData["curr_code"].ToString();
                }
                   

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);



                sql = " select a.* from (  ";
                sql += " select  jv_acc_id,  '1' as roworder,  null as roworder2, null as jv_pkid, to_date('{FDATE}') as jvh_date, '{CURRCODE}' as curcode, ";
                sql += "  null as jvh_vrno,   null as jvh_type,   null as JVH_REMARKS, null as jvh_reference,  null as rec_category,  ";
                sql += "  null as jv_debit, null as jv_credit,   null as jv_exrate,  null as inr,  ";
                sql += "  sum(case when nvl(jv_Debit,0) >0 then jv_ftotal else 0 end) - sum(case when nvl(jv_credit,0) >0 then jv_ftotal else 0 end) as OPENING  ";
                sql += "  from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id where  JVH_YEAR  = {YEARCODE} and a.REC_COMPANY_CODE ='{COMPCODE}' and a.REC_BRANCH_CODE  ='{BRCODE}' and ";
                sql += "  JV_ACC_ID   ='{AGENTID}' and JVH_TYPE  != 'OS' and JVH_TYPE  != 'OB' ";
                sql += "  and JVH_TYPE  != 'OC' and (JVH_TYPE = 'OP' or JVH_DATE  <'{FDATE}') and ";
                sql += "  jv_curr_id ='{CURRID}' ";
                sql += "  group by jv_acc_id ";
                sql += "  union all  ";
                sql += "  select  jv_acc_id,  case when nvl(a.jvh_location,'OTHERS') = 'OTHERS' then   '2' else '3' end as roworder,  nvl(a.jvh_location,'OTHERS') as roworder2, ";
                sql += "  jv_pkid,  jvh_date, '{CURRCODE}' as curcode, jvh_vrno,  jvh_type, jvh_remarks,  jvh_reference,  a.rec_category,  ";
                sql += "  case when nvl(jv_Debit,0) >0 then jv_ftotal else 0 end as jv_debit, ";
                sql += "  case when nvl(jv_credit,0) >0 then jv_ftotal else 0 end as jv_Credit, jv_exrate,  ";
                sql += "  case when nvl(jv_Debit,0) >0 then jv_debit else jv_credit end as inr , 0 as opening  ";
                sql += "  from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id  where  JVH_YEAR  = {YEARCODE} and ";
                sql += "  a.REC_COMPANY_CODE ='{COMPCODE}' and a.REC_BRANCH_CODE  ='{BRCODE}' and JV_ACC_ID   ='{AGENTID}' ";
                sql += "  and JVH_TYPE  != 'OB' and JVH_TYPE  != 'OS'  and JVH_TYPE  != 'OC' and JVH_TYPE  != 'OP'  ";
                sql += "  and JVH_DATE  >='{FDATE}' and JVH_DATE   <= '{EDATE}' and ";
                sql += "  jv_curr_id ='{CURRID}'  ";
                sql += " ) a   order by  roworder, roworder2,  jvh_date,jvh_reference";
               


                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{AGENTID}",agent_id);
                sql = sql.Replace("{CURRID}", curr_id);
                sql = sql.Replace("{CURRCODE}",curr_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                decimal opening = 0;
                tot_credit = 0;
                tot_debit = 0;
                tot_diff = 0;
                foreach (DataRow Dr in Dt_List.Rows)
                {
                 

                    mrow = new CostStmtReport();
                    mrow.rowtype = "DETAIL";
                    mrow.rowcolor = "BLACK";
                    opening = Lib.Conv2Decimal(Dr["OPENING"].ToString());
                    if(opening != 0)
                    {
                        if(opening < 0)
                        {
                            mrow.jv_credit = Math.Abs(opening);
                            tot_credit += Math.Abs(opening);
                        }
                        else
                        {
                            mrow.jv_debit = Math.Abs(opening);
                            tot_debit += Math.Abs(opening);
                        }
                    }
                    else
                    {
                        mrow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                        mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                    }

                    mrow.jvh_reference = Dr["jvh_reference"].ToString();
                    mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mrow.jvh_type = Dr["jvh_type"].ToString();
                    mrow.jvh_vrno = Dr["jvh_vrno"].ToString();
                    mrow.jvh_remarks = Dr["JVH_REMARKS"].ToString();
                    mrow.curr_code = Dr["curcode"].ToString();
                    mrow.jv_exrate = Lib.Conv2Decimal(Dr["jv_exrate"].ToString());
                   
                    mrow.reccategory = Dr["rec_category"].ToString();
                    mrow.roworder2 = Dr["roworder2"].ToString();
                    mrow.branch = Dr["roworder2"].ToString();
                    mList.Add(mrow);
                    tot_debit += Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    tot_credit += Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                }
                if (mList.Count > 1)
                {
                    if(tot_debit > tot_credit)
                    {
                        tot_diff = tot_debit - tot_credit;
                        mrow = new CostStmtReport();
                        mrow.rowtype = "TOTALDUE";
                        mrow.rowcolor = "RED";
                        mrow.jvh_reference = "TOTAL DUE FROM "+agent_name;
                        mrow.jv_credit = Lib.Conv2Decimal(Lib.NumericFormat(tot_diff.ToString(), 2));
                        mList.Add(mrow);
                        tot_credit += tot_diff;
                    }
                    else
                    {
                        tot_diff = tot_credit - tot_debit;
                        mrow = new CostStmtReport();
                        mrow.rowtype = "TOTALDUE";
                        mrow.rowcolor = "RED";
                        mrow.jvh_reference = "TOTAL DUE TO " + agent_name;
                        mrow.jv_debit = Lib.Conv2Decimal(Lib.NumericFormat(tot_diff.ToString(), 2));
                        mList.Add(mrow);
                        tot_debit += tot_diff;
                    }
                    mrow = new CostStmtReport();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "RED";
                    mrow.jvh_reference = "TOTAL";
                    mrow.jv_debit = Lib.Conv2Decimal(Lib.NumericFormat(tot_debit.ToString(), 2));
                    mrow.jv_credit = Lib.Conv2Decimal(Lib.NumericFormat(tot_credit.ToString(), 2));

                    mList.Add(mrow);
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintCostStatementReport();
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

        private void PrintCostStatementReport()
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
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "CostStatmentReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 25;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 7;
                WS.Columns[4].Width = 256 * 8;
                WS.Columns[5].Width = 256 * 40;
                WS.Columns[6].Width = 256 * 6;
                WS.Columns[7].Width = 256 * 8;
                WS.Columns[8].Width = 256 * 12;
                WS.Columns[9].Width = 256 * 12;
                WS.Columns[10].Width = 256 * 12;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
              

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
                Lib.WriteData(WS, iRow, 1, "STATEMENT OF "+agent_name, _Color, true, "", "L", "", 13, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, "PERIOD FROM "+from_date +" TO "+to_date , _Color, true, "", "L", "", 13, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 11;
                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, "REFERENCE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PARTICULARS ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-RATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                _Size = 10;
                string row_order2 = "";
                string row_order = "";
                foreach (CostStmtReport Rec in mList )
                {
                    row_order2 = nvl(Rec.roworder2, "").ToString();

                    if (Rec.rowtype.ToString()  == "TOTAL")
                    {
                        iRow++; iCol = 1;_Border = "TB";_Bold = true;
                        Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jvh_reference, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jv_debit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jv_credit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);

                    }
                    if (Rec.rowtype.ToString() == "TOTALDUE")
                    {
                        iRow++; iRow++; iCol = 1;_Bold = true;
                        Lib.WriteData(WS, iRow, 1, nvl(Rec.jvh_reference, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                       
                        Lib.WriteData(WS, iRow, 8, nvl(Rec.jv_debit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, 9, nvl(Rec.jv_credit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        

                    }
                    if (Rec.rowtype.ToString() == "DETAIL")
                    {
                        iRow++; iCol = 1;

                        if (nvl(Rec.roworder2, "").ToString() == "")
                        {
                            if (nvl(Rec.reccategory, "").ToString() == "")
                            {
                                iCol = 1;
                                Lib.WriteData(WS, iRow, 1, "OPENING BALANCE AS PER STATEMENT ON " + Rec.jvh_date, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                                Lib.WriteData(WS, iRow, 6, nvl(Rec.curr_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                                Lib.WriteData(WS, iRow, 8, nvl(Rec.jv_debit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, 9, nvl(Rec.jv_credit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                            }
                            else
                            {

                                iRow++; iCol = 1;
                                Lib.WriteData(WS, iRow, 2, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, _Bold, _Border, "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                                Lib.WriteData(WS, iRow, 3, nvl(Rec.jvh_type, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                                Lib.WriteData(WS, iRow, 4, nvl(Rec.jvh_vrno, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                                Lib.WriteData(WS, iRow, 5, nvl(Rec.jvh_remarks, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                                Lib.WriteData(WS, iRow, 6, nvl(Rec.curr_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                                Lib.WriteData(WS, iRow, 7, nvl(Rec.jv_exrate, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, 8, nvl(Rec.jv_debit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, 9, nvl(Rec.jv_credit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                                Lib.WriteData(WS, iRow, 10, nvl(Rec.reccategory, ""), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                            }

                        }
                        else
                        {
                            if (row_order2 != row_order)
                            {
                                iRow++; iCol = 1;
                                Lib.WriteData(WS, iRow, 1, "A/C. " + nvl(Rec.roworder2, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                                iRow++; iRow++; iCol = 1;
                            }


                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jvh_reference, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, _Bold, _Border, "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jvh_type, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jvh_vrno, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jvh_remarks, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.curr_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jv_exrate, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jv_debit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.jv_credit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                            Lib.WriteData(WS, iRow, iCol++, nvl(Rec.reccategory, ""), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        }
                    }
                    row_order = row_order2;
                    row_order2 = "";
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
        public object nvl(object svalue, object sret)
        {
            if (svalue == null)
                return sret;
            else
                return svalue;
        }
    }
}

