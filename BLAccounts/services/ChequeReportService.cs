using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLAccounts
{
    public class ChequeReportService : BaseReport
    {
        public string report_folder = "";
        public string folderid = "";
        public string File_Name = "";
        public string File_Type = "";
        public string File_Display_Name = "myreport.pdf";
        public string company_code = "";
        public string branch_code = "";
        public string user_code = "";
        public string PKID = "";
        public string BL_FORMAT_ID = "";
        public bool IsAcPayee = false;
        string sql = "";
        DataTable Dt_COLPOS = new DataTable();
        DataTable Dt_Data = new DataTable();

        private bool IsNotOverCheque = false;
        private float Xtolrnce = 0;
        private DataRow DR = null;
        private string BANK_NAME = "";
        private string CHQNO = "";
        private string JVNO = "";
        private string CHQDATE = "";
        private string CHQNAME1 = "";
        private string CHQNAME2 = "";
        private string CHQWORDS1 = "";
        private string CHQWORDS2 = "";
        private string CHQAMOUNT = "";
        private int CHQNAME1_LEN = 70;
        private int CHQWORDS1_LEN = 65;
        private int x1 = 0, y1 = 0, h1 = 0, w1 = 0, fsize = 0;
        private string fname = "", sStyle = "", sBorder = "";

        DBConnection Con_Oracle = null;

        public void Process()
        {
            try
            {
                ReadData();
                if (DR == null)
                    throw new Exception("No Details to Print...Print CHQUE");

                if (DR["JV_DUE_DATE"].Equals(DBNull.Value))
                    throw new Exception("Cheque Date Not Enterd");

                if (DR["JV_PAID_TO"].Equals(DBNull.Value))
                    throw new Exception("Paid To Not Enterd");

                SetData();


                string fname = "myreport";
                fname = DR["JV_PAID_TO"].ToString().Replace(" ", "");
                if (fname.Length > 30)
                    fname = fname.Substring(0, 30);
                File_Display_Name = Lib.ProperFileName(fname) + ".pdf";
                File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
                File_Type = "pdf";

                BeginReport(Page_Height, Page_Width);
                PrintCheque();
                EndReport();
                if (ExportList != null)
                {
                    Export2Pdf mypdf = new Export2Pdf();
                    mypdf.ExportList = ExportList;
                    mypdf.FileName = File_Name;
                    mypdf.Page_Height = Page_Height;
                    mypdf.Page_Width = Page_Width;
                    mypdf.Process();
                }

                UpdateChqprint(DR["jv_parent_id"].ToString());
            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw ex;
            }
        }

        private void UpdateChqprint(string JVHID)
        {
            try
            {
                int ChqPrintCount = 0;
                Con_Oracle = new DBConnection();

                sql  = "select nvl(max(jvh_chq_printed),0) + 1 as chqprint ";
                sql += " from ledgerh a";
                sql += " where a.jvh_pkid = '{JVHID}' ";
                sql = sql.Replace("{JVHID}", JVHID);

                DataTable Dt_Temp = new DataTable();
                Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                if(Dt_Temp.Rows.Count>0)
                {
                    ChqPrintCount = Lib.Conv2Integer(Dt_Temp.Rows[0]["chqprint"].ToString());
                }

                sql = "update ledgerh set jvh_chq_printed = " + ChqPrintCount.ToString() + " where jvh_pkid ='" + JVHID + "'";

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                string remarks = "TO: " + (CHQNAME1 + " " + CHQNAME2).Trim();
                remarks += ", BNK:" + BANK_NAME;
                remarks += ", CHQ#:" + CHQNO;
                remarks += ", DT:" + CHQDATE;
                remarks += ", AMT:" + CHQAMOUNT;
                Lib.AuditLog("CHQ-PRINT", "BP", "PRINT", company_code, branch_code, user_code, JVHID, "CHQ-" + CHQNO, remarks);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
        }

        private void ReadData()
        {
            Con_Oracle = new DBConnection();
            DR = null;
            sql = "select jv_parent_id,jv_due_date,jv_paid_to,jv_total,jv_chqno,jv_bank,jv_branch from ledgert where jv_pkid = '" + this.PKID + "'";
            Dt_Data = new DataTable();
            Dt_Data = Con_Oracle.ExecuteQuery(sql);
            if (Dt_Data.Rows.Count > 0)
                DR = Dt_Data.Rows[0];

            IsNotOverCheque = false;
            if (DR != null)
            {
                CHQNO = DR["jv_chqno"].ToString();
                BANK_NAME = DR["jv_bank"].ToString()+" " + DR["jv_branch"].ToString();

                sql = "select jvh_not_over_chq from ledgerh where jvh_pkid = '" + DR["jv_parent_id"].ToString() + "'";
                Object sVal = Con_Oracle.ExecuteScalar(sql);
                if (sVal.ToString() == "Y")
                    IsNotOverCheque = true;

                sql = "";
                sql = " select a.BLF_NAME, nvl(BLF_LEFT_MARGIN,0) + BLF_COL_X as BLF_COL_X ,";
                sql += " nvl(BLF_TOP_MARGIN,0) + BLF_COL_Y as BLF_COL_Y,";
                sql += " BLF_COL_HEIGHT,BLF_COL_WIDTH,BLF_COL_FONT_SIZE,BLF_COL_NAME,BLF_COL_FONT,BLF_COL_BORDER,BLF_COL_STYLE";
                sql += " from printformatm a inner join printformatd b on a.blf_pkid = b.blf_col_header_id";
                sql += " where blf_type = 'CHQUE'  ";
                sql += " and blf_pkid  ='" + BL_FORMAT_ID + "'";
                sql += " and blf_col_enabled = 'Y' ";
                sql += " order by BLF_COL_X, BLF_COL_Y";
                Dt_COLPOS = new DataTable();
                Dt_COLPOS = Con_Oracle.ExecuteQuery(sql);
                

                DataRow[] drs = Dt_COLPOS.Select("BLF_COL_NAME ='CHQNAME1'");
                if (drs.Length > 0)
                    CHQNAME1_LEN = Lib.Conv2Integer(drs[0]["BLF_COL_WIDTH"].ToString());
                drs = Dt_COLPOS.Select("BLF_COL_NAME ='CHQWORDS1'");
                if (drs.Length > 0)
                    CHQWORDS1_LEN = Lib.Conv2Integer(drs[0]["BLF_COL_WIDTH"].ToString());
            }

            Con_Oracle.CloseConnection();
        }

        private void SetData()
        {
           
            CHQDATE = "";
            if (!DR["JV_DUE_DATE"].Equals(DBNull.Value))
                CHQDATE = ((DateTime)DR["JV_DUE_DATE"]).ToString("dd/MM/yyyy");

            string Name1 = DR["JV_PAID_TO"].ToString();
            SplitWords(CHQNAME1_LEN, Name1, ref CHQNAME1, ref CHQNAME2);

            CHQAMOUNT = string.Format("{0:#0.00}", double.Parse(DR["JV_TOTAL"].ToString()));

            decimal nTotal = Lib.Convert2Decimal(CHQAMOUNT);

            string Words = Number2Word_RS.Convert(nTotal.ToString(), "", "Paise");
            Words = Words.ToUpper().Trim();
            
            GetPosition("CHQWORDS1");
            if (fname != "")
                ifontName = fname;
            if (fsize != 0)
                ifontSize = fsize;

            string[] ArryWrds = Lib.ConvertString2Lines(Words, CHQWORDS1_LEN - 50, "WORD", ifontName, ifontSize, sStyle);
            CHQWORDS1 = ArryWrds.Length > 0 ? ArryWrds[0] : "";
            CHQWORDS2 = ArryWrds.Length > 1 ? ArryWrds[1] : "";
        }
        private void SplitWords(int Len1, string Words, ref string W1, ref string W2)
        {
            W1 = ""; W2 = "";
            if (Words.Length <= Len1)
            {
                W1 = Words;
                return;
            }
            if (Words.Substring(Len1, 1) == " ")
            {
                W1 = Words.Substring(0, Len1);
                W2 = Words.Substring(Len1).Trim();
                return;
            }
            int k = Len1;
            Boolean bOk = false;
            for (; k >= 0; k--)
            {
                if (Words.Substring(k, 1) == " ")
                {
                    bOk = true;
                    W1 = Words.Substring(0, k);
                    W2 = Words.Substring(k).Trim();
                    break;
                }
            }
            if (bOk == false)
            {
                W1 = Words;
                W2 = "";
            }

        }
        private Boolean GetPosition(string FldName)
        {
            Boolean bRet = false;
            try
            {
                 
                foreach (DataRow dr in Dt_COLPOS.Select("blf_col_name ='" + FldName.Trim() + "'"))
                {
                    x1 = Lib.Conv2Integer(dr["BLF_COL_X"].ToString());
                    y1 = Lib.Conv2Integer(dr["BLF_COL_Y"].ToString());
                    h1 = Lib.Conv2Integer(dr["BLF_COL_HEIGHT"].ToString());
                    w1 = Lib.Conv2Integer(dr["BLF_COL_WIDTH"].ToString());
                    fsize = Lib.Conv2Integer(dr["BLF_COL_FONT_SIZE"].ToString());

                    fname = dr["BLF_COL_FONT"].ToString();
                    sStyle = dr["BLF_COL_STYLE"].ToString();
                    sBorder = dr["BLF_COL_BORDER"].ToString();
                    bRet = true;
                }
            }
            catch (Exception)
            {
                bRet = false;
            }
            return bRet;
        }
        private void PrintCheque()
        {
            Row = 10;
            AddPage(Page_Height, Page_Width);
            FillData();
        }
        private void FillData()
        {
            string str = "";

            if (GetPosition("CHQPAYEE"))
                AddXYLabel(x1, y1, h1, w1, (IsAcPayee == true) ? "A/C PAYEE" : "", fname, fsize, (IsAcPayee == true) ? sBorder : "", sStyle, 0, 0, 4, 16, 0, 0, 0, 16);

            if (GetPosition("CHQDATE"))
            {
                if (w1 < 50)
                {
                    CHQDATE = "";
                    if (!DR["JV_DUE_DATE"].Equals(DBNull.Value))
                        CHQDATE = ((DateTime)DR["JV_DUE_DATE"]).ToString("ddMMyyyy");
                    for (int i = 0; i < CHQDATE.Length; i++)
                    {
                        AddXYLabel(x1, y1, h1, w1, CHQDATE[i].ToString(), fname, fsize, sBorder, sStyle, 0, 0, 4, 16, 0, 0, 0, 16);
                        x1 += w1;
                    }
                }
                else
                {
                    AddXYLabel(x1, y1, h1, w1, Lib.DatetoStringDisplayformat(DR["JV_DUE_DATE"]), fname, fsize, sBorder, sStyle, 0, 0, 4, 16, 0, 0, 0, 16);
                }
            }

            if (GetPosition("CHQNAME1"))
                AddXYLabel(x1, y1, h1, w1, CHQNAME1, fname, fsize, sBorder, sStyle, 0, 0, 2, 16, 0, 0, Xtolrnce, fsize + 2);

            if (GetPosition("CHQNAME2"))
                AddXYLabel(x1, y1, h1, w1, CHQNAME2, fname, fsize, sBorder, sStyle, 0, 0, 2, 16, 0, 0, Xtolrnce, fsize + 2);

            if (IsNotOverCheque)
            {
                if (GetPosition("CHQNOTOVER"))
                {
                    str = "NOT OVER - ";
                    if (Lib.Convert2Decimal(CHQAMOUNT) > 0)
                    {
                        //str = Common.NumericFormat(CHQAMOUNT, 2);
                        str += string.Format("{0:#,0.00}", double.Parse(CHQAMOUNT));
                    }
                    AddXYLabel(x1, y1, h1, w1, str, fname, fsize, sBorder, sStyle, 0, 0, 2, 16, 0, 0, Xtolrnce, fsize + 2);
                }
            }
            else
            {
                if (GetPosition("CHQWORDS1"))
                    AddXYLabel(x1, y1, h1, w1, CHQWORDS1, fname, fsize, sBorder, sStyle, 0, 0, 2, 16, 0, 0, Xtolrnce, fsize + 2);

                if (GetPosition("CHQWORDS2"))
                    AddXYLabel(x1, y1, h1, w1, CHQWORDS2, fname, fsize, sBorder, sStyle, 0, 0, 2, 16, 0, 0, Xtolrnce, fsize + 2);

                if (GetPosition("CHQAMOUNT"))
                {
                    str = "";
                    if (Lib.Convert2Decimal(CHQAMOUNT) > 0)
                    {
                        //str = Common.NumericFormat(CHQAMOUNT, 2);
                        str = string.Format("{0:#,0.00}", double.Parse(CHQAMOUNT));
                    }
                    AddXYLabel(x1, y1, h1, w1, str, fname, fsize, sBorder, sStyle, 0, 0, 4, 16, 0, 0, Xtolrnce, fsize + 2);
                }

            }

        }
    }
}


