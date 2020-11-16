using System;
using System.Data;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.CI;

namespace BLXml
{
    public class CostInvoice : BaseReport
    {
        private DataTable DT_MASTER = new DataTable();
        private DataTable DT_HOUSE = new DataTable();
        private DataTable DT_CNTR = new DataTable();
        private DataTable DT_COSTD = new DataTable();
        private DataRow DR_MASTER = null;
        public Boolean IsError = false;
        public string ErrorMessage = "";
        private InvoiceMessage InvMessage = null;

        private string ErrorValues = "";
        private string sql = "";
        private string MessageNumber = "";
        private string InvoiceNumber = "";
        public string COST_PKID = "";
        DBConnection Con_Oracle = null;
        string comp_name = "", comp_add1 = "", comp_add2 = "", comp_add3 = "", comp_add4 = "";
        public void Generate()
        {
            try
            {
                ErrorMessage = "";
                IsError = false;
                ReadData();
                if (DT_MASTER.Rows.Count <= 0)
                {
                    ErrorMessage = "Details not Found";
                    return;
                }

                IsError = false;
                GenerateXmlFiles(COST_PKID);
                WriteXmlFiles();

                DT_MASTER.Rows.Clear();
                DT_HOUSE.Rows.Clear();
                DT_CNTR.Rows.Clear();
                DT_COSTD.Rows.Clear();
            }
            catch (Exception ex)
            {
                IsError = true;
                ErrorMessage += " |" + ex.Message.ToString();
                //  MessageBox.Show(ex.ToString());
            }
        }
        private void ReadData()
        {

            Con_Oracle = new DBConnection();

            sql = "";

            sql = " select cost_pkid,  cost_refno, cost_folderno,cost_date,cost_drcr, b.hbl_bl_no, hbl_folder_no, b.hbl_type,b.rec_category, ";
            sql += " agent.cust_code as agent_code,  ";
            sql += " agent.cust_name as agent_name,  ";
            sql += " agentadd.add_line1 as agent_line1,";
            sql += " agentadd.add_line2 as agent_line2,";
            sql += " agentadd.add_line3 as agent_line3,";
            sql += " agentadd.add_line4 as agent_line4,";
            sql += " vsl.param_name as vessel_name, hbl_vessel_no,";
            sql += " pol.param_code as pol_code,pol.param_name as pol_name,";
            sql += " pod.param_code as pod_code,pod.param_name as pod_name,";
            sql += " curr.param_code as curr_code,";
            sql += " cost_exrate,cost_buy_pp,cost_buy_cc,";
            sql += " cost_sell_pp,cost_sell_cc,cost_rebate,";
            sql += " cost_ex_works,cost_hand_charges,cost_kamai,";
            sql += " cost_other_charges,cost_asper_amount,cost_buy_tot,";
            sql += " cost_sell_tot,cost_profit ,cost_our_profit,cost_your_profit,";
            sql += " cost_drcr_amount,cost_drcr_amount_inr,cost_expense,cost_income";
            sql += " from costingm a";
            sql += " inner join hblm b on a.cost_mblid = b.hbl_pkid";
            //sql += " left join customerm agent on b.hbl_agent_id = agent.cust_pkid";
            //sql += " left join addressm agentadd on hbl_agent_br_id = agentadd.add_pkid";
            sql += " left join customerm agent on a.cost_jv_agent_id = agent.cust_pkid";
            sql += " left join addressm agentadd on a.cost_jv_agent_br_id = agentadd.add_pkid";
            sql += " left join param vsl on hbl_vessel_id = vsl.param_pkid";
            sql += " left join param pol on hbl_pol_id = pol.param_pkid";
            sql += " left join param pod on hbl_pod_id = pod.param_pkid";
            sql += " left join param curr on cost_currency_id = curr.param_pkid";
            sql += " where cost_pkid ='" + COST_PKID + "'";

            DT_MASTER = Con_Oracle.ExecuteQuery(sql);

            sql = "";
            sql += " select cost_pkid, h.hbl_bl_no, cons.cust_name as consignee_name ";
            sql += " from costingm a";
            sql += " inner join hblm m on a.cost_mblid = m.hbl_pkid";
            sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
            sql += " left join customerm cons on h.hbl_imp_id = cons.cust_pkid";
            sql += " where cost_pkid ='" + COST_PKID + "'";

            DT_HOUSE = Con_Oracle.ExecuteQuery(sql);

            sql = "";
            sql += " select cost_pkid,cntr_no,ctype.param_code as cntr_type ";
            sql += " from costingm a";
            sql += " inner join containerm c on cost_mblid = cntr_booking_id ";
            sql += " left join param ctype on cntr_type_id = ctype.param_pkid ";
            sql += " where cost_pkid ='" + COST_PKID + "'";

            DT_CNTR = Con_Oracle.ExecuteQuery(sql);


            sql = "";
            sql += " select costd_parent_id,costd_blno,costd_acc_name,costd_acc_qty,costd_acc_rate ,costd_acc_amt ,costd_ctr, costd_remarks ";
            sql += " ,costd_srate,costd_brate,costd_split ";
            sql += " from costingd ";
            sql += " where costd_parent_id ='" + COST_PKID + "'";
            sql += " and nvl(costd_category,'COSTING') = 'INVOICE' ";
            sql += " order by costd_ctr";
            DT_COSTD = Con_Oracle.ExecuteQuery(sql);

            Con_Oracle.CloseConnection();

        }

        private void GenerateXmlFiles(string COST_ID)
        {
            InvMessage = new InvoiceMessage();
            InvMessage.Items = Generate_InvoiceHeader(COST_ID);
        }

        private object[] Generate_InvoiceHeader(string COST_ID)
        {
            object[] Items = null;
            int ArrIndex = 0;
            try
            {
                Items = new object[2];
                Items[ArrIndex++] = Generate_MessageInfo();
                Items[ArrIndex++] = Generate_HeaderData(COST_ID);
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return Items;
        }
        private InvoiceMessageMessageInfo Generate_MessageInfo()
        {
            InvoiceMessageMessageInfo Rec = null;
            try
            {
                this.MessageNumber = XmlLib.GetNewMessageNumber();

                Rec = new InvoiceMessageMessageInfo();
                Rec.MessageSender = XmlLib.messageSenderField;
                Rec.MessageNumber = this.MessageNumber;
                Rec.MessageRecipient = XmlLib.messageRecipientField;
                Rec.CreatedDateTime = ConvertYMDDate(DateTime.Now.ToString());
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return Rec;
        }

        private InvoiceMessageInvoice Generate_HeaderData(string COST_ID)
        {
            InvoiceMessageInvoice Rec = null;
            string PreData = "1";
            InvoiceNumber = "";
            try
            {

                foreach (DataRow Dr in DT_MASTER.Select("cost_pkid ='" + COST_ID + "'", "cost_pkid"))
                {
                    if (PreData != Dr["cost_pkid"].ToString())
                    {
                        PreData = Dr["cost_pkid"].ToString();

                        Rec = new InvoiceMessageInvoice();
                        Rec.invoice_type = Dr["cost_drcr"].ToString() == "DR" ? "DEBIT NOTE" : "CREDIT NOTE";
                        Rec.invoice_number = Dr["cost_refno"].ToString();
                        InvoiceNumber = Rec.invoice_number.ToString();
                        Rec.invoice_date = ConvertYMDDate(Dr["cost_date"].ToString());
                        Rec.invoice_amount = Dr["cost_drcr_amount"].ToString();
                        Rec.cntrs = Generate_InvoiceCntrs(COST_ID);
                        Rec.feeder_vessel = Dr["vessel_name"].ToString() + " " + Dr["hbl_vessel_no"].ToString();
                        Rec.master = Dr["hbl_bl_no"].ToString();
                        Rec.houseLineItems = Generate_InvoiceHouses(COST_ID);
                        Rec.pol_code = Dr["POL_CODE"].ToString();
                        Rec.pol_name = Dr["POL_NAME"].ToString();
                        Rec.pod_code = Dr["POD_CODE"].ToString();
                        Rec.pod_name = Dr["POD_NAME"].ToString();
                        Rec.currency = Dr["CURR_CODE"].ToString();
                        Rec.our_refno = Dr["COST_FOLDERNO"].ToString();
                        Rec.Parties = Generate_InvoiceParties(Dr["cost_pkid"].ToString());
                        Rec.InvoiceLineItems = Generate_InvoiceLineItems(Dr["cost_pkid"].ToString());
                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return Rec;
        }

        private InvoiceMessageInvoiceCntrsCntr[] Generate_InvoiceCntrs(string COST_ID)
        {
            InvoiceMessageInvoiceCntrsCntr Rec = null;
            InvoiceMessageInvoiceCntrsCntr[] mCntrList = null;
            int ArrIndex = 0;
            try
            {
                mCntrList = new InvoiceMessageInvoiceCntrsCntr[DT_CNTR.Rows.Count];
                foreach (DataRow Dr in DT_CNTR.Rows)
                {
                    Rec = new InvoiceMessageInvoiceCntrsCntr();
                    Rec.Value = Lib.GetCntrno(Dr["cntr_no"].ToString()) + "[" + Dr["cntr_type"].ToString() + "]";
                    mCntrList[ArrIndex++] = Rec;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mCntrList;
        }

        private LineItem[] Generate_InvoiceHouses(string COST_ID)
        {
            LineItem Rec = null;
            LineItem[] mHouseList = null;
            int ArrIndex = 0;
            try
            {
                mHouseList = new LineItem[DT_HOUSE.Rows.Count];
                foreach (DataRow Dr in DT_HOUSE.Select("cost_pkid ='" + COST_ID + "'", "cost_pkid"))
                {
                    Rec = new LineItem();
                    Rec.HouseBLNo = Dr["HBL_BL_NO"].ToString();
                    Rec.consignee = Dr["CONSIGNEE_NAME"].ToString();
                    mHouseList[ArrIndex++] = Rec;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mHouseList;
        }
        private InvoiceMessageInvoicePartiesParty[] Generate_InvoiceParties(string COST_ID)
        {
            InvoiceMessageInvoicePartiesParty Rec = null;
            InvoiceMessageInvoicePartiesParty[] mPartyList = null;
            int ArrIndex = 0;
            string PreData = "1";
            try
            {

                mPartyList = new InvoiceMessageInvoicePartiesParty[1];
                foreach (DataRow Dr in DT_MASTER.Select("cost_pkid ='" + COST_ID + "'", "cost_pkid"))
                {
                    if (PreData != Dr["cost_pkid"].ToString())
                    {
                        PreData = Dr["cost_pkid"].ToString();
                        Rec = new InvoiceMessageInvoicePartiesParty();
                        Rec.Type = "AGENT";
                        Rec.Code = Dr["AGENT_CODE"].ToString();
                        Rec.Name = Dr["AGENT_NAME"].ToString();
                        Rec.AddressLine1 = Dr["AGENT_LINE1"].ToString();
                        Rec.AddressLine2 = Dr["AGENT_LINE2"].ToString();
                        Rec.AddressLine3 = Dr["AGENT_LINE3"].ToString();
                        Rec.AddressLine4 = Dr["AGENT_LINE4"].ToString();
                        mPartyList[ArrIndex++] = Rec;
                    }
                }

            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mPartyList;
        }

        private LineItem[] Generate_InvoiceLineItems(string COST_ID)
        {
            LineItem Rec = null;
            LineItem[] mInvList = null;
            int ArrIndex = 0;
            string CurrCode = "";
            try
            {
                foreach (DataRow Dr in DT_MASTER.Select("cost_pkid ='" + COST_ID + "'", "cost_pkid"))
                {
                    if (Dr["CURR_CODE"].ToString().Length > 0)
                    {
                        CurrCode = Dr["CURR_CODE"].ToString();
                        break;
                    }
                }
                mInvList = new LineItem[DT_COSTD.Rows.Count];
                foreach (DataRow Dr in DT_COSTD.Select("costd_parent_id = '" + COST_ID + "'", "costd_ctr"))
                {
                    Rec = new LineItem();
                    Rec.slno = (ArrIndex + 1).ToString();
                    Rec.refno = Dr["costd_blno"].ToString();
                    Rec.reftype = "HBL";
                    Rec.description = Dr["costd_acc_name"].ToString();
                    Rec.remarks = Dr["costd_remarks"].ToString();
                    Rec.currency = CurrCode;
                    if (Lib.Convert2Decimal(Dr["costd_srate"].ToString()) != 0)
                        Rec.SellingRate = Dr["costd_srate"].ToString();
                    else
                        Rec.SellingRate = "";
                    if (Lib.Convert2Decimal(Dr["costd_brate"].ToString()) != 0)
                        Rec.BuyingRate = Dr["costd_brate"].ToString();
                    else
                        Rec.BuyingRate = "";
                    if (Lib.Convert2Decimal(Dr["costd_split"].ToString()) != 0)
                        Rec.split = Dr["costd_split"].ToString();
                    else
                        Rec.split = "";
                    Rec.qty = Dr["costd_acc_qty"].ToString();
                    Rec.rate = Dr["costd_acc_rate"].ToString();
                    Rec.amount = Dr["costd_acc_amt"].ToString();
                    mInvList[ArrIndex++] = Rec;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mInvList;
        }
        private void WriteXmlFiles()
        {
            try
            {
                if (InvMessage == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | Costing Invoice Not Generated.";
                    return;
                }


                string FileName = "IN";
                FileName = Lib.ProperFileName(InvoiceNumber);
                if (FileName.Trim() == "")
                    FileName = "IN";
                XmlLib.CreateSentFolder();
                //FileName = Lib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";

                FileName = XmlLib.sentFolder + FileName + ".XML";
                if (File.Exists(FileName))
                    File.Delete(FileName);

                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");
                XmlSerializer mySerializer = new XmlSerializer(typeof(InvoiceMessage));
                StreamWriter writer = new StreamWriter(FileName);
                mySerializer.Serialize(writer, InvMessage, ns);
                writer.Close();

                string File_Name = FileName.Replace(".XML", ".PDF");
                if (File.Exists(File_Name))
                    File.Delete(File_Name);
                ProcessInvoice(File_Name);
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }
        private string ConvertYMDDate(string sDate)
        {
            if (sDate != null)
            {
                if (sDate.Trim().Length > 0)
                    sDate = Convert.ToDateTime(sDate).ToString("yyyy-MM-dd HH:mm:ss");
            }
            return sDate;
        }

        public void ProcessInvoice(string File_Name)
        {
            try
            {
                DR_MASTER = null;
                foreach (DataRow Dr in DT_MASTER.Select("cost_pkid ='" + COST_PKID + "'", "cost_pkid"))
                {
                    DR_MASTER = Dr;
                    break;
                }

                if (DR_MASTER == null)
                    return;

                BeginReport(Page_Height, Page_Width);
                PrintInvoice();
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
            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw ex;
            }
        }

        private void PrintInvoice()
        {
            InitPage();
            AddPage(1100, 800);
            //LoadImage(ImagePath + "\\sgtLogo.png", 20, Row + 10, 68, 70);
            WriteCompanyDetails();
            WriteDetails();
        }
        private void InitPage()
        {
            HCOL1 = 10;
            HCOL2 = HCOL1 + 100;
            HCOL3 = HCOL2 + 10;
            HCOL4 = HCOL3 + 300;
            HCOL5 = HCOL4 + 130;
            HCOL6 = HCOL5 + 50;
            HCOL7 = HCOL6 + 10;
            HCOL8 = HCOL7 + 80;
            HCOL9 = HCOL8 + 100;

            ifontName = "Calibri";
            ifontSize = 10;

            Row = 40;
            ROW_HT = 15;
        }
        private void WriteCompanyDetails()
        {
            comp_name = ""; comp_add1 = ""; comp_add2 = ""; comp_add3 = ""; comp_add4 = "";

            Dictionary<string, object> mSearchData = new Dictionary<string, object>();
            LovService mService = new LovService();
            mSearchData.Add("table", "COMP_ADDRESS");
            mSearchData.Add("comp_code", XmlLib.Company_Code);
            DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
            if (Dt_CompAddress != null)
            {
                foreach (DataRow Dr in Dt_CompAddress.Rows)
                {
                    comp_name = Dr["COMP_NAME"].ToString();
                    comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                    comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                    comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                    // comp_add4 = "Email : " + Dr["COMP_email"].ToString() + " Web : " + Dr["COMP_WEB"].ToString();
                    comp_add4 = "Email : hodoc@cargomar.in Web : " + Dr["COMP_WEB"].ToString();
                    break;
                }
            }


            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_name, ifontName, ifontSize + 6, "", "BC");
            Row += ROW_HT;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_add1, ifontName, ifontSize, "", "C");
            Row += ROW_HT;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_add2, ifontName, ifontSize, "", "C");
            Row += ROW_HT;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_add3, ifontName, ifontSize, "", "C");
            Row += ROW_HT;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_add4, ifontName, ifontSize, "", "C");

            Row += ROW_HT;
            Row += ROW_HT;
        }

        private void WriteDetails()
        {
            DrawHLine(10, Row, Page_Width - 20);
            Row += ROW_HT;
            string sTitle = "";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                sTitle = "DEBIT NOTE";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
                sTitle = "CREDIT NOTE";
            AddXYLabel(10, Row, ROW_HT, 800 - 4, sTitle, ifontName, ifontSize + 6, "", "BC");
            Row += ROW_HT;
            DrawHLine(10, Row + ROW_HT, Page_Width - 20);

            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, DR_MASTER["AGENT_NAME"].ToString(), ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "NUMBER", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL7, Row, ROW_HT, HCOL9 - HCOL7, DR_MASTER["COST_REFNO"].ToString(), ifontName, ifontSize, "", "L");

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, DR_MASTER["AGENT_LINE1"].ToString(), ifontName, ifontSize, "", "L");
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "DATE", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL7, Row, ROW_HT, HCOL9 - HCOL7, Lib.DatetoStringDisplayformat(DR_MASTER["cost_date"]), ifontName, ifontSize, "", "L");

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL9 - HCOL1, DR_MASTER["AGENT_LINE2"].ToString(), ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL9 - HCOL1, DR_MASTER["AGENT_LINE3"].ToString(), ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL9 - HCOL1, DR_MASTER["AGENT_LINE4"].ToString(), ifontName, ifontSize, "", "L");

            Row += ROW_HT;
            Row += ROW_HT;

            sTitle = "";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                sTitle = "WE DEBIT YOUR ACCOUNT FOR THE FOLLOWING";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
                sTitle = "WE CREDIT YOUR ACCOUNT FOR THE FOLLOWING";

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL9 - HCOL1, sTitle, ifontName, ifontSize, "TB", "CB");

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "FEEDER VESSEL", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL9 - HCOL3, DR_MASTER["VESSEL_NAME"].ToString() + " " + DR_MASTER["HBL_VESSEL_NO"].ToString(), ifontName, ifontSize, "", "L");

            if (DR_MASTER["HBL_BL_NO"].ToString().Trim() != "")
            {
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "MBL NO", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL9 - HCOL3, DR_MASTER["HBL_BL_NO"].ToString(), ifontName, ifontSize, "", "L");
            }

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "CONTAINER", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            string sCntr = "";
            int iCount = 0;
            foreach (DataRow Dr in DT_CNTR.Rows)
            {
                iCount++;
                if (sCntr != "")
                    sCntr += ",";
                sCntr += Lib.GetCntrno(Dr["cntr_no"].ToString()) + "[" + Dr["cntr_type"].ToString() + "]";
                if (iCount % 6 == 0)
                {
                    if (iCount > 6)
                        Row += ROW_HT;
                    AddXYLabel(HCOL3, Row, ROW_HT, HCOL9 - HCOL3, sCntr, ifontName, ifontSize, "", "L");
                    sCntr = "";
                }
            }
            if (sCntr != "")
            {
                if (iCount > 6)
                    Row += ROW_HT;
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL9 - HCOL3, sCntr, ifontName, ifontSize, "", "L");
            }

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "HBLNO/CONSIGNEE", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            iCount = 0;
            string sBLs = "";
            foreach (DataRow Dr in DT_HOUSE.Rows)
            {
                iCount++;
                sBLs = Dr["hbl_bl_no"].ToString() + " / " + Dr["consignee_name"].ToString();
                if (iCount > 1)
                    Row += ROW_HT;
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL9 - HCOL3, sBLs, ifontName, ifontSize, "", "L");
            }

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "PORT OF LOADING", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL9 - HCOL3, DR_MASTER["POL_NAME"].ToString(), ifontName, ifontSize, "", "L");

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "DESTINATION", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL9 - HCOL3, DR_MASTER["POD_NAME"].ToString(), ifontName, ifontSize, "", "L");

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "OUR REFNO", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL9 - HCOL3, DR_MASTER["COST_FOLDERNO"].ToString(), ifontName, ifontSize, "", "L");

            Row += ROW_HT;
            Row += ROW_HT;

            bool IsRemarksExist = false;
            foreach (DataRow Dr in DT_COSTD.Rows)
            {
                if (Dr["costd_remarks"].ToString().Trim().Length > 0)
                {
                    IsRemarksExist = true;
                    break;
                }
            }

            AddXYLabel(HCOL1, Row, ROW_HT, (HCOL1 + 100) - HCOL1, "REFNO", ifontName, ifontSize, "", "LB");//LTBR
            AddXYLabel(HCOL1 + 100, Row, ROW_HT, HCOL4 - (HCOL1 + 100), "PARTICULARS", ifontName, ifontSize, "", "LB");//LTBR
            if (IsRemarksExist)
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, "REMARKS", ifontName, ifontSize, "", "LB");//LTBR
            AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, "AMOUNT(" + DR_MASTER["curr_code"].ToString() + ")", ifontName, ifontSize, "", "RB", 0, 0, 0, 0, 0, 0, -5);//LTBR
            DrawHLine(10, Row + ROW_HT, Page_Width - 20);
            string[] AccNameArry;
            string[] AccRemarksArry;

            int NameColWidth = 0;
            int RemarksColWidth = 0;
            int iLen = 0;

            if (IsRemarksExist)
            {
                NameColWidth = Lib.Conv2Integer((HCOL4 - (HCOL1 + 100)).ToString());
                RemarksColWidth = Lib.Conv2Integer((HCOL8 - HCOL4).ToString());
                iLen = 0;
                foreach (DataRow Dr in DT_COSTD.Rows)
                {
                    AccNameArry = Lib.ConvertString2Lines(Dr["costd_acc_name"].ToString(), NameColWidth, "WORD");
                    AccRemarksArry = Lib.ConvertString2Lines(Dr["costd_remarks"].ToString(), RemarksColWidth, "WORD");

                    if (AccNameArry != null)
                        iLen = AccNameArry.Length;

                    if (AccRemarksArry != null)
                        if (AccRemarksArry.Length > iLen)
                            iLen = AccRemarksArry.Length;

                    Row += ROW_HT;
                    AddXYLabel(HCOL1, Row, ROW_HT, (HCOL1 + 100) - HCOL1, Dr["costd_blno"].ToString(), ifontName, ifontSize, "", "L");//LR
                    if (AccNameArry != null)
                        AddXYLabel(HCOL1 + 100, Row, ROW_HT, HCOL4 - (HCOL1 + 100), AccNameArry.Length > 0 ? AccNameArry[0] : "", ifontName, ifontSize, "", "L");//LR
                    if (AccRemarksArry != null)
                        AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, AccRemarksArry.Length > 0 ? AccRemarksArry[0] : "", ifontName, ifontSize, "", "L");//LR
                    AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, Lib.NumericFormat(Dr["costd_acc_amt"].ToString(), 2), ifontName, ifontSize, "", "R", 0, 0, 0, 0, 0, 0, -5);//LR
                    for (int i = 1; i < iLen; i++)
                    {
                        Row += ROW_HT;
                        AddXYLabel(HCOL1, Row, ROW_HT, (HCOL1 + 100) - HCOL1, "", ifontName, ifontSize, "", "L");//LR
                        if (AccNameArry != null)
                            AddXYLabel(HCOL1 + 100, Row, ROW_HT, HCOL4 - (HCOL1 + 100), AccNameArry.Length > i ? AccNameArry[i].Trim() : "", ifontName, ifontSize, "", "L");//LR
                        if (AccRemarksArry != null)
                            AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, AccRemarksArry.Length > i ? AccRemarksArry[i].Trim() : "", ifontName, ifontSize, "", "L");//LR
                        AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, "", ifontName, ifontSize, "", "R", 0, 0, 0, 0, 0, 0, -5);//LR
                    }
                }
            }
            else
            {
                NameColWidth = Lib.Conv2Integer((HCOL8 - (HCOL1 + 100)).ToString());
                iLen = 0;
                foreach (DataRow Dr in DT_COSTD.Rows)
                {
                    AccNameArry = Lib.ConvertString2Lines(Dr["costd_acc_name"].ToString(), NameColWidth, "WORD");
                    if (AccNameArry != null)
                        iLen = AccNameArry.Length;

                    Row += ROW_HT;
                    AddXYLabel(HCOL1, Row, ROW_HT, (HCOL1 + 100) - HCOL1, Dr["costd_blno"].ToString(), ifontName, ifontSize, "", "L");//LR
                    if (AccNameArry != null)
                        AddXYLabel(HCOL1 + 100, Row, ROW_HT, HCOL8 - (HCOL1 + 100), AccNameArry.Length > 0 ? AccNameArry[0] : "", ifontName, ifontSize, "", "L");//LR
                    AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, Lib.NumericFormat(Dr["costd_acc_amt"].ToString(), 2), ifontName, ifontSize, "", "R", 0, 0, 0, 0, 0, 0, -5);//LR
                    for (int i = 1; i < iLen; i++)
                    {
                        Row += ROW_HT;
                        AddXYLabel(HCOL1, Row, ROW_HT, (HCOL1 + 100) - HCOL1, "", ifontName, ifontSize, "", "L");//LR
                        if (AccNameArry != null)
                            AddXYLabel(HCOL1 + 100, Row, ROW_HT, HCOL8 - (HCOL1 + 100), AccNameArry.Length > i ? AccNameArry[i].Trim() : "", ifontName, ifontSize, "", "L");//LR
                        AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, "", ifontName, ifontSize, "", "R", 0, 0, 0, 0, 0, 0, -5);//LR
                    }
                }
            }

            Row += ROW_HT;
            DrawHLine(10, Row, Page_Width - 20);
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, "TOTAL", ifontName, ifontSize, "", "LB");//LTB
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, "", ifontName, ifontSize, "", "LB");//TB
            AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, Lib.NumericFormat(DR_MASTER["cost_drcr_amount"].ToString(), 2), ifontName, ifontSize, "", "RB", 0, 0, 0, 0, 0, 0, -5);//LTBR
            Row += ROW_HT;
            Row += ROW_HT;
            //DrawHLine(10, Row, Page_Width - 20);

            decimal nDrCRAmt = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString());

            if (nDrCRAmt < 0)
                nDrCRAmt = Math.Abs(nDrCRAmt);

            string sAmt = Lib.NumericFormat(nDrCRAmt.ToString(), 2);

            string sWords = "";
            if (DR_MASTER["curr_code"].ToString() != "INR")
                sWords = Number2Word_USD.Convert(sAmt, DR_MASTER["CURR_CODE"].ToString(), "CENTS");
            if (DR_MASTER["curr_code"].ToString() == "INR")
                sWords = Number2Word_RS.Convert(sAmt, "INR", "PAISE");

            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL9 - HCOL1, sWords, ifontName, ifontSize, "TB", "LB");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL9 - HCOL1, "E.&.O.E", ifontName, ifontSize, "TB", "LB");

        }
    }
}
