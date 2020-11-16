using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using System.IO;

namespace BLOperations.models
{
    public class AirBLReport : BaseReport
    {

        public Bl mRow = null;
        public string Report_Caption = "";
        public string PKID = "";
        public string InvokeType = "";
        public Boolean Chk_BL_Original = false;
        public Boolean Chk_Side1 = true;
        public Boolean Chk_Side2 = false;
        public string FooterNote = "";
        public Dictionary<int, string> fList = null;
        public string BL_TYPE = "MAWB";
        public string RootPath = "";

        private float DescStartRow = 0;
        private float OthChrgStartRow = 0;
        private int ifontSizesm = 8;
        private float ROW_HTsm = 0;

        private decimal Agent_Tot_PP = 0, Agent_Tot_CC = 0;
        private decimal Carrier_Tot_PP = 0, Carrier_Tot_CC = 0;
        private float dCol1 = 0, dCol2 = 0, dCol3 = 0, dCol4 = 0, dCol5 = 0, dCol6 = 0, dCol7 = 0;
        private float dCol8 = 0, dCol9 = 0, dCol10 = 0, dCol11 = 0, dCol12 = 0, dCol13 = 0, dCol14 = 0;
        
        private int R1 = 0;
        private int COL01 = 0;
        private int COL02 = 0;
        private int COL03 = 0;
        private int COL04 = 0;
        private int COL05 = 0;
        private int COL06 = 0;
        private int COL07 = 0;
        private int COL08 = 0;
        private int COL09 = 0;
        private int COL10 = 0;
        private int COL11 = 0;
        private int COL12 = 0;
        private int COL13 = 0;
        private int COL14 = 0;
        private int COL15 = 0;
        private int COL16 = 0;
        private int COL17 = 0;
        private int COL18 = 0;
        //private int COL19 = 0;
        //private int COL20 = 0;

        private int WD1 = 15;
        private int WD2 = 15;
        private int WD3 = 15;

        private int COLTOT = 38;

        string[] CH1 = { "", "", "", "", "", "", "", "" };
        string[] CH2 = { "", "", "", "", "", "", "", "" };

        private string BsideStyle = "";
        private char sColSplit = '~';
        private char sStyleSplit = '#';
        private string PrintFormatName = "CARGOMAR";
        private int ImageHeight = 48;
        private int ImageWidth = 48;
        private string sError = "";

        private string OthCharges = "";
        private string OthWeight = "";
        private string OthRate = "";
        private string OthTotal = "";
        private string OthPrintS = "";
        private string OthPrintC = "";

        public AirBLReport()
        {

        }
        public void ProcessData()
        {
            try
            {
                Init();

                if (mRow == null)
                    throw new Exception("No Details to Print...");

                sError = AllValid();
                if (sError.ToString().Trim()!="" )
                    throw new Exception(sError);
                PrintFormatName = mRow.bl_print_format_name;
                PrintData();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }
        private void Init()
        {
            RootPath = RootPath + "\\Images";
        }

        private string AllValid()
        {
            string str = "";

            //if (BL_TYPE == "HAWB")
            //{
            //    sql = "Select hbls_mbl_id from airhbl_summary where length(hbls_mbl_id)>0 and hbls_hbl_id = '" + this.PKID.ToString() + "'";
            //    DataTable dt_temp = new DataTable();
            //    dt_temp = orCon.RunSql(sql);
            //    if (dt_temp.Rows.Count <= 0)
            //    {
            //        bRet = false;
            //        MessageBox.Show("Master Details not Found", "Print");
            //        return bRet;
            //    }
            //}

            //for (int c = 1; c < fList.Count; c++)
            //{
            //    FooterNote = fList[c];
            //    if (FooterNote.Trim().Length <= 0)
            //        continue;
            //    if (FooterNote.ToUpper().Contains("FOR SHIPPER") && !(Chk_BL_Original == true && DR["BL_IS_ORIGINAL"].ToString().Trim() == "Y"))
            //    {
            //        MessageBox.Show("To print Original 3 - (For Shipper) Please save as Original Shipper and continue...", "Print AWB");
            //        return;
            //    }
            //    else if (FooterNote.ToUpper().Contains("CONSIGNEE") && DR["BL_IS_ORIGINAL_CNEE"].ToString().Trim() != "Y")
            //    {
            //        MessageBox.Show("To print Original 2 - (For Consignee) Please save as Original Consignee and continue...", "Print");
            //        return;
            //    }
            //    else if (FooterNote.ToUpper().Contains("FOR ISSUING CARRIER") && DR["BL_IS_ORIGINAL_CARR"].ToString().Trim() != "Y")
            //    {
            //        MessageBox.Show("To print Original 1 - (For Issuing Carrier) Please save as Original Carrier and continue...", "Print");
            //        return;
            //    }
            //}

            return str;
        }
        private void PrintData()
        {
            Row = 10;
            R1 = 0;
            addList("XLCOLUMN",
                "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "14", "8", "8", "8",
                "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "14",
                "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8",
                "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8",
                "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8",
                "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8",
                "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8", "8"
                );

            BeginReport(1100, 800);

            FooterNote = "";
            for (int c = 1; c < fList.Count; c++)
            {
                FooterNote = fList[c];
                if (FooterNote.Trim().Length <= 0)
                    continue;

                if (FooterNote.ToUpper().Contains("SHIPPER"))
                    InvokeType = "SHPR";
                else if (FooterNote.ToUpper().Contains("CONSIGNEE"))
                    InvokeType = "CNEE";
                else
                    InvokeType = "OTHR";
                WriteData();
            }
            EndReport();
        }
        private void WriteData()
        {
            Row = 10;
            if (Chk_Side1)
            {
                AddPage(1100, 800);
                WriteAddress();
                WriteHouse();
            }

            if (Chk_Side2)
            {
                CanWrite = true;
                AddPage(1100, 800);
                WriteBackSide();
                CanWrite = true;
            }

        }

        private void WriteAddress()
        {
            HCOL1 = 20; HCOL2 = 190; HCOL3 = 280; HCOL4 = 380; HCOL5 = 500; HCOL6 = 600; HCOL7 = 700; HCOL8 = 775; HCOL9 = 950; HCOL10 = 1050;
            ROW_HT = 15;
            ROW_HTsm = ROW_HT - 3;
            ifontSizesm = 6;
            ifontSize = 9;
        }

        private void WriteHouse()
        {
            // First Quarter
            COL01 = 2; COL02 = 5; COL03 = 10; COL04 = 25; COL05 = 50; COL06 = 96;
            WD1 = COL04 - COL01 - 1;
            WD2 = COL05 - COL04 - 1;
            WD3 = COL06 - COL05 - 1;

            COLTOT = COL06;
            Dictionary<int, string> AddrList = null;
            int iAddr = 0;
            string Str = "";
            string[] sData = null;
            string s1 = "";
            string s2 = "";
            Row += ROW_HT; R1++;
            if (mRow.bl_mbl_no.ToString().Contains("-"))
            {
                sData = mRow.bl_mbl_no.ToString().Split('-');
                s1 = sData[0].ToString().Trim();
                s2 = sData[1].ToString();
            }
            else
            {
                if (mRow.bl_mbl_no.ToString().Trim().Length > 3)
                {
                    s1 = mRow.bl_mbl_no.Substring(0, 3).Trim();
                    s2 = mRow.bl_mbl_no.ToString().Substring(3);
                }
            }

            AddXYLabel(HCOL1, Row, ROW_HT, 70 - HCOL1, s1, ifontName, ifontSize, "", "BC", R1, COL01, 0, 16, 0);
            Str = mRow.bl_pol_code.ToString();
            if(Str.StartsWith("IN") && Str.Length>=5) //INPOL4
            {
                Str = Str.Substring(2, 3);
            }
            AddXYLabel(70, Row, ROW_HT, 120 - 70, Str, ifontName, ifontSize, "L", "BC", R1, COL02, 0, 16, 0);// MblRec.mbld_pol_code
            AddXYLabel(120, Row, ROW_HT, 100, s2, ifontName, ifontSize, "L", "B", R1, COL03, 0, 16, 0);
            AddXYLabel(HCOL5, Row - 2, ROW_HT + 2, HCOL8 - HCOL5 - 20, (BL_TYPE == "MAWB") ? mRow.bl_mbl_no : "HAWB No. " + mRow.hbl_bl_no.ToString(), ifontName, ifontSize + 4, "", "RB", R1, COL05, WD3, 16, 0);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL2 - HCOL1, "Shipper's Name and Address", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL2, Row, ROW_HTsm, HCOL4 - HCOL2, "Shipper's Account Number", ifontName, ifontSizesm, "TL", "C", R1, COL04, WD2, 16, 0);
            AddXYLabel(HCOL4, Row, ROW_HTsm, HCOL8 - HCOL4, "Not Negotiable", ifontName, ifontSizesm, "TLR", "", R1, COL05, WD3, 16, 0);

            iAddr = 0;
            AddrList = new Dictionary<int, string>();
            if (mRow.bl_shipper_add1.ToString().Trim().Length > 0)
            {
                AddrList[iAddr] = mRow.bl_shipper_add1.ToString();
                iAddr++;
            }
            if (mRow.bl_shipper_add2.ToString().Trim().Length > 0)
            {
                AddrList[iAddr] = mRow.bl_shipper_add2.ToString();
                iAddr++;
            }
            if (mRow.bl_shipper_add3.ToString().Trim().Length > 0)
            {
                AddrList[iAddr] = mRow.bl_shipper_add3.ToString();
                iAddr++;
            }
            if (mRow.bl_shipper_add4.ToString().Trim().Length > 0)
            {
                AddrList[iAddr] = mRow.bl_shipper_add4.ToString();
                iAddr++;
            }

            Row += ROW_HTsm; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, "", ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, "", ifontName, ifontSize, "LB", "", R1, COL04, WD2, 16, 0);
            AddXYLabel(HCOL4, Row, ROW_HT + 5, HCOL8 - HCOL4, BL_TYPE == "MAWB" ? "Air Waybill" : "House Air Waybill", ifontName, ifontSize, "LR", "B", R1, COL05, WD3, 16, 0);
            Row += ROW_HT; R1++;


            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_shipper_name.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL4 + 50 - HCOL4, "Issued by", ifontName, ifontSizesm, "L", "", R1, COL05, 0, 16, 0);
            AddXYLabel(HCOL4 + 50, Row + 5, ROW_HT + 10, HCOL8 - (HCOL4 + 50), mRow.bl_issued_by1.ToString(), ifontName, ifontSize + 3, "R", "B", R1, COL05 + 6, WD3 - 6, 16, 0);

            if (BL_TYPE == "HAWB")
            {
                if (PrintFormatName.StartsWith("BANSARD"))
                {
                    ImageHeight = 200;
                    ImageWidth = 65;
                    if (PrintFormatName.Trim() != "")
                        LoadImage(RootPath + "\\" + PrintFormatName + ".jpg", HCOL4 + 160, Row - 18, ImageHeight, ImageWidth);
                }
                else
                {
                    if (PrintFormatName.Trim() != "")
                        LoadImage(RootPath + "\\" + PrintFormatName + ".jpg", HCOL4 + 2, Row + 15, ImageHeight, ImageWidth);

                    if (mRow.bl_iata_code.ToString().Trim() != "" && PrintFormatName.StartsWith("CARGOMAR"))
                    {
                        LoadImage(RootPath + "\\IATA.jpg", HCOL8 - 65, Row - 25, 60, 40);
                    }
                }
            }


            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, (AddrList.Count > 0) ? AddrList[0] : "", ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL4 + 50 - HCOL4, "", ifontName, ifontSizesm, "L", "", R1, COL05, 0, 16, 0);
            AddXYLabel(HCOL4 + 50, Row + 10, ROW_HT, HCOL8 - (HCOL4 + 50), mRow.bl_issued_by2.ToString(), ifontName, ifontSizesm, "R", "", R1, COL05 + 6, WD3 - 6, 16, 0);
            //new row
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, (AddrList.Count > 1) ? AddrList[1] : "", ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL4 + 50 - HCOL4, "", ifontName, ifontSizesm, "L", "", R1, COL05, 0, 16, 0);
            AddXYLabel(HCOL4 + 50, Row + 4, ROW_HT + 4, HCOL8 - (HCOL4 + 50), mRow.bl_issued_by3.ToString(), ifontName, ifontSizesm, "R", "", R1, COL05 + 6, WD3 - 6, 16, 0);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, (AddrList.Count > 2) ? AddrList[2] : "", ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL4 + 50 - HCOL4, "", ifontName, ifontSize, "L", "", R1, COL05, 5, 16, 0);

            if (PrintFormatName.StartsWith("BANSARD"))
                AddXYLabel(HCOL4 + 150, Row + 2, ROW_HT + 2, HCOL8 - (HCOL4 + 150), mRow.bl_issued_by4.ToString(), ifontName, ifontSizesm, "R", "B", R1, COL05 + 6, WD3 - 6, 16, 0);
            else
                AddXYLabel(HCOL4 + 50, Row + 2, ROW_HT + 2, HCOL8 - (HCOL4 + 50), mRow.bl_issued_by4.ToString(), ifontName, ifontSizesm, "R", "", R1, COL05 + 6, WD3 - 6, 16, 0);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, (AddrList.Count > 3) ? AddrList[3] : "", ifontName, ifontSize, "LR", "", R1, COL01, 0, 16, 0);

            if (PrintFormatName.StartsWith("BANSARD"))
                AddXYLabel(HCOL4 + 150, Row, ROW_HT, HCOL8 - (HCOL4 + 150), mRow.bl_issued_by5.ToString(), ifontName, ifontSizesm, "R", "B", R1, COL05 + 6, WD3 - 6, 16, 0);
            else
                AddXYLabel(HCOL4 + 50, Row, ROW_HT, HCOL8 - (HCOL4 + 50), mRow.bl_issued_by5.ToString(), ifontName, ifontSizesm, "R", "", R1, COL05 + 6, WD3 - 6, 16, 0);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL2 - HCOL1, "Consignee's Name and Address", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL2, Row, ROW_HTsm, HCOL4 - HCOL2, "Consignee's Account Number", ifontName, ifontSizesm, "TLR", "C", R1, COL04, WD2, 16, 0);
            AddXYLabel(HCOL4, Row, ROW_HTsm, HCOL8 - HCOL4, "Copies 1,2 and 3 of this Air Waybill are Originals and have the same validity", ifontName, ifontSizesm, "LBRT", "", R1, COL05, WD3, 16, 0);

            // DrawVLine(HCOL8, Row, ROW_HT * 5 + ROW_HTsm);

            iAddr = 0;
            AddrList = new Dictionary<int, string>();
            if (mRow.bl_consignee_add1.ToString().Trim().Length > 0)
            {
                AddrList[iAddr] = mRow.bl_consignee_add1.ToString();
                iAddr++;
            }
            if (mRow.bl_consignee_add2.ToString().Trim().Length > 0)
            {
                AddrList[iAddr] = mRow.bl_consignee_add2.ToString();
                iAddr++;
            }
            if (mRow.bl_consignee_add3.ToString().Trim().Length > 0)
            {
                AddrList[iAddr] = mRow.bl_consignee_add3.ToString();
                iAddr++;
            }
            if (mRow.bl_consignee_add4.ToString().Trim().Length > 0)
            {
                AddrList[iAddr] = mRow.bl_consignee_add4.ToString();
                iAddr++;
            }

            Row += ROW_HTsm; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, "", ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, "", ifontName, ifontSize, "LBR", "", R1, COL04, WD2, 16, 0);//Consgne Acc No



            DrawVLine(HCOL8, Row, ROW_HT * 5 + ROW_HT);

            s1 = "";
            Str = "It is agreed that the goods described herein are accepted in apparent good order and condition (except as noted) ";
            s1 += Str;
            AddXYLabel(HCOL4 + 2, Row + 2, ROW_HTsm, HCOL8 - (HCOL4 + 8), Str, ifontName, 5, "", "J");
            Str = "for carriage SUBJECT TO THE CONDITIONS OF CONTRACT ON THE REVERSE HEREOF.ALL GOODS MAY BE ";
            s1 += Str;
            AddXYLabel(HCOL4 + 2, Row + 2 + (ROW_HTsm * 1), ROW_HTsm, HCOL8 - (HCOL4 + 8), Str, ifontName, 5, "", "J");
            Str = "CARRIED  BY ANY OTHER MEANS INCLUDING ROAD OR ANY OTHER CARRIER UNLESS SPECIFIC CONTRARY ";
            s1 += Str;
            AddXYLabel(HCOL4 + 2, Row + 2 + (ROW_HTsm * 2), ROW_HTsm, HCOL8 - (HCOL4 + 8), Str, ifontName, 5, "", "J");
            Str = "INSTRUCTIONS ARE GIVEN HEREON BY THE SHIPPER, AND SHIPPER AGREES THAT THE SHIPMENT MAY BE ";
            s1 += Str;
            AddXYLabel(HCOL4 + 2, Row + 2 + (ROW_HTsm * 3), ROW_HTsm, HCOL8 - (HCOL4 + 8), Str, ifontName, 5, "", "J");
            Str = "CARRIED VIA INTERMEDIATE STOPPING PLACES WHICH THE CARRIER DEEMS APPROPRIATE.THE SHIPPER'S ";
            s1 += Str;
            AddXYLabel(HCOL4 + 2, Row + 2 + (ROW_HTsm * 4), ROW_HTsm, HCOL8 - (HCOL4 + 8), Str, ifontName, 5, "", "J");
            Str = "ATTENTION IS DRAWN TO THE NOTICE CONCERNING CARRIER'S LIMITATION OF LIABILITY.Shipper may ";
            s1 += Str;
            AddXYLabel(HCOL4 + 2, Row + 2 + (ROW_HTsm * 5), ROW_HTsm, HCOL8 - (HCOL4 + 8), Str, ifontName, 5, "", "J");
            Str = "increase such limitation of liability by declaring a higher value for carriage and paying a supplemental charge if required.";
            s1 += Str;
            AddXYLabel(HCOL4 + 2, Row + 2 + (ROW_HTsm * 6), ROW_HTsm, HCOL8 - (HCOL4 + 8), Str, ifontName, 5, "", "J");

            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", s1, "7", "LR", "TW", R1.ToString(), COL05.ToString(), WD3.ToString(), "16", "5");



            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_consignee_name.ToString(), ifontName, ifontSize, "LR", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, (AddrList.Count > 0) ? AddrList[0] : "", ifontName, ifontSize, "LR", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, (AddrList.Count > 1) ? AddrList[1] : "", ifontName, ifontSize, "LR", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, (AddrList.Count > 2) ? AddrList[2] : "", ifontName, ifontSize, "LR", "", R1, COL01, WD1 + WD2 + 1, 16, 0);

            //new row
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, (AddrList.Count > 3) ? AddrList[3] : "", ifontName, ifontSize, "LR", "", R1, COL01, WD1 + WD2 + 1, 16, 0);


            if (BL_TYPE == "MAWB")
            {
                Row += ROW_HT; R1++;
                AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL4 - HCOL1, "Issuing Carrier's Agent Name and City", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                AddXYLabel(HCOL4, Row, ROW_HTsm, HCOL8 - HCOL4, "Accounting information", ifontName, ifontSizesm, "TLR", "", R1, COL05, WD3, 16, 0);
                Row += ROW_HTsm; R1++;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_issu_agnt_name.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                if (mRow.bl_ex_works.ToString() == "Y")
                    Str = "EX-WORKS";
                else
                    Str = (mRow.bl_frt_status.ToString().Trim() == "P") ? "FREIGHT PREPAID" : (mRow.bl_frt_status.ToString().Trim() == "C") ? "FREIGHT COLLECT" : "";
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, Str, ifontName, ifontSize, "LR", "C", R1, COL05, WD3, 16, 0); //CONSOL NO : MblRec.mbld_refno
                Row += ROW_HT; R1++;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_issu_agnt_city.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info1.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);

                Row += ROW_HT; R1++;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, "", ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info2.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);

                Row += ROW_HT; R1++;
                AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL2 - HCOL1, "Agent's IATA Code", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1, 16, 0);
                AddXYLabel(HCOL2, Row, ROW_HTsm, HCOL4 - HCOL2, "Account No.", ifontName, ifontSizesm, "TL", "", R1, COL04, WD2, 16, 0);
                AddXYLabel(HCOL4, Row, ROW_HTsm, HCOL8 - HCOL4, mRow.bl_account_info3.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);
                Row += ROW_HTsm; R1++;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, mRow.bl_iata_code.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, mRow.bl_acc_no.ToString(), ifontName, ifontSize, "L", "", R1, COL04, WD2, 16, 0);
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info4.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);

                Row += ROW_HT; R1++;
                AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL4 - HCOL1, "Airport of Departure(Addr. of first Carrier) and requested Routing", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                AddXYLabel(HCOL4, Row, ROW_HTsm, (HCOL4 + 160) - HCOL4, "Reference Number", ifontName, ifontSizesm, "LT", "C", R1, COL05, WD3, 16, 0);
                AddXYLabel(HCOL4 + 159, Row, ROW_HTsm, (HCOL6 + 55) - (HCOL4 + 158), "Optional Shipping Information", ifontName, ifontSizesm, "lTBr", "", R1, COL05, WD3, 16, 0);
                AddXYLabel(HCOL6 + 55, Row, ROW_HTsm, (HCOL8) - (HCOL6 + 55), "", ifontName, ifontSize, "TR", "", R1, COL05, WD3, 16, 0);

                Row += ROW_HTsm; R1++;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_pol.ToString(), ifontName, ifontSize, "LB", "", R1, COL01, WD1 + WD2 + 1, 16, 0); //MblRec.mbl_pol_name

                //AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, Str, ifontName, ifontSize, "LRB", "C", R1, COL05, WD3, 16, 0);
                //addList("EXTRA_XL_TEXT", "0", "0", "0", "0", Str, "7", "LRB", "C", R1.ToString(), COL05.ToString(), WD3.ToString(), "16", "0");
                AddXYLabel(HCOL4, Row, ROW_HT, (HCOL4 + 160) - HCOL4, "", ifontName, ifontSize, "L", "", R1, COL05, WD3, 16, 0);
                AddXYLabel(HCOL4 + 160, Row, ROW_HT, (HCOL6 + 55) - (HCOL4 + 158), "", ifontName, ifontSize, "L", "", R1, COL05, WD3, 16, 0);
                AddXYLabel(HCOL6 + 55, Row, ROW_HT, (HCOL8) - (HCOL6 + 55), "", ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);
            }
            else
            {
                if (PrintFormatName.StartsWith("BANSARD"))
                {
                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL2 - HCOL1, "Issuing Consolidators's Agent Name", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL2, Row, ROW_HTsm, HCOL4 - HCOL2, "Agent's IATA Code", ifontName, ifontSizesm, "TL", "", R1, COL04, WD2, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HTsm, HCOL8 - HCOL4, "Accounting information", ifontName, ifontSizesm, "TLR", "", R1, COL05, WD3, 16, 0);
                    Row += ROW_HTsm; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, mRow.bl_issu_agnt_name.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, mRow.bl_iata_code.ToString(), ifontName, ifontSize, "L", "", R1, COL04, WD2, 16, 0);
                    if (mRow.bl_ex_works.ToString() == "Y")
                        Str = "EX-WORKS";
                    else
                        Str = (mRow.bl_frt_status.ToString().Trim() == "P") ? "FREIGHT PREPAID" : (mRow.bl_frt_status.ToString().Trim() == "C") ? "FREIGHT COLLECT" : "";
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, Str, ifontName, ifontSize, "LR", "C", R1, COL05, WD3, 16, 0); //CONSOL NO : MblRec.mbld_refno
                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_issu_agnt_city.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, "Account No.", ifontName, ifontSizesm, "TL", "", R1, COL04, WD2, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info1.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);
                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, "", ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, mRow.bl_acc_no.ToString(), ifontName, ifontSize, "L", "", R1, COL04, WD2, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info2.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);



                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL2 - HCOL1, "IATA CARRIER", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1, 16, 0);
                    AddXYLabel(HCOL2, Row, ROW_HTsm, HCOL4 - HCOL2, "Master Air Waybill Number", ifontName, ifontSizesm, "TL", "", R1, COL04, WD2, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HTsm, HCOL8 - HCOL4, mRow.bl_account_info3.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);
                    Row += ROW_HTsm; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, mRow.bl_iata_carrier.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, mRow.bl_mbl_no.ToString(), ifontName, ifontSize, "L", "", R1, COL04, WD2, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info4.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);

                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL4 - HCOL1, "Airport of Departure(Addr. of first Carrier) and requested Routing", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HTsm, (HCOL4 + 160) - HCOL4, "Reference Number", ifontName, ifontSizesm, "LT", "C", R1, COL05, WD3, 16, 0);
                    AddXYLabel(HCOL4 + 159, Row, ROW_HTsm, (HCOL6 + 55) - (HCOL4 + 158), "Optional Shipping Information", ifontName, ifontSizesm, "lTBr", "", R1, COL05, WD3, 16, 0);
                    AddXYLabel(HCOL6 + 55, Row, ROW_HTsm, (HCOL8) - (HCOL6 + 55), "", ifontName, ifontSize, "TR", "", R1, COL05, WD3, 16, 0);

                    Row += ROW_HTsm; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_pol.ToString(), ifontName, ifontSize, "LB", "", R1, COL01, WD1 + WD2 + 1, 16, 0); //MblRec.mbl_pol_name

                    //AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, Str, ifontName, ifontSize, "LRB", "C", R1, COL05, WD3, 16, 0);
                    //addList("EXTRA_XL_TEXT", "0", "0", "0", "0", Str, "7", "LRB", "C", R1.ToString(), COL05.ToString(), WD3.ToString(), "16", "0");
                    AddXYLabel(HCOL4, Row, ROW_HT, (HCOL4 + 160) - HCOL4, "", ifontName, ifontSize, "L", "", R1, COL05, WD3, 16, 0);
                    AddXYLabel(HCOL4 + 160, Row, ROW_HT, (HCOL6 + 55) - (HCOL4 + 158), "", ifontName, ifontSize, "L", "", R1, COL05, WD3, 16, 0);
                    AddXYLabel(HCOL6 + 55, Row, ROW_HT, (HCOL8) - (HCOL6 + 55), "", ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);
                }
                else
                {
                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL4 - HCOL1, "Issuing Consolidator's Agent Name", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HTsm, HCOL6 - 20 - HCOL4, "IATA Carrier", ifontName, ifontSizesm, "TL", "", R1, COL05, WD3, 16, 0);
                    AddXYLabel(HCOL6 - 20, Row, ROW_HTsm, HCOL8 - (HCOL6 - 20), "Master Air Waybill Number", ifontName, ifontSizesm, "TLR", "", R1, COL05, WD3, 16, 0);

                    Row += ROW_HTsm; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_issu_agnt_name.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL6 - 20 - HCOL4, mRow.bl_iata_carrier.ToString(), ifontName, ifontSize, "BL", "C", R1, COL05, WD3, 16, 0);
                    AddXYLabel(HCOL6 - 20, Row, ROW_HT, HCOL8 - (HCOL6 - 20), mRow.bl_mbl_no.ToString(), ifontName, ifontSize + 2, "BLR", "CB", R1, COL05, WD3, 16, 0);

                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_issu_agnt_city.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, "Accounting information", ifontName, ifontSizesm, "LR", "", R1, COL05, WD3, 16, 0);
                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, "", ifontName, ifontSize, "L", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    if (mRow.bl_ex_works.ToString() == "Y")
                        Str = "EX-WORKS";
                    else
                        Str = (mRow.bl_frt_status.ToString().Trim() == "P") ? "FREIGHT PREPAID" : (mRow.bl_frt_status.ToString().Trim() == "C") ? "FREIGHT COLLECT" : "";
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, Str, ifontName, ifontSize, "LR", "C", R1, COL05, WD3, 16, 0);

                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL2 - HCOL1, "Agent's IATA Code", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1, 16, 0);
                    AddXYLabel(HCOL2, Row, ROW_HTsm, HCOL4 - HCOL2, "Account No.", ifontName, ifontSizesm, "TL", "", R1, COL04, WD2, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HTsm, HCOL8 - HCOL4, mRow.bl_account_info1.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);
                    Row += ROW_HTsm; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, mRow.bl_iata_code.ToString(), ifontName, ifontSize, "L", "", R1, COL01, WD1, 16, 0);
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, mRow.bl_acc_no.ToString(), ifontName, ifontSize, "L", "", R1, COL04, WD2, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info2.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);

                    Row += ROW_HT; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL4 - HCOL1, "Airport of Departure(Addr. of first Carrier) and requested Routing", ifontName, ifontSizesm, "TL", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info3.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);

                    Row += ROW_HTsm; R1++;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL4 - HCOL1, mRow.bl_pol.ToString(), ifontName, ifontSize, "LB", "", R1, COL01, WD1 + WD2 + 1, 16, 0);
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL8 - HCOL4, mRow.bl_account_info4.ToString(), ifontName, ifontSize, "LR", "", R1, COL05, WD3, 16, 0);
                }
            }



            // END OF FIRST PART

            // START Of PART 2

            COL01 = 2;
            COL02 = COL01 + 6;
            COL03 = COL02 + 8;
            COL04 = COL03 + 1;
            COL05 = COL04 + 11;
            COL06 = COL05 + 2;
            COL07 = COL06 + 5;
            COL08 = COL07 + 5;
            COL09 = COL08 + 5;
            COL10 = COL09 + 5;  // Currency
            COL11 = COL10 + 5;
            COL12 = COL11 + 3;
            COL13 = COL12 + 3;
            COL14 = COL13 + 3;
            COL15 = COL14 + 3;
            COL16 = COL15 + 3;
            COL17 = COL16 + 13;
            COL18 = COLTOT;


            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 40 - HCOL1, "to", ifontName, ifontSizesm, "L", "", R1, COL01, COL02 - COL01 - 1, 0, 0);
            AddXYLabel(HCOL1 + 40, Row, ROW_HTsm, HCOL1 + 110 - (HCOL1 + 40), "By first Carrier", ifontName, ifontSizesm, "L", "", R1, COL02, COL03 - COL02 - 1, 0, 0);
            AddXYLabel(HCOL1 + 105, Row, ROW_HTsm, HCOL2 + 40 - (HCOL1 + 115), "Routing and Destination", ifontName, ifontSizesm, "Blr", "");
            PrintEasy("Routing and Destination", ROW_HTsm, ifontSizesm, "B", "", COL03, COL04, COL05);
            AddXYLabel(HCOL2 + 40, Row, ROW_HTsm, HCOL2 + 80 - (HCOL2 + 40), "to", ifontName, ifontSizesm, "L", "", R1, COL06, COL07 - COL06 - 1, 0, 0);
            AddXYLabel(HCOL2 + 80, Row, ROW_HTsm, HCOL2 + 120 - (HCOL2 + 80), "by", ifontName, ifontSizesm, "L", "", R1, COL07, COL08 - COL07 - 1, 0, 0);
            AddXYLabel(HCOL2 + 120, Row, ROW_HTsm, HCOL2 + 160 - (HCOL2 + 120), "to", ifontName, ifontSizesm, "L", "", R1, COL08, COL09 - COL08 - 1, 0, 0);
            AddXYLabel(HCOL2 + 160, Row, ROW_HTsm, HCOL4 - (HCOL2 + 160), "by", ifontName, ifontSizesm, "L", "", R1, COL09, COL10 - COL09 - 1, 0, 0);
            AddXYLabel(HCOL4, Row, ROW_HTsm, (HCOL4 + 38) - (HCOL4), "Currency", ifontName, ifontSizesm, "TL", "", R1, COL10, COL11 - COL10 - 1, 0, 0);
            AddXYLabel(HCOL4 + 36, Row, ROW_HTsm, (HCOL4 + 60) - (HCOL4 + 35), "Chgs.", ifontName, ifontSizesm, "TL", "C", R1, COL11, COL12 - COL11 - 1, 0, 0);

            AddXYLabel(HCOL4 + 60, Row, ROW_HTsm, (HCOL4 + 110) - (HCOL4 + 60), "WT/VAL", ifontName, ifontSizesm, "TL", "C", R1, COL12, COL14 - COL12 - 1, 0, 0);
            AddXYLabel(HCOL4 + 110, Row, ROW_HTsm, (HCOL4 + 160) - (HCOL4 + 110), "OTHER", ifontName, ifontSizesm, "TL", "C", R1, COL14, COL16 - COL14 - 1, 0, 0);

            AddXYLabel(HCOL4 + 160, Row, ROW_HTsm, (HCOL6 + 55) - (HCOL4 + 158), "Declared Value for Carriage", ifontName, ifontSizesm, "TL", "", R1, COL16, COL17 - COL16 - 1, 0, 0);
            AddXYLabel(HCOL6 + 55, Row, ROW_HTsm, (HCOL8) - (HCOL6 + 55), "Declared Value for Customs", ifontName, ifontSizesm, "TLR", "", R1, COL17, COL18 - COL17 - 1, 0, 0);

            Row += ROW_HTsm; R1++;

            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 40 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, COL01, 0, 0, 0);
            AddXYLabel(HCOL1 + 40, Row, ROW_HTsm, HCOL1 + 110 - (HCOL1 + 40), "", ifontName, ifontSizesm, "L", "", R1, COL02, 0, 0, 0);
            AddXYLabel(HCOL1 + 110, Row, ROW_HTsm, HCOL2 + 40 - (HCOL1 + 110), "", ifontName, ifontSizesm, "", "");
            AddXYLabel(HCOL2 + 40, Row, ROW_HTsm, HCOL2 + 80 - (HCOL2 + 40), "", ifontName, ifontSizesm, "L", "", R1, COL06, 0, 0, 0);
            AddXYLabel(HCOL2 + 80, Row, ROW_HTsm, HCOL2 + 120 - (HCOL2 + 80), "", ifontName, ifontSizesm, "L", "", R1, COL07, 0, 0, 0);
            AddXYLabel(HCOL2 + 120, Row, ROW_HTsm, HCOL2 + 160 - (HCOL2 + 120), "", ifontName, ifontSizesm, "L", "", R1, COL08, 0, 0, 0);
            AddXYLabel(HCOL2 + 160, Row, ROW_HTsm, HCOL4 - (HCOL2 + 160), "", ifontName, ifontSizesm, "L", "", R1, COL09, 0, 0, 0);
            AddXYLabel(HCOL4, Row, ROW_HTsm, (HCOL4 + 35) - (HCOL4), "", ifontName, ifontSizesm, "L", "", R1, COL10, 0, 0, 0);
            AddXYLabel(HCOL4 + 36, Row, ROW_HTsm, (HCOL4 + 60) - (HCOL4 + 35), "Code", ifontName, ifontSizesm, "L", "C", R1, COL11, 0, 0, 0);
            AddXYLabel(HCOL4 + 60, Row, ROW_HTsm, (HCOL4 + 85) - (HCOL4 + 60), "PPD", ifontName, ifontSizesm, "TL", "C", R1, COL12, COL13 - COL12 - 1, 0, 0);
            AddXYLabel(HCOL4 + 85, Row, ROW_HTsm, (HCOL4 + 110) - (HCOL4 + 85), "COLL", ifontName, ifontSizesm, "TL", "C", R1, COL13, COL14 - COL13 - 1, 0, 0);
            AddXYLabel(HCOL4 + 110, Row, ROW_HTsm, (HCOL4 + 135) - (HCOL4 + 110), "PPD", ifontName, ifontSizesm, "TL", "C", R1, COL14, COL15 - COL14 - 1, 0, 0);
            AddXYLabel(HCOL4 + 135, Row, ROW_HTsm, (HCOL4 + 160) - (HCOL4 + 135), "COLL", ifontName, ifontSizesm, "TL", "C", R1, COL15, COL16 - COL15 - 1, 0, 0);
            AddXYLabel(HCOL4 + 160, Row, ROW_HTsm, (HCOL6 + 60) - (HCOL4 + 160), "", ifontName, ifontSizesm, "L", "", R1, COL16, 0, 0, 0);
            AddXYLabel(HCOL6 + 55, Row, ROW_HTsm, (HCOL8) - (HCOL6 + 55), "", ifontName, ifontSizesm, "LR", "", R1, COL17, COL18 - COL17 - 1, 13, 0);




            Row += ROW_HTsm; R1++;

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 40 - HCOL1,  mRow.bl_to1.ToString(), ifontName, ifontSize, "LB", "", R1, COL01, COL02 - COL01 - 1, 0, 0); //mbl_to_port1
            AddXYLabel(HCOL1 + 40, Row, ROW_HT, HCOL2 + 40 - (HCOL1 + 40), mRow.bl_by1.ToString(), ifontName, ifontSize, "LB", "", R1, COL02, COL06 - COL02 - 1, 0, 0);//mbl_by_carrier1
            AddXYLabel(HCOL2 + 40, Row, ROW_HT, HCOL2 + 80 - (HCOL2 + 40), mRow.bl_to2.ToString(), ifontName, ifontSize, "LB", "", R1, COL06, COL07 - COL06 - 1, 0, 0);//mbl_to_port2
            AddXYLabel(HCOL2 + 80, Row, ROW_HT, HCOL2 + 120 - (HCOL2 + 80), mRow.bl_by2.ToString(), ifontName, ifontSize, "LB", "", R1, COL07, COL08 - COL07 - 1, 0, 0);//mbl_by_carrier2
            AddXYLabel(HCOL2 + 120, Row, ROW_HT, HCOL2 + 160 - (HCOL2 + 120), mRow.bl_to3.ToString(), ifontName, ifontSize, "LB", "", R1, COL08, COL09 - COL08 - 1, 0, 0);
            AddXYLabel(HCOL2 + 160, Row, ROW_HT, HCOL4 - (HCOL2 + 160), mRow.bl_by3.ToString(), ifontName, ifontSize, "LB", "", R1, COL09, COL10 - COL09 - 1, 0, 0);
            AddXYLabel(HCOL4, Row, ROW_HT, (HCOL4 + 35) - (HCOL4), mRow.bl_currency.ToString(), ifontName, ifontSize, "L", "C", R1, COL10, COL11 - COL10 - 1, 0, 0); //MblRec.mbl_currency
            AddXYLabel(HCOL4 + 36, Row, ROW_HT, (HCOL4 + 60) - (HCOL4 + 35), mRow.bl_frt_status.ToString().Trim() + mRow.bl_oc_status.ToString().Trim(), ifontName, ifontSize, "L", "C", R1, COL11, COL12 - COL11 - 1, 0, 0);
            AddXYLabel(HCOL4 + 60, Row, ROW_HT, (HCOL4 + 85) - (HCOL4 + 60), (mRow.bl_frt_status.ToString().Trim() == "P") ? "X" : "", ifontName, ifontSize, "TL", "C", R1, COL12, COL13 - COL12 - 1, 0, 0);
            AddXYLabel(HCOL4 + 85, Row, ROW_HT, (HCOL4 + 110) - (HCOL4 + 85), (mRow.bl_frt_status.ToString().Trim() == "C") ? "X" : "", ifontName, ifontSize, "LT", "C", R1, COL13, COL14 - COL13 - 1, 0, 0);
            AddXYLabel(HCOL4 + 110, Row, ROW_HT, (HCOL4 + 135) - (HCOL4 + 110), (mRow.bl_oc_status.ToString().Trim() == "P") ? "X" : "", ifontName, ifontSize, "LT", "C", R1, COL14, COL15 - COL14 - 1, 0, 0);
            AddXYLabel(HCOL4 + 135, Row, ROW_HT, (HCOL4 + 160) - (HCOL4 + 135), (mRow.bl_oc_status.ToString().Trim() == "C") ? "X" : "", ifontName, ifontSize, "LT", "C", R1, COL15, COL16 - COL15 - 1, 0, 0);
            AddXYLabel(HCOL4 + 160, Row, ROW_HT, (HCOL6 + 55) - (HCOL4 + 160), mRow.bl_carriage_value.ToString(), ifontName, ifontSize, "L", "", R1, COL16, COL17 - COL16 - 1, 0, 0);
            AddXYLabel(HCOL6 + 55, Row, ROW_HT, (HCOL8) - (HCOL6 + 55), mRow.bl_customs_value.ToString(), ifontName, ifontSize, "LR", "", R1, COL17, COL18 - COL17 - 1, 0, 0);


            // END Of PART 2


            COL01 = 2;
            COL02 = COL01 + 23;
            COL03 = COL02 + 6;
            COL04 = COL03 + 1;
            COL05 = COL04 + 4;
            COL06 = COL05 + 7;
            COL07 = COL06 + 1;
            COL08 = COL07 + 6; // amt of insurance
            COL09 = COL08 + 10;
            COL10 = COLTOT;

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "Airport of Destination", ifontName, ifontSizesm, "L", "", R1, COL01, COL02 - COL01 - 1, 0, 0);
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL2 + 55 - (HCOL2), BL_TYPE == "MAWB" ? "" : "Flight/Date", ifontName, ifontSizesm, "L", "", R1, COL02, COL03 - COL02 - 1, 0, 0);//Flight/Date
            AddXYLabel(HCOL2 + 50, Row, ROW_HTsm, HCOL3 + 50 - (HCOL2 + 50), BL_TYPE == "MAWB" ? "Requested Flight/Date" : "For Carrier Use Only", ifontName, ifontSizesm, "Blr", "C");//For Carrier Use only
            PrintEasy("Requested Flight/Date", ROW_HTsm, ifontSizesm, "B", "", COL03, COL04, COL06);//For Carriers Use Only
            AddXYLabel(HCOL3 + 53, Row, ROW_HT, HCOL4 - (HCOL3 + 50), BL_TYPE == "MAWB" ? "" : "Flight/Date", ifontName, ifontSizesm, "", "", R1, COL07, COL08 - COL07 - 1, 0, 0);//Flight/Date
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL4 + 110 - HCOL4, "Amount of Insurance", ifontName, ifontSizesm, "TL", "", R1, COL08, COL09 - COL08 - 1, 0, 0);

            s1 = "";
            Str = "INSURANCE - if carrier offers insurance and such insurance is ";
            s1 += Str;
            AddXYLabel(HCOL4 + 110, Row, 10, (HCOL8) - (HCOL4 + 110), Str, ifontName, 6, "", "");
            Str = "requested in accordance  with  conditions on reverse hereof indicate ";
            s1 += Str;
            AddXYLabel(HCOL4 + 110, Row + 9, 10, (HCOL8) - (HCOL4 + 110), Str, ifontName, 6, "", "");
            Str = "amount to be insured in figures in box  marked \"Amount of Insurance\".";
            s1 += Str;
            AddXYLabel(HCOL4 + 110, Row + 18, 10, (HCOL8) - (HCOL4 + 110), Str, ifontName, 6, "", "");
            AddXYLabel(HCOL4 + 110, Row, ROW_HT, (HCOL8) - (HCOL4 + 110), "", ifontName, ifontSizesm, "LTR", "");
            WD3 = COL10 - COL09 - 1;
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", s1, "7", "LRBT", "TW", R1.ToString(), COL09.ToString(), WD3.ToString(), "0", "1");

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, mRow.bl_pod.ToString(), ifontName, ifontSize, "L", "", R1, COL01, COL02 - COL01 - 1, 0, 0);//mbl_pod_name
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2 + 10, mRow.bl_flight1.ToString(), ifontName, ifontSize - 1, "L", "", R1, COL02, COL03 - COL02 - 1, 0, 0);
            DrawVLine(HCOL3 + 8, Row - 3, ROW_HT + 3);

            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "7", "L", "", R1.ToString(), COL05.ToString(), "0", "0", "0");

            AddXYLabel(HCOL3 + 8, Row, ROW_HT, HCOL4 - HCOL3, mRow.bl_flight2.ToString(), ifontName, ifontSize - 1, "", "", R1, COL06, COL08 - COL06 - 1, 0, 0);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL4 + 110 - HCOL4, mRow.bl_ins_amt.ToString(), ifontName, ifontSize, "L", "", R1, COL08, COL09 - COL08 - 1, 0, 0);
            AddXYLabel(HCOL4 + 110, Row, ROW_HT, (HCOL8) - (HCOL4 + 110), "", ifontName, ifontSize, "LR", "");



            WD3 = COLTOT - COL01 - 1;
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL8 - HCOL1, "Handling Information", ifontName, ifontSizesm, "LTR", "", R1, COL01, WD3, 0, 0);
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL8 - HCOL1, "", ifontName, ifontSize, "LR", "");
            float RoutRow = Row - 4;
            AddXYLabel(HCOL1, RoutRow, ROW_HT - 3, HCOL8 - HCOL1, mRow.bl_hand_info1.ToString(), ifontName, ifontSize - 1, "", "");
            RoutRow += ROW_HT - 3;
            AddXYLabel(HCOL1, RoutRow, ROW_HT - 3, HCOL8 - 100 - HCOL1, mRow.bl_hand_info2.ToString(), ifontName, ifontSize - 1, "", "");
            RoutRow += ROW_HT - 3;
            AddXYLabel(HCOL1, RoutRow, ROW_HT - 3, HCOL8 - 100 - HCOL1, mRow.bl_hand_info3.ToString(), ifontName, ifontSize - 1, "", "");

            R1++; addList("EXTRA_XL_TEXT", "0", "0", "0", "0", mRow.bl_hand_info1.ToString(), "8", "LR", "", R1.ToString(), COL01.ToString(), WD3.ToString(), "10", "0");
            R1++; addList("EXTRA_XL_TEXT", "0", "0", "0", "0", mRow.bl_hand_info2.ToString(), "8", "LR", "", R1.ToString(), COL01.ToString(), WD3.ToString(), "10", "0");
            R1++; addList("EXTRA_XL_TEXT", "0", "0", "0", "0", mRow.bl_hand_info3.ToString(), "8", "LR", "", R1.ToString(), COL01.ToString(), WD3.ToString(), "10", "0");


            COL02 = 50;
            COL03 = COL16;
            COL04 = COL17;

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 90 - HCOL1, "", ifontName, ifontSize, "L", "", R1, COL01, 0, 0, 0);
            AddXYLabel(HCOL8 - 100, Row, ROW_HT, HCOL8 - (HCOL8 - 100), "SCI", ifontName, ifontSize - 2, "LRT", "C", R1, COL04, COLTOT - COL04 - 1, 12, 0);

            //Row += ROW_HT; R1++;
            //AddXYLabel( HCOL1, Row, ROW_HTsm, HCOL5 + 60 - HCOL1, "These commodities,technology or software were exported from the United States", ifontName, 9, "L", "", R1, COL01, 0, 0, 0);
            //AddXYLabel( HCOL5 + 60, Row, ROW_HTsm, HCOL8 - 100 - (HCOL5 + 60), "Diversion contrary to", ifontName, 9, "", "", R1, COL03, 0, 0, 0);
            //AddXYLabel( HCOL8 - 100, Row, ROW_HTsm, HCOL8 - (HCOL8 - 100), "", ifontName, ifontSize, "LR", "C", R1, COL04, COLTOT - COL04 - 1, 11, 0);

            Row += ROW_HTsm; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL8 - 100 - HCOL1, "", ifontName, 5, "L", "");
            if (BL_TYPE == "MAWB")
                Str = "(For USA use only) These commodities, technology or software were exported from the United States in accordance with the Export Administration Regulations. Diversion contrary to USA law prohibited.";
            else
                Str = "(For USA use only) These commodities licensed by USA for ultimate destination .............................................................. Diversion contrary to USA law prohibited";
            AddXYLabel(HCOL1, Row + 2, ROW_HT, HCOL8 - 100 - HCOL1, Str, ifontName, 5, "", "", R1, COL01, 0, 0, 0);
            // AddXYLabel(HCOL5 - 140, Row, ROW_HT, HCOL5 + 60 - (HCOL5 - 60), "", ifontName, ifontSize, "", "", R1, COL02, 0, 0, 0);//MblRec.mbld_country_name
            // AddXYLabel( HCOL5 + 60, Row, ROW_HT, HCOL8 - 100 - (HCOL5 + 60), "U.S. law prohibited.", ifontName, 9, "", "", R1, COL03, 0, 0, 0);

            AddXYLabel(HCOL8 - 100, Row, ROW_HT, HCOL8 - (HCOL8 - 100), "", ifontName, ifontSize, "LR", "", R1, COL04, COLTOT - COL04 - 1, 12, 0);
            Row += ROW_HT;


            dCol1 = 20; dCol2 = 60; dCol3 = 136; dCol4 = 148; dCol5 = 158; dCol6 = 170; dCol7 = 250; //dCol3 = 140;
            dCol8 = 260; dCol9 = 340; dCol10 = 350; dCol11 = 430; dCol12 = 440; dCol13 = 530; dCol14 = 540;

            DrawHLine(dCol1, Row, HCOL8 - dCol1);
            DrawVLine(dCol1, Row, ROW_HT * 19 + ROW_HTsm * 2);
            AddXYLabel(dCol1, Row, ROW_HTsm, dCol2 - dCol1, "No of", ifontName, ifontSizesm, "", "C");

            DrawVLine(dCol2, Row, ROW_HT * 19 + ROW_HTsm * 2);
            AddXYLabel(dCol2, Row + 5, ROW_HTsm, dCol3 - dCol2, "Gross", ifontName, ifontSizesm, "", "C");
            DrawVLine(dCol3, Row, ROW_HT * 19 + ROW_HTsm * 2);
            AddXYLabel(dCol3, Row + 5, ROW_HTsm, dCol4 - dCol3, "kg", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol4, Row, ROW_HTsm, dCol5 - dCol4, "", ifontName, ifontSizesm, "", "");
            SetFillRectangle(dCol4, Row, ROW_HT * 19 + ROW_HTsm * 2, 10);
            AddXYLabel(dCol5, Row, ROW_HTsm, dCol6 - dCol5, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol6, Row, ROW_HTsm, dCol7 - dCol6, "", ifontName, ifontSizesm, "B", "C");
            AddXYLabel(dCol6 - 10, Row, ROW_HTsm, dCol7 - dCol6, "Rate Class", ifontName, ifontSizesm, "", "L");
            AddXYLabel(dCol7, Row, ROW_HTsm, dCol8 - dCol7, "", ifontName, ifontSizesm, "", "");
            SetFillRectangle(dCol7, Row, ROW_HT * 19 + ROW_HTsm * 2, 10);
            AddXYLabel(dCol8, Row + 5, ROW_HTsm, dCol9 - dCol8, "Chargeable", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol9, Row, ROW_HTsm, dCol10 - dCol9, "", ifontName, ifontSizesm, "", "");
            SetFillRectangle(dCol9, Row, ROW_HT * 19 + ROW_HTsm * 2, 10);
            AddXYLabel(dCol10, Row + 5, ROW_HTsm, dCol11 - dCol10, "Rate", ifontName, ifontSizesm, "", "L");
            AddXYLabel(dCol11, Row, ROW_HTsm, dCol12 - dCol11, "", ifontName, ifontSizesm, "", "");
            SetFillRectangle(dCol11, Row, ROW_HT * 19 + ROW_HTsm * 2, 10);
            AddXYLabel(dCol12, Row + 10, ROW_HTsm, dCol13 - dCol12, "Total", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol13, Row, ROW_HTsm, dCol14 - dCol13, "", ifontName, ifontSizesm, "", "");
            SetFillRectangle(dCol13, Row, ROW_HT * 19 + ROW_HTsm * 2, 10);
            AddXYLabel(dCol14, Row + 5, ROW_HTsm, HCOL8 - dCol14, "Nature and Quantity of Goods", ifontName, ifontSizesm, "", "C");
            DrawVLine(HCOL8, Row, ROW_HT * 19 + ROW_HTsm * 2);

            Row += ROW_HTsm;
            AddXYLabel(dCol1, Row, ROW_HTsm, dCol2 - dCol1, "Pieces", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol2, Row + 5, ROW_HTsm, dCol3 - dCol2, "Weight", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol3, Row + 5, ROW_HTsm, dCol4 - dCol3, "lb", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol4, Row, ROW_HTsm, dCol5 - dCol4, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol5, Row, ROW_HTsm, dCol6 - dCol5, "", ifontName, ifontSizesm, "", "");
            DrawVLine(dCol6, Row, ROW_HT * 19 + ROW_HTsm);
            AddXYLabel(dCol6, Row, ROW_HTsm, dCol7 - dCol6, "Commodity", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol7, Row, ROW_HTsm, dCol8 - dCol7, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol8, Row + 5, ROW_HTsm, dCol9 - dCol8, "Weight", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol9, Row, ROW_HTsm, dCol10 - dCol9, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol10, Row + 5, ROW_HTsm, dCol11 - dCol10, "", ifontName, ifontSizesm, "", "C");
            AddXYLabel(dCol11, Row, ROW_HTsm, dCol12 - dCol11, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol12, Row, ROW_HTsm, dCol13 - dCol12, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol13, Row, ROW_HTsm, dCol14 - dCol13, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol14, Row + 5, ROW_HTsm, HCOL8 - dCol14, "(Incl. Dimensions or Volume)", ifontName, ifontSizesm, "", "C");

            Row += ROW_HTsm;

            AddXYLabel(dCol1, Row, ROW_HTsm, dCol2 - dCol1, "RCP", ifontName, ifontSizesm, "B", "C");
            AddXYLabel(dCol2, Row, ROW_HTsm, dCol3 - dCol2, "", ifontName, ifontSizesm, "B", "");
            AddXYLabel(dCol3, Row, ROW_HTsm, dCol4 - dCol3, "", ifontName, ifontSizesm, "B", "");
            AddXYLabel(dCol4, Row, ROW_HTsm, dCol5 - dCol4, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol5, Row, ROW_HTsm, dCol6 - dCol5, "", ifontName, ifontSizesm, "B", "");
            AddXYLabel(dCol6, Row, ROW_HTsm, dCol7 - dCol6, "Item No", ifontName, ifontSizesm, "B", "C");
            AddXYLabel(dCol7, Row, ROW_HTsm, dCol8 - dCol7, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol8, Row, ROW_HTsm, dCol9 - dCol8, "", ifontName, ifontSizesm, "B", "");
            AddXYLabel(dCol9, Row, ROW_HTsm, dCol10 - dCol9, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol10, Row, ROW_HTsm, dCol11 - dCol10, "Charge", ifontName, ifontSizesm, "B", "R");

            AddXYLabel(dCol11, Row, ROW_HTsm, dCol12 - dCol11, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol12, Row, ROW_HTsm, dCol13 - dCol12, "", ifontName, ifontSizesm, "B", "");
            AddXYLabel(dCol13, Row, ROW_HTsm, dCol14 - dCol13, "", ifontName, ifontSizesm, "", "");
            AddXYLabel(dCol14, Row, ROW_HTsm, HCOL8 - dCol14, "", ifontName, ifontSizesm, "B", "");

            DrawVLine(dCol10 - 20, Row - 18, 25, 25, 70);

            Row += ROW_HT;


            DescStartRow = Row;

            // start newly added
            WriteDescription();
            Row = DescStartRow;
            // end newly added 

            Row += ROW_HT * 17;
            Row += ROW_HT;
            DrawHLine(HCOL1, Row, HCOL8 - HCOL1);
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 10 - (HCOL1), "", ifontName, ifontSizesm, "L", "C");
            AddXYLabel(HCOL1 + 10, Row, ROW_HTsm, HCOL1 + 95 - (HCOL1 + 10 + 10), "Prepaid", ifontName, ifontSizesm, "Blr", "C");
            AddXYLabel(HCOL1 + 95, Row, ROW_HTsm, HCOL1 + 205 - (HCOL1 + 95 + 10), "Weight Charge", ifontName, ifontSizesm, "Blr", "C");
            AddXYLabel(HCOL1 + 205, Row, ROW_HTsm, HCOL1 + 300 - (HCOL1 + 205 + 10), "Collect", ifontName, ifontSizesm, "Blr", "C");
            AddXYLabel(HCOL1 + 300, Row, ROW_HTsm, HCOL8 - (HCOL1 + 300), "Other Charges", ifontName, ifontSizesm, "LR", "");

            COL01 = 2;
            COL02 = COL01 + 1;
            COL03 = COL02 + 1;

            COL04 = COL03 + 10;
            COL05 = COL04 + 1;
            COL06 = COL05 + 1;
            COL07 = COL06 + 10;
            COL08 = COL07 + 1;
            COL09 = COL08 + 1;
            COL10 = COL09 + 10;
            COL11 = COL10 + 2;
            COL12 = COL11 + 1;
            COL13 = COL12 + 1;
            //COL10 = COLTOT;
            R1++;


            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "9", "L", "", R1.ToString(), COL01.ToString(), "0", "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "9", "l", "", R1.ToString(), COL02.ToString(), "0", "12");
            WD1 = COL04 - COL03 - 1;
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "Prepaid", "9", "B", "C", R1.ToString(), COL03.ToString(), WD1.ToString(), "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "9", "r", "", R1.ToString(), COL04.ToString(), "0", "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "9", "l", "", R1.ToString(), COL05.ToString(), "0", "12");
            WD1 = COL07 - COL06 - 1;
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "Weight Charge", "9", "B", "C", R1.ToString(), COL06.ToString(), WD1.ToString(), "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "9", "r", "", R1.ToString(), COL07.ToString(), "0", "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "9", "l", "", R1.ToString(), COL08.ToString(), "0", "12");
            WD1 = COL10 - COL09 - 1;
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "Collect", "9", "B", "C", R1.ToString(), COL09.ToString(), WD1.ToString(), "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "9", "r", "", R1.ToString(), COL10.ToString(), "0", "12");
            WD1 = COLTOT - COL11 - 1;
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "Other Charges", "9", "LR", "", R1.ToString(), COL11.ToString(), WD1.ToString(), "12");

            Row += ROW_HTsm;
            OthChrgStartRow = Row;

            // start newly added
            WriteOtherCharges();
            Row = OthChrgStartRow;
            // end newly added 

            string frtppAsArrngd = "", frtccAsArrngd = "";
            decimal frtppAmt = 0, frtccAmt = 0;
            if (mRow.bl_frt_status.ToString().Trim() == "P")
            {
                frtppAmt = Lib.Conv2Decimal(mRow.bl_total.ToString());
                if ((mRow.bl_asarranged_shipper.ToString().Trim() == "Y") || (mRow.bl_asarranged_consignee.ToString().Trim() == "Y"))
                    frtppAsArrngd = "AS AGREED";
                else
                    frtppAsArrngd = "";
            }
            else if (mRow.bl_frt_status.ToString().Trim() == "C")
            {
                frtccAmt = Lib.Conv2Decimal(mRow.bl_total.ToString());
                if ((mRow.bl_asarranged_shipper.ToString().Trim() == "Y") || (mRow.bl_asarranged_consignee.ToString().Trim() == "Y"))
                    frtccAsArrngd = "AS AGREED";
                else
                    frtccAsArrngd = "";
            }

            if (frtppAsArrngd.Trim().Length > 0)
                Str = frtppAsArrngd;
            else
                Str = (frtppAmt > 0) ? frtppAmt.ToString() : "";


            COL01 = 2;
            COL02 = COL01 + 18;
            COL03 = COL02 + 20;
            COL04 = COL03 + 28;
            COL05 = COL04 + 20;

            WD1 = COL02 - COL01 - 1;
            WD2 = COL03 - COL02 - 1;

            ROW_HT = ROW_HT + 3;

            R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, Str, ifontName, ifontSize, "LB", "C", R1, COL01, COL02 - COL01 - 1, 16, 0);
            if (frtccAsArrngd.Trim().Length > 0)
                Str = frtccAsArrngd;
            else
                Str = (frtccAmt > 0) ? frtccAmt.ToString() : "";
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), Str, ifontName, ifontSize, "LB", "C", R1, COL02, COL03 - COL02 - 1, 16);
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL4 + 50 - (HCOL1 + 300), "", ifontName, ifontSizesm, "L", "");//other charges 
            AddXYLabel(HCOL4 + 50, Row, ROW_HT, HCOL4 + 100 - (HCOL4 + 50), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL4 + 100, Row, ROW_HT, HCOL5 + 30 - (HCOL4 + 100), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL5 + 40, Row, ROW_HT, HCOL5 + 140 - (HCOL5 + 30), "", ifontName, ifontSizesm, "", "");
            AddXYLabel(HCOL5 + 140, Row, ROW_HT, HCOL5 + 190 - (HCOL5 + 140), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL5 + 190, Row, ROW_HT, HCOL8 - (HCOL5 + 190), "", ifontName, ifontSizesm, "R", "R");

            WD3 = COLTOT - COL04 - 1;
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH1[0].ToString(), "9", "L", "", R1.ToString(), COL03.ToString(), "0", "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH2[0].ToString(), "9", "R", "", R1.ToString(), COL04.ToString(), WD3.ToString(), "12");


            Row += ROW_HT; R1++;


            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 95 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, COL01, 0, 16, 0);
            PrintEasy("Valuation Charge", ROW_HTsm, ifontSizesm, "B", "C", 15, 16, 26);
            AddXYLabel(HCOL1 + 95, Row, ROW_HTsm, HCOL1 + 205 - (HCOL1 + 95 + 10), "Valuation Charge", ifontName, ifontSizesm, "Blr", "C");
            AddXYLabel(HCOL1 + 300, Row, ROW_HTsm, HCOL4 + 50 - (HCOL1 + 300), "", ifontName, ifontSizesm, "L", "");//other charges 
            AddXYLabel(HCOL4 + 50, Row, ROW_HTsm, HCOL4 + 100 - (HCOL4 + 50), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL4 + 100, Row, ROW_HTsm, HCOL5 + 30 - (HCOL4 + 100), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL5 + 40, Row, ROW_HTsm, HCOL5 + 140 - (HCOL5 + 30), "", ifontName, ifontSizesm, "", "");
            AddXYLabel(HCOL5 + 140, Row, ROW_HTsm, HCOL5 + 190 - (HCOL5 + 140), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL5 + 190, Row, ROW_HTsm, HCOL8 - (HCOL5 + 190), "", ifontName, ifontSizesm, "R", "R");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH1[1].ToString(), "9", "L", "", R1.ToString(), COL03.ToString(), "0", "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH2[1].ToString(), "9", "R", "", R1.ToString(), COL04.ToString(), WD3.ToString(), "12");


            Row += ROW_HTsm; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, "", ifontName, ifontSize, "LB", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), "", ifontName, ifontSize, "LB", "", R1, COL02, WD2, 16, 0);
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL4 + 50 - (HCOL1 + 300), "", ifontName, ifontSizesm, "L", "");//other charges 
            AddXYLabel(HCOL4 + 50, Row, ROW_HT, HCOL4 + 100 - (HCOL4 + 50), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL4 + 100, Row, ROW_HT, HCOL5 + 30 - (HCOL4 + 100), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL5 + 40, Row, ROW_HT, HCOL5 + 140 - (HCOL5 + 30), "", ifontName, ifontSizesm, "", "");
            AddXYLabel(HCOL5 + 140, Row, ROW_HT, HCOL5 + 190 - (HCOL5 + 140), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL5 + 190, Row, ROW_HT, HCOL8 - (HCOL5 + 190), "", ifontName, ifontSizesm, "R", "R");

            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH1[2].ToString(), "9", "L", "", R1.ToString(), COL03.ToString(), "0", "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH2[2].ToString(), "9", "R", "", R1.ToString(), COL04.ToString(), WD3.ToString(), "12");

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 125 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, COL01, 0, 16, 0);
            AddXYLabel(HCOL1 + 125, Row, ROW_HTsm, HCOL1 + 175 - (HCOL1 + 125), "Tax", ifontName, ifontSizesm, "lBr", "C");
            PrintEasy("Tax", ROW_HTsm, ifontSizesm, "B", "C", 17, 18, 23);
            AddXYLabel(HCOL1 + 300, Row, ROW_HTsm, HCOL4 + 50 - (HCOL1 + 300), "", ifontName, ifontSizesm, "L", "");//other charges 
            AddXYLabel(HCOL4 + 50, Row, ROW_HTsm, HCOL4 + 100 - (HCOL4 + 50), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL4 + 100, Row, ROW_HTsm, HCOL5 + 30 - (HCOL4 + 100), "", ifontName, ifontSizesm, "", "R");


            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH1[3].ToString(), "9", "L", "", R1.ToString(), COL03.ToString(), "0", "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH2[3].ToString(), "9", "R", "", R1.ToString(), COL04.ToString(), WD3.ToString(), "12");

            // NEW Changes Start
            AddXYLabel(HCOL5 + 40, Row, ROW_HTsm, HCOL5 + 140 - (HCOL5 + 30), "", ifontName, ifontSizesm, "", "");
            AddXYLabel(HCOL5 + 140, Row, ROW_HTsm, HCOL5 + 190 - (HCOL5 + 140), "", ifontName, ifontSizesm, "", "R");
            AddXYLabel(HCOL5 + 190, Row, ROW_HTsm, HCOL8 - (HCOL5 + 190), "", ifontName, ifontSizesm, "R", "R");
            // NEW Changes End

            Row += ROW_HTsm; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, "", ifontName, ifontSize, "LB", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), "", ifontName, ifontSize, "LB", "", R1, COL02, WD2, 16, 0);
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL4 + 50 - (HCOL1 + 300), "", ifontName, ifontSizesm, "LB", "");//other charges 
            AddXYLabel(HCOL4 + 50, Row, ROW_HT, HCOL4 + 100 - (HCOL4 + 50), "", ifontName, ifontSizesm, "B", "R");
            AddXYLabel(HCOL4 + 100, Row, ROW_HT, HCOL5 + 30 - (HCOL4 + 100), "", ifontName, ifontSizesm, "B", "R");
            AddXYLabel(HCOL5 + 30, Row, ROW_HT, HCOL5 + 140 - (HCOL5 + 30), "", ifontName, ifontSizesm, "B", "");
            AddXYLabel(HCOL5 + 140, Row, ROW_HT, HCOL5 + 190 - (HCOL5 + 140), "", ifontName, ifontSizesm, "B", "R");
            AddXYLabel(HCOL5 + 190, Row, ROW_HT, HCOL8 - (HCOL5 + 190), "", ifontName, ifontSizesm, "BR", "R");

            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH1[4].ToString(), "9", "L", "", R1.ToString(), COL03.ToString(), "0", "12");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", CH2[4].ToString(), "9", "R", "", R1.ToString(), COL04.ToString(), WD3.ToString(), "12");

            FindTotal();

            Row += ROW_HT; R1++;

            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "9", "L", "", R1.ToString(), COL01.ToString(), "0", "12");
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 50 - HCOL1, "", ifontName, ifontSizesm, "L", "");
            AddXYLabel(HCOL1 + 50, Row, ROW_HTsm, HCOL1 + 250 - (HCOL1 + 50), "Total other Charges Due Agent", ifontName, ifontSizesm, "lBr", "C");
            PrintEasy("Total other Charges Due Agent", 12, ifontSizesm, "B", "C", 5, 6, 36);
            DrawVLine(HCOL1 + 300, Row, ROW_HT + ROW_HTsm);
            DrawVLine(HCOL8, Row, ROW_HT + ROW_HTsm);
            s1 = "";
            Str = "Shipper certifies that the particulars on the face hereof are correct and that insofar as any part of the consignment";
            s1 += Str;
            AddXYLabel(HCOL1 + 300, Row, 10, HCOL8 - (HCOL1 + 300), Str, ifontName, 6, "", "");
            Str = "contains dangerous goods such part is properly described by name and is in proper condition for carriage by air";
            s1 += Str;
            AddXYLabel(HCOL1 + 300, Row + 9, 10, HCOL8 - (HCOL1 + 300), Str, ifontName, 6, "", "");
            Str = "according to the applicable Dangerous Goods Regulations.";
            s1 += Str;
            AddXYLabel(HCOL1 + 300, Row + 18, 10, HCOL8 - (HCOL1 + 300), Str, ifontName, 6, "", "");
            WD3 = COLTOT - COL03 - 1;
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", s1, "8", "LTR", "TW", R1.ToString(), COL03.ToString(), WD3.ToString(), "12", "2");


            Row += ROW_HTsm; R1++;
            string OthppAsArrngd = "", OthccAsArrngd = "";
            if (mRow.bl_oc_status.ToString().Trim() == "P")
            {
                if ((mRow.bl_asarranged_shipper.ToString().Trim() == "Y") || (mRow.bl_asarranged_consignee.ToString().Trim() == "Y"))
                    OthppAsArrngd = "AS AGREED";
                else
                    OthppAsArrngd = "";
            }
            else if (mRow.bl_oc_status.ToString().Trim() == "C")
            {
                if ((mRow.bl_asarranged_shipper.ToString().Trim() == "Y") || (mRow.bl_asarranged_consignee.ToString().Trim() == "Y"))
                    OthccAsArrngd = "AS AGREED";
                else
                    OthccAsArrngd = "";
            }
            if (OthppAsArrngd.Trim().Length > 0)
            {
                Str = (BL_TYPE == "MAWB") ? OthppAsArrngd : "";
            }
            else
                Str = (Agent_Tot_PP > 0) ? Agent_Tot_PP.ToString() : "";
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, Str, ifontName, ifontSize, "LB", "C", R1, COL01, WD1, 16, 0);

            if (OthccAsArrngd.Trim().Length > 0)
            {
                Str = (BL_TYPE == "MAWB") ? OthccAsArrngd : "";
            }
            else
                Str = (Agent_Tot_CC > 0) ? Agent_Tot_CC.ToString() : "";
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), Str, ifontName, ifontSize, "LB", "C", R1, COL02, WD2, 16, 0);



            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 50 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, COL01, 0, 16, 0);
            AddXYLabel(HCOL1 + 50, Row, ROW_HTsm, HCOL1 + 250 - (HCOL1 + 50), "Total other Charges Due Carrier", ifontName, ifontSizesm, "Blr", "C");
            AddXYLabel(HCOL1 + 300, Row, ROW_HTsm, HCOL8 - (HCOL1 + 300), "", ifontName, ifontSizesm, "LR", "");



            PrintEasy("Total other Charges Due Carrier", ROW_HTsm, ifontSizesm, "B", "C", 5, 6, 36);


            Row += ROW_HTsm; R1++;
            if (OthppAsArrngd.Trim().Length > 0)
            {
                Str = (BL_TYPE == "MAWB") ? OthppAsArrngd : "";
            }
            else
                Str = (Carrier_Tot_PP > 0) ? Carrier_Tot_PP.ToString() : "";
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, Str, ifontName, ifontSize, "L", "C", R1, COL01, WD1, 16, 0);

            if (OthccAsArrngd.Trim().Length > 0)
            {
                Str = (BL_TYPE == "MAWB") ? OthccAsArrngd : "";
            }
            else
                Str = (Carrier_Tot_CC > 0) ? Carrier_Tot_CC.ToString() : "";

            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), Str, ifontName, ifontSize, "L", "C", R1, COL02, WD2, 16, 0);
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL8 - (HCOL1 + 300), mRow.bl_by1_agent.ToString(), ifontName, ifontSize, "LR", "", R1, COL03, WD3, 16, 0);

            Row += ROW_HT; R1++;
            SetFillRectangle(HCOL1, Row, ROW_HT * 2, 300);

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, "", ifontName, ifontSize, "LT", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), "", ifontName, ifontSize, "LT", "", R1, COL02, WD2, 16, 0);
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL8 - (HCOL1 + 300), mRow.bl_by2_agent.ToString(), ifontName, ifontSize, "LR", "", R1, COL03, WD3, 16, 0);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, "", ifontName, ifontSize, "LB", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), "", ifontName, ifontSize, "LB", "", R1, COL02, WD2, 16, 0);

            //DrawDotLine(HCOL1 + 300, Row, HCOL8 - (HCOL1 + 300));
            DrawDashLine(HCOL1 + 300, Row, HCOL8 - (HCOL1 + 300));
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL8 - (HCOL1 + 300), "Signature of Shipper or His Agent", ifontName, ifontSizesm, "LR", "C", R1, COL03, WD3, 16, 0);

            R1++;

            DrawXLHBorder(R1, COL01, COLTOT - 1, 1, "B");


            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 20 - (HCOL1), "", ifontName, ifontSizesm, "L", "", R1, COL01, 0, 16, 0);
            AddXYLabel(HCOL1 + 20, Row, ROW_HTsm, HCOL1 + 130 - (HCOL1 + 20), "Total prepaid", ifontName, ifontSizesm, "lBr", "C");

            AddXYLabel(HCOL1 + 130, Row, ROW_HTsm, HCOL1 + 150 - (HCOL1 + 130), "", ifontName, ifontSizesm, "R", "");
            AddXYLabel(HCOL1 + 170, Row, ROW_HTsm, HCOL1 + 280 - (HCOL1 + 170), "Total Collect", ifontName, ifontSizesm, "lBr", "C");

            PrintEasy("Total Prepaid", ROW_HTsm, ifontSizesm, "B", "C", 3, 4, 17);

            AddXYLabel(HCOL1 + 280, Row, ROW_HTsm, HCOL1 + 300 - (HCOL1 + 280), "", ifontName, ifontSizesm, "", "");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "8", "L", "", R1.ToString(), COL02.ToString(), "0", "16", "0");
            PrintEasy("Total Collect", ROW_HTsm, ifontSizesm, "B", "C", 22, 23, 37);
            AddXYLabel(HCOL1 + 300, Row, ROW_HTsm, HCOL8 - (HCOL1 + 300), mRow.bl_by1_carrier.ToString(), ifontName, ifontSize, "LRT", "", R1, COL03, WD3, 16, 0);


            Row += ROW_HTsm; R1++;
            if (OthppAsArrngd.Trim().Length > 0 || frtppAsArrngd.Trim().Length > 0)
            {
                Str = "AS AGREED";
            }
            else
                Str = (Agent_Tot_PP + Carrier_Tot_PP + frtppAmt > 0) ? (Agent_Tot_PP + Carrier_Tot_PP + frtppAmt).ToString() : "";
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, Str, ifontName, ifontSize, "LB", "C", R1, COL01, WD1, 16, 0);

            if (OthccAsArrngd.Trim().Length > 0 || frtccAsArrngd.Trim().Length > 0)
            {
                Str = "AS AGREED";
            }
            else
                Str = (Agent_Tot_CC + Carrier_Tot_CC + frtccAmt > 0) ? (Agent_Tot_CC + Carrier_Tot_CC + frtccAmt).ToString() : "";
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), Str, ifontName, ifontSize, "LB", "C", R1, COL02, WD2, 16, 0);
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL8 - (HCOL1 + 300), mRow.bl_by2_carrier.ToString(), ifontName, ifontSize, "LR", "", R1, COL03, WD3, 16, 0);


            Row += ROW_HT; R1++;

            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 10 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, COL01, 0, 16, 0);
            AddXYLabel(HCOL1 + 10, Row, ROW_HTsm, HCOL1 + 160 - (HCOL1 + 30), "Currency Conversion Rates", ifontName, ifontSizesm, "lBr", "C");
            PrintEasy("Currency Conversion Rates", ROW_HTsm, ifontSizesm, "B", "C", 3, 4, 18);


            AddXYLabel(HCOL1 + 150, Row, ROW_HTsm, HCOL1 + 160 - (HCOL1 + 150), "", ifontName, ifontSizesm, "L", "", R1, COL02, 0, 16, 0);

            AddXYLabel(HCOL1 + 160, Row, ROW_HTsm, HCOL1 + 300 - (HCOL1 + 160 + 10), "CC Charges in Dest. Currency", ifontName, ifontSizesm, "lBr", "C");

            PrintEasy("CC Charges in Dest. Currency", ROW_HTsm, ifontSizesm, "B", "C", 21, 22, 38);

            AddXYLabel(HCOL1 + 300, Row, ROW_HTsm, HCOL6 - (HCOL1 + 300), "", ifontName, ifontSize, "L", "C");

            //Str = (mRow.bl_issued_date_print.ToString().Trim() != "") ? Convert.ToDateTime(DR[AIRBL.COL_BL_ISSUED_DATE]).ToShortDateString() : "";
            Str = mRow.bl_issued_date_print.ToString();
            s1 = "   " + Str;
            Str += "   " + mRow.bl_issued_place.ToString();
            s1 += "   " + mRow.bl_issued_place.ToString();
            AddXYLabel(HCOL1 + 360, Row - 3, ROW_HTsm, HCOL6 - (HCOL1 + 360), Str, ifontName, ifontSize, "", "L");
            AddXYLabel(HCOL6, Row, ROW_HTsm, HCOL8 - (HCOL6), "", ifontName, ifontSize, "R", "L");
            AddXYLabel(HCOL6, Row - 3, ROW_HTsm, HCOL8 - (HCOL6), mRow.bl_issued_by.ToString(), ifontName, ifontSize, "", "L");
            s1 += "    " + mRow.bl_issued_by.ToString();
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", s1, "8", "L", "", R1.ToString(), COL03.ToString(), "0", "16", "0");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "8", "L", "", R1.ToString(), COLTOT.ToString(), "0", "12", "0");


            Row += ROW_HTsm; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, "", ifontName, ifontSize, "LB", "", R1, COL01, WD1, 16, 0);
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), "", ifontName, ifontSize, "LB", "", R1, COL02, WD2, 16, 0);
            Str = "Executed on      (Date)                  at            (Place)                        Signature of Issuing Carrier or its Agent";
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL8 - (HCOL1 + 300), Str, ifontName, ifontSizesm, "LBR", "", R1, COL03, WD3, 16, 0);
            DrawDashLine(HCOL1 + 300, Row, HCOL8 - (HCOL1 + 300));

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HTsm, HCOL1 + 150 - HCOL1, "", ifontName, ifontSize, "L", "", R1, COL01, 0, 16, 0);
            AddXYLabel(HCOL1 + 30, Row + 5, ROW_HTsm, HCOL1 + 150 - HCOL1, "For Carriers Use only", ifontName, ifontSizesm, "", "", R1, COL01 + 1, 0, 16, 0);
            AddXYLabel(HCOL1 + 150, Row, ROW_HTsm, HCOL1 + 170 - (HCOL1 + 150), "", ifontName, ifontSizesm, "L", "", R1, COL02, 0, 16, 0);
            AddXYLabel(HCOL1 + 170, Row, ROW_HTsm, HCOL1 + 300 - (HCOL1 + 170 + 20), "Charges at Destination", ifontName, ifontSizesm, "lBr", "C");
            PrintEasy("Charges at Destination", ROW_HTsm, ifontSizesm, "B", "C", 21, 22, 38);
            AddXYLabel(HCOL1 + 300, Row, ROW_HTsm, HCOL1 + 320 - (HCOL1 + 300), "", ifontName, ifontSizesm, "L", "", R1, COL03, 0, 16, 0);
            AddXYLabel(HCOL1 + 320, Row, ROW_HTsm, HCOL1 + 440 - (HCOL1 + 320), "Total Collect charges", ifontName, ifontSizesm, "lBr", "C");
            PrintEasy("Total Collect Charges", ROW_HTsm, ifontSizesm, "B", "C", 42, 43, 58);
            AddXYLabel(HCOL1 + 460, Row, ROW_HTsm, HCOL8 - (HCOL1 + 440), "", ifontName, ifontSizesm, "L", "", R1, 61, 0, 16, 0);
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL8 - HCOL6, "", ifontName, ifontSize + 3, "R", "B");
            AddXYLabel(HCOL5, Row + 5, ROW_HT, HCOL8 - HCOL5 - 20, (BL_TYPE == "MAWB") ? mRow.bl_mbl_no.ToString() : "HAWB No. " + mRow.hbl_bl_no.ToString(), ifontName, ifontSize + 4, "", "RB", R1, 71, 0, 16, 0);

            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "8", "L", "", R1.ToString(), COLTOT.ToString(), "0", "12", "0");

            Row += ROW_HTsm; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL1 + 150 - HCOL1, "", ifontName, ifontSize, "LB", "", R1, COL01, 0, 16, 0);
            AddXYLabel(HCOL1 + 30, Row - 1, ROW_HT, HCOL1 + 150 - HCOL1, "at Destination", ifontName, ifontSizesm, "", "", R1, COL01 + 1, 0, 16, 0);
            AddXYLabel(HCOL1 + 150, Row, ROW_HT, HCOL1 + 300 - (HCOL1 + 150), "", ifontName, ifontSize, "LB", "", R1, COL02, 0, 16, 0);
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL1 + 460 - (HCOL1 + 300), "", ifontName, ifontSize, "LB", "", R1, COL03, 0, 16, 0);
            AddXYLabel(HCOL1 + 460, Row, ROW_HT, HCOL8 - (HCOL1 + 460), "", ifontName, ifontSize, "LBR", "", R1, 61, 0, 16, 0);
            Str = "";
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL8 - HCOL5, Str, ifontName, ifontSize, "BR", "", R1, 71, 0, 16, 0);
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", "8", "L", "", R1.ToString(), COLTOT.ToString(), "0", "12", "0");
            R1++;
            DrawXLHBorder(R1, COL01, COLTOT - 1, 1, "T");
            
            Str = FooterNote;
            Row += ROW_HT + 5; R1++;
            AddXYLabel(HCOL1 + 300, Row, ROW_HT, HCOL8 - HCOL4, Str, ifontName, ifontSize, "", "");//, R1, 71, 0, 16, 0

        }

        private void PrintEasy(string str, float Ht, int fontsize, string sBorder, string sStyle, int sCol, int Col, int eCol)
        {
            //addList("EXTRA_XL_TEXT", "0", "0", "0", "0", str, fontsize.ToString(), "LRB", "C", R1.ToString(), COL05.ToString(), WD3.ToString(), "16", "0");
            int wd = eCol - Col - 1;
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", fontsize.ToString(), "l", "", R1.ToString(), sCol.ToString(), "0", Ht.ToString(), "0");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", str, fontsize.ToString(), sBorder, sStyle, R1.ToString(), Col.ToString(), wd.ToString(), Ht.ToString(), "0");
            addList("EXTRA_XL_TEXT", "0", "0", "0", "0", "", fontsize.ToString(), "r", "", R1.ToString(), eCol.ToString(), "0", Ht.ToString(), "0");
        }
        private void FindTotal()
        {
            Agent_Tot_PP = 0; Agent_Tot_CC = 0; Carrier_Tot_PP = 0; Carrier_Tot_CC = 0;
            if (mRow.bl_oc_status.ToString().Trim() == "P")
            {
                GetOthData(mRow.bl_charges1_agent);
                Agent_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges2_agent);
                Agent_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges3_agent);
                Agent_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges4_agent);
                Agent_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges5_agent);
                Agent_Tot_PP += Lib.Conv2Decimal(OthTotal);
                 

                GetOthData(mRow.bl_charges1_carrier);
                Carrier_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges2_carrier);
                Carrier_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges3_carrier);
                Carrier_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges4_carrier);
                Carrier_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges5_carrier);
                Carrier_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges6_carrier);
                Carrier_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges7_carrier);
                Carrier_Tot_PP += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges8_carrier);
                Carrier_Tot_PP += Lib.Conv2Decimal(OthTotal);
            }
            if (mRow.bl_oc_status.ToString().Trim() == "C")
            {
                GetOthData(mRow.bl_charges1_agent);
                Agent_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges2_agent);
                Agent_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges3_agent);
                Agent_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges4_agent);
                Agent_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges5_agent);
                Agent_Tot_CC += Lib.Conv2Decimal(OthTotal);

                GetOthData(mRow.bl_charges1_carrier);
                Carrier_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges2_carrier);
                Carrier_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges3_carrier);
                Carrier_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges4_carrier);
                Carrier_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges5_carrier);
                Carrier_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges6_carrier);
                Carrier_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges7_carrier);
                Carrier_Tot_CC += Lib.Conv2Decimal(OthTotal);
                GetOthData(mRow.bl_charges8_carrier);
                Carrier_Tot_CC += Lib.Conv2Decimal(OthTotal);
            }
        }

        private void WriteDescription()
        {
            string str = "";
            Row = DescStartRow;

            int iCtr = 0;

            // WriteHead(); using for xl

            string sDesc = "";


            string s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14;
            s1 = ""; s2 = ""; s3 = ""; s4 = ""; s5 = ""; s6 = ""; s7 = ""; s8 = ""; s9 = ""; s10 = ""; s11 = ""; s12 = ""; s13 = ""; s14 = "";
            s14 = sDesc;

            R1++;
            str = (Lib.Conv2Integer(mRow.bl_pcs.ToString()) > 0) ? mRow.bl_pcs.ToString() : "";
            AddXYLabel(dCol1, Row, ROW_HT, dCol2 - dCol1 - 5, str, ifontName, ifontSize, "", "R");
            s1 = mRow.bl_pcs.ToString();

            str = (Lib.Conv2Decimal(mRow.bl_grwt.ToString()) > 0) ? mRow.bl_grwt.ToString() : "";
            AddXYLabel(dCol2, Row, ROW_HT, dCol3 - dCol2 - 5, str, ifontName, ifontSize, "", "R");
            s2 = str;

            str = (mRow.bl_wt_unit.ToString().Trim().Length > 0) ? mRow.bl_wt_unit.ToString().Trim() : "";//.Substring(0, 1)
            AddXYLabel(dCol3, Row, ROW_HT, dCol4 - dCol3, str, ifontName, ifontSize, "", "L");//9
            s3 = str;
            s4 = "";
            str = (mRow.bl_class.ToString().Trim().Length > 0) ? mRow.bl_class.ToString().Trim() : "";//.Substring(0, 1)
            AddXYLabel(dCol5 - 1, Row, ROW_HT, dCol4 - dCol3, str, ifontName, ifontSize, "", "");
            s5 = str;
            AddXYLabel(dCol6, Row, ROW_HT, dCol7 - dCol6, mRow.bl_comm.ToString(), ifontName, ifontSize, "", "L");
            s6 = mRow.bl_comm.ToString();
            s7 = "";

            str = (Lib.Conv2Decimal(mRow.bl_chwt.ToString()) > 0) ? mRow.bl_chwt.ToString() : "";
            AddXYLabel(dCol8, Row, ROW_HT, dCol9 - dCol8 - 5, str, ifontName, ifontSize, "", "R");
            s8 = str;

            str = (mRow.bl_wt_unit.ToString().Trim().Length > 0) ? mRow.bl_wt_unit.ToString().Trim() : "";//.Substring(0, 1)
            AddXYLabel(dCol9, Row, ROW_HT, dCol10 - dCol9, str, ifontName, ifontSize, "", "");
            s9 = str;



            if ((mRow.bl_asarranged_shipper.ToString().Trim() == "Y") || (mRow.bl_asarranged_consignee.ToString().Trim() == "Y"))
            {
                str = "AS AGREED";
                AddXYLabel(dCol10, Row, ROW_HT, dCol11 - dCol10 - 5, "", ifontName, ifontSize, "", "R");
                AddXYLabel(dCol12, Row, ROW_HT, dCol13 - dCol12 + 10, str, ifontName, ifontSize, "", "L");
                s10 = "";
                s11 = "";
                s12 = str;
            }
            else
            {
                AddXYLabel(dCol10, Row, ROW_HT, dCol11 - dCol10 - 5, (Lib.Conv2Decimal(mRow.bl_rate.ToString()) > 0) ? mRow.bl_rate.ToString() : "", ifontName, ifontSize, "", "R");
                AddXYLabel(dCol12, Row, ROW_HT, dCol13 - dCol12 - 5, (Lib.Conv2Decimal(mRow.bl_total.ToString()) > 0) ? mRow.bl_total.ToString() : "", ifontName, ifontSize, "", "R");
                s10 = (Lib.Conv2Decimal(mRow.bl_rate.ToString()) > 0) ? mRow.bl_rate.ToString() : "";
                s11 = "";
                s12 = (Lib.Conv2Decimal(mRow.bl_total.ToString()) > 0) ? mRow.bl_total.ToString() : "";
            }
            iCtr++;
            // WriteXlDesc(iCtr, s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14);
            R1++;


            Row += ROW_HT;
            s1 = ""; s2 = ""; s3 = ""; s4 = ""; s5 = ""; s6 = ""; s7 = ""; s8 = ""; s9 = ""; s10 = ""; s11 = ""; s12 = ""; s13 = ""; s14 = "";

            // str = (MblRec.mbld_lbs > 0) ? MblRec.mbld_lbs.ToString() : "";
            str = "";
            AddXYLabel(dCol2, Row, ROW_HT, dCol3 - dCol2 - 5, str, ifontName, ifontSize, "", "R");
            s2 = str;
            //  str = (MblRec.mbld_lbs > 0) ? "L" : "";
            str = "";
            AddXYLabel(dCol3, Row, ROW_HT, dCol4 - dCol3, str, ifontName, ifontSize, "", "");
            s3 = str;
            iCtr++;
            // WriteXlDesc(iCtr, s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14);
            R1++;

            while (iCtr < 15)
            {
                iCtr++;
                s1 = ""; s2 = ""; s3 = ""; s4 = ""; s5 = ""; s6 = ""; s7 = ""; s8 = ""; s9 = ""; s10 = ""; s11 = ""; s12 = ""; s13 = ""; s14 = "";
                //  WriteXlDesc(iCtr, s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14);
                R1++;
            }

            Row += ROW_HT * 15;

            s1 = ""; s2 = ""; s3 = ""; s4 = ""; s5 = ""; s6 = ""; s7 = ""; s8 = ""; s9 = ""; s10 = ""; s11 = ""; s12 = ""; s13 = ""; s14 = "";

            DrawHLine(dCol1, Row, dCol4 - dCol1);
            str = (Lib.Conv2Integer(mRow.bl_pcs.ToString()) > 0) ? mRow.bl_pcs.ToString() : "";
            AddXYLabel(dCol1, Row, ROW_HT, dCol2 - dCol1 - 5, str, ifontName, ifontSize, "", "R");
            s1 = str;
            str = (Lib.Conv2Decimal(mRow.bl_grwt.ToString()) > 0) ? mRow.bl_grwt.ToString() : "";
            AddXYLabel(dCol2, Row, ROW_HT, dCol3 - dCol2 - 5, str, ifontName, ifontSize, "", "R");
            s2 = str;
            str = (mRow.bl_wt_unit.ToString().Trim().Length > 0) ? mRow.bl_wt_unit.ToString().Trim() : "";//.Substring(0, 1)
            AddXYLabel(dCol3, Row, ROW_HT, dCol4 - dCol3, str, ifontName, ifontSize, "", "");
            s3 = str;

            DrawHLine(dCol8, Row, dCol9 - dCol8);
            str = (Lib.Conv2Decimal(mRow.bl_chwt.ToString()) > 0) ? mRow.bl_chwt.ToString() : "";
            AddXYLabel(dCol8, Row, ROW_HT, dCol9 - dCol8 - 5, str, ifontName, ifontSize, "", "R");
            s8 = str;
            str = (mRow.bl_wt_unit.ToString().Trim().Length > 0) ? mRow.bl_wt_unit.ToString().Trim() : "";//.Substring(0, 1)
            AddXYLabel(dCol9, Row, ROW_HT, dCol10 - dCol9, str, ifontName, ifontSize, "", "");
            s9 = str;
            DrawHLine(dCol10, Row, dCol11 - dCol10);
            DrawHLine(dCol12, Row, dCol13 - dCol12);
            // if ((InvokeType == "SHPR" && DR[AIRBL.COL_BL_ASARRANGED_SHIPPER].ToString().Trim() == "Y") || (InvokeType == "CNEE" && DR[AIRBL.COL_BL_ASARRANGED_CONSIGNEE].ToString().Trim() == "Y"))
            if ((mRow.bl_asarranged_shipper.ToString().Trim() == "Y") || (mRow.bl_asarranged_consignee.ToString().Trim() == "Y"))
            {
                str = "AS AGREED";
                AddXYLabel(dCol10, Row, ROW_HT, dCol11 - dCol10 - 5, "", ifontName, ifontSize, "", "R");
                AddXYLabel(dCol12, Row, ROW_HT, dCol13 - dCol12 + 10, str, ifontName, ifontSize, "", "L");
                s10 = "";
                s11 = "";
                s12 = str;
            }
            else
            {
                AddXYLabel(dCol10, Row, ROW_HT, dCol11 - dCol10 - 5, (Lib.Conv2Decimal(mRow.bl_rate.ToString()) > 0) ? mRow.bl_rate.ToString() : "", ifontName, ifontSize, "", "R");
                AddXYLabel(dCol12, Row, ROW_HT, dCol13 - dCol12 - 5, (Lib.Conv2Decimal(mRow.bl_total.ToString()) > 0) ? mRow.bl_total.ToString() : "", ifontName, ifontSize, "", "R");
                s10 = (Lib.Conv2Decimal(mRow.bl_rate.ToString()) > 0) ? mRow.bl_rate.ToString() : "";
                s11 = "";
                s12 = (Lib.Conv2Decimal(mRow.bl_total.ToString()) > 0) ? mRow.bl_total.ToString() : "";
            }

            DrawXLHBorder(R1, COL01, COL03, 1, "T");
            DrawXLHBorder(R1, COL08, COL11 - 1, 1, "T");
            DrawXLHBorder(R1, COL12, COL13 - 1, 1, "T");
            R1++;
            iCtr++;
            //WriteXlDesc(iCtr, s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14);
            R1++;
            s1 = ""; s2 = ""; s3 = ""; s4 = ""; s5 = ""; s6 = ""; s7 = ""; s8 = ""; s9 = ""; s10 = ""; s11 = ""; s12 = ""; s13 = ""; s14 = "";
            iCtr++;
            // WriteXlDesc(iCtr, s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14);
            R1++;
            DrawXLHBorder(R1, COL01, COLTOT - 1, 1, "T");

            Row = DescStartRow;
            //int dCtr = 0;
            //foreach (DataRow dr in Dt_DESC.Rows)
            //{
            //    dCtr = Common.Convert2Integer(dr["bl_desc_ctr"].ToString());
            //    Row = DescStartRow + (float)(dCtr - 1) * ROW_HT;
            //    AddXYLabel(dCol1, Row + ROW_HT, ROW_HT, dCol14 - dCol1, dr["bl_marks"].ToString(), ifontName, ifontSize - 1, "", "L");
            //    AddXYLabel(dCol14, Row + ROW_HT, ROW_HT, HCOL8 - dCol14 + 40, dr["bl_desc"].ToString(), ifontName, ifontSize - 1, "", "L");
            //}
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark1 != null ? mRow.bl_mark1.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc1 != null ? mRow.bl_desc1.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark2 != null ? mRow.bl_mark2.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc2 != null ? mRow.bl_desc2.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark3 != null ? mRow.bl_mark3.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc3 != null ? mRow.bl_desc3.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark4 != null ? mRow.bl_mark4.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc4 != null ? mRow.bl_desc4.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark5 != null ? mRow.bl_mark5.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc5 != null ? mRow.bl_desc5.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark6 != null ? mRow.bl_mark6.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc6 != null ? mRow.bl_desc6.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark7 != null ? mRow.bl_mark7.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc7 != null ? mRow.bl_desc7.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark8 != null ? mRow.bl_mark8.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc8 != null ? mRow.bl_desc8.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark9 != null ? mRow.bl_mark9.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc9 != null ? mRow.bl_desc9.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark10 != null ? mRow.bl_mark10.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc10 != null ? mRow.bl_desc10.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark11 != null ? mRow.bl_mark11.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc11 != null ? mRow.bl_desc11.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark12 != null ? mRow.bl_mark12.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc12 != null ? mRow.bl_desc12.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark13 != null ? mRow.bl_mark13.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc13 != null ? mRow.bl_desc13.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark14 != null ? mRow.bl_mark14.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc14 != null ? mRow.bl_desc14.ToString() : "", ifontName, ifontSize - 1, "", "L");
            Row += ROW_HT;
            AddXYLabel(dCol1, Row, ROW_HT, dCol14 - dCol1, mRow.bl_mark15 != null ? mRow.bl_mark15.ToString() : "", ifontName, ifontSize - 1, "", "L");
            AddXYLabel(dCol14, Row, ROW_HT, HCOL8 - dCol14 + 40, mRow.bl_desc15 != null ? mRow.bl_desc15.ToString() : "", ifontName, ifontSize - 1, "", "L");

        }


        private void WriteOtherCharges()
        {
            // OthChrgStartRow += 3;
            Row = OthChrgStartRow + 5;
            float OthChrgHCOL = HCOL1 + 300;
            int PrntRowCount = 0;
            string Chrgs = "";

            int c1 = 0;

            if (GetOthData(mRow.bl_charges1_carrier))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Carrier
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH1[c1] = Chrgs; c1++;
                }
            if (GetOthData(mRow.bl_charges2_carrier))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Carrier
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH1[c1] = Chrgs; c1++;
                }
            if (GetOthData(mRow.bl_charges3_carrier))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Carrier
                    Row += ROW_HTsm;
                    PrntRowCount++;
                    CH1[c1] = Chrgs; c1++;
                }

            //Row = OthChrgStartRow + 5;
            //OthChrgHCOL += 110;

            if (GetOthData(mRow.bl_charges4_carrier))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Carrier
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH1[c1] = Chrgs; c1++;
                }
            if (GetOthData(mRow.bl_charges5_carrier))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Carrier
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH1[c1] = Chrgs; c1++;
                }
            if (GetOthData(mRow.bl_charges6_carrier))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Carrier
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH1[c1] = Chrgs; c1++;
                }

            Row = OthChrgStartRow + 5;
            OthChrgHCOL += 130;

            if (GetOthData(mRow.bl_charges7_carrier))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Carrier
                    Row += ROW_HTsm;
                    PrntRowCount++;
                    CH1[c1] = Chrgs; c1++;
                }
            if (GetOthData(mRow.bl_charges8_carrier))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Carrier
                    Row += ROW_HTsm;
                    PrntRowCount++;
                    CH1[c1] = Chrgs; c1++;
                }



            Row = OthChrgStartRow + 5;
            OthChrgHCOL += 130;
            c1 = 0;

            if (GetOthData(mRow.bl_charges1_agent))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Agent
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH2[c1] = Chrgs; c1++;
                }
            if (GetOthData(mRow.bl_charges2_agent))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges Ahgent        
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH2[c1] = Chrgs; c1++;
                }
            if (GetOthData(mRow.bl_charges3_agent))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges agent
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH2[c1] = Chrgs; c1++;
                }
            //Row = OthChrgStartRow + 5;
            //OthChrgHCOL += 110;

            if (GetOthData(mRow.bl_charges4_agent))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges agent
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH2[c1] = Chrgs; c1++;
                }

            if (GetOthData(mRow.bl_charges5_agent))
                if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                  (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                {
                    Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                    AddXYLabel(OthChrgHCOL, Row, ROW_HTsm, 200, Chrgs, ifontName, ifontSizesm + 2, "", "");//other charges agent
                    Row += ROW_HTsm ;
                    PrntRowCount++;
                    CH2[c1] = Chrgs; c1++;
                }

        }

        private string GetFormatOtherChrgs(string Charges, string Rates, string TotChrgs)
        {
            string Str = "";
            if (Charges.Trim().Length <= 0)
                return Str;
            Str = Charges;
            Str += " : ";
            //if (Lib.Conv2Decimal(Rates) > 0)
            //{
            //    Str += mRow.bl_currency.ToString();
            //    Str += Rates;
            //    Str += "/" +mRow.bl_wt_unit.ToString();
            //    Str += " = ";
            //}
            //Str += mRow.bl_currency.ToString();
            Str += (Lib.Conv2Decimal(TotChrgs) > 0) ? TotChrgs : "";
            return Str;
        }

        private void WriteBackSide()
        {
            string Flname = "";
            string sline = null;

            HCOL1 = 46;
            HCOL2 = HCOL1 + 18;
            HCOL3 = HCOL2 + 40;
            HCOL4 = HCOL3 + 283;
            HCOL5 = HCOL4 + 26;
            HCOL6 = HCOL5 + 18;
            HCOL7 = HCOL6 + 40;
            HCOL8 = HCOL7 + 283;

            Row = 50;

            ROW_HT = 12; ifontName = "Arial"; ifontSize = 7; int bfontSize = 10;

            Flname = RootPath + "\\MAWB-BACKSIDE.TXT";
            if (!File.Exists(Flname))
                throw new Exception("BACK SIDE FILE NOT FOUND");

            StreamReader reader = new StreamReader(Flname);
            string[] ColLines = null;
            Boolean IsColumn1 = true;
            float Col1_StartRow = 0;
            while ((sline = reader.ReadLine()) != null)
            {
                sline = sline.Trim();

                if (sline == "{COLUMN 1}")
                    Col1_StartRow = Row;
                else if (sline == "{COLUMN 2}")
                {
                    Row = Col1_StartRow;
                    IsColumn1 = false;
                }
                else if (sline == "{HLINE}")
                    DrawHLine(HCOL1, Row, HCOL8 - HCOL1);
                else
                {
                    ColLines = sline.Split(sColSplit);
                    if (IsColumn1)
                    {

                        if (ColLines.Length == 1)
                            AddXYLabel(HCOL1, Row, ROW_HT, HCOL8 - HCOL1, GetFormatLine(sline), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        else if (ColLines.Length == 2)
                        {
                            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, GetFormatLine(ColLines[0]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, GetFormatLine(ColLines[1]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        }
                        else if (ColLines.Length == 3)
                        {
                            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, GetFormatLine(ColLines[0]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, GetFormatLine(ColLines[1]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, GetFormatLine(ColLines[2]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        }
                    }
                    else
                    {
                        if (ColLines.Length == 1)
                            AddXYLabel(HCOL5, Row, ROW_HT, HCOL8 - HCOL5, GetFormatLine(sline), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        else if (ColLines.Length == 2)
                        {
                            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, GetFormatLine(ColLines[0]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL6, Row, ROW_HT, HCOL8 - HCOL6, GetFormatLine(ColLines[1]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        }
                        else if (ColLines.Length == 3)
                        {
                            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, GetFormatLine(ColLines[0]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, GetFormatLine(ColLines[1]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL7, Row, ROW_HT, HCOL8 - HCOL7, GetFormatLine(ColLines[2]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        }

                    }
                    Row += ROW_HT;
                }
            }
            if (reader != null)
                reader.Close();

            // AddXYLabel(this.Page_Width, this.Page_Height, 0, 0, "COPY", ifontName, 120, "", "W"); //water Mark
        }

        private string GetFormatLine(string sLine)
        {
            string[] sData = null;
            BsideStyle = "";
            sLine = sLine.Trim();
            sData = sLine.Split(sStyleSplit);
            if (sData.Length > 1)
            {
                BsideStyle = sData[0].Trim();
                sLine = sData[1].Trim();
            }

            return sLine;
        }
    

        private Boolean GetOthData(object obj)
        {
            Boolean bOk = false;
            OthCharges = "";
            OthWeight = "";
            OthRate = "";
            OthTotal = "";
            OthPrintS = "";
            OthPrintC = "";
            if (obj != null)
                if (obj.ToString().Contains(","))
                {
                    bOk = true;
                    string[] sdata = obj.ToString().Split(',');
                    OthCharges = sdata[0];
                    OthWeight = sdata[1];
                    OthRate = sdata[2];
                    OthTotal = sdata[3];
                    OthPrintS = sdata[4];
                    OthPrintC = sdata[5];
                }
            return bOk;
        }
    }
}
