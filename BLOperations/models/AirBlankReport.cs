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
    public class AirBlankReport : BaseReport
    {

        public Bl mRow = null;
        public string RootPath = "";
        public DataTable Dt_COLPOS = new DataTable();
        public string FooterNote = "";
        public string InvokeType = "";

        private int R1 = 0;
        private const int XL_COLA = 1;
        private const int XL_COLB = 2;
        private const int XL_COLC = 3;
        private const int XL_COLD = 4;
        private const int XL_COLE = 5;
        private const int XL_COLF = 6;
        private const int XL_COLG = 7;
        private const int XL_COLH = 8;
        private const int XL_COLI = 9;
        private const int XL_COLJ = 10;
        private const int XL_COLK = 11;
        private const int XL_COL_TOT = 10;
        private float Xtolrnce = 5;
        private string sError = "";
        private int x1 = 0, y1 = 0, h1 = 0, w1 = 0, fsize = 0;
        private string fname = "", sStyle = "";

        private string OthCharges = "";
        private string OthWeight = "";
        private string OthRate = "";
        private string OthTotal = "";
        private string OthPrintS = "";
        private string OthPrintC = "";
        private decimal Agent_Tot_PP = 0, Agent_Tot_CC = 0;
        private decimal Carrier_Tot_PP = 0, Carrier_Tot_CC = 0;

        string[] CH1 = { "", "", "", "", "", "" };
        string[] CH2 = { "", "", "", "", "", "" };


        public AirBlankReport()
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
                if (sError.ToString().Trim() != "")
                    throw new Exception(sError);

                if (FooterNote.ToUpper().Contains("SHIPPER"))
                    InvokeType = "SHPR";
                else if (FooterNote.ToUpper().Contains("CONSIGNEE"))
                    InvokeType = "CNEE";
                else
                    InvokeType = "OTHR";

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

            return str;
        }
        private void PrintData()
        {
            Row = 10;
            R1 = 0;
            BeginReport(1100, 800);
            AddPage(1100, 800);
            WriteAddress();
            WriteHouse();
            EndReport();
        }
        private void WriteAddress()
        {
            HCOL1 = 20; HCOL2 = 190; HCOL3 = 280; HCOL4 = 380; HCOL5 = 500; HCOL6 = 600; HCOL7 = 700; HCOL8 = 775; HCOL9 = 950; HCOL10 = 1050;
            ROW_HT = 15;
                        ifontSize = 9;
        }
        private void WriteHouse()
        {
            Dictionary<int, string> AddrList = null;
            int iAddr = 0;
            string Str = "";
            string[] sData = null;
            string s1 = "";
            string s2 = "";
            R1++;
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
            if (GetPosition("AIRLINE_CODE"))
                AddXYLabel(x1, y1, h1, w1, s1, fname, fsize, "", sStyle);
            if (GetPosition("POL_CODE"))
            {
                Str = mRow.bl_pol_code.ToString();
                if (Str.StartsWith("IN") && Str.Length >= 5) //INPOL4
                {
                    Str = Str.Substring(2, 3);
                }
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);// MblRec.mbld_pol_code
            }
            if (GetPosition("MAWB_NO_TOP_LEFT"))
                AddXYLabel(x1, y1, h1, w1, s2, fname, fsize, "", sStyle);

            if (GetPosition("MAWB_NO_TOP_RIGHT"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_mbl_no.ToString(), fname, fsize, "", sStyle);
            if (GetPosition("HAWB_NO_TOP"))
                AddXYLabel(x1, y1, h1, w1, mRow.hbl_bl_no.ToString(), fname, fsize, "", sStyle);

            R1++;
            R1++;
            if (GetPosition("SHIPPER_NAME"))
            {
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

                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_shipper_name.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, (AddrList.Count > 0) ? AddrList[0] : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, (AddrList.Count > 1) ? AddrList[1] : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, (AddrList.Count > 2) ? AddrList[2] : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, (AddrList.Count > 3) ? AddrList[3] : "", fname, fsize, "", sStyle);
            }

            R1++;

            if (GetPosition("ISSUED_BY1"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_issued_by1.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_issued_by2.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_issued_by3.ToString(), fname, fsize, "", sStyle);
            }
            R1++;
            R1++;
            R1++;
            R1++;

            if (GetPosition("CONSIGNEE_NAME"))
            {
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
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_consignee_name.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, (AddrList.Count > 0) ? AddrList[0] : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, (AddrList.Count > 1) ? AddrList[1] : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, (AddrList.Count > 2) ? AddrList[2] : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, (AddrList.Count > 3) ? AddrList[3] : "", fname, fsize, "", sStyle);
            }
            R1++;
            R1++;
            R1++;
            R1++;
            R1++;
            R1++;
            R1++;
            if (GetPosition("ISSUING_AGENT_NAME"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_issu_agnt_name.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_issu_agnt_city.ToString(), fname, fsize, "", sStyle);
            }
            R1++;

            if (GetPosition("IATA_CARRIER"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_iata_carrier.ToString(), fname, fsize, "", sStyle);
            R1++;
            if (GetPosition("ACCOUNTING_INFORMATION1"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_account_info1.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_account_info2.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_account_info3.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_account_info4.ToString(), fname, fsize, "", sStyle);
            }

            R1++;
            R1++;
            R1++;

            if (GetPosition("IATA_CODE"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_iata_code.ToString(), fname, fsize, "", sStyle);
            if (GetPosition("ACCOUNT_NO"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_acc_no.ToString(), fname, fsize, "", sStyle);
            R1++;
            R1++;
            if (GetPosition("AIRPORT_DEPARTURE"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_pol.ToString(), fname, fsize, "", sStyle); //MblRec.mbl_pol_name
            // END OF FIRST PART

            // START Of PART 2
            

            if (GetPosition("TO1"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_to1.ToString(), fname, fsize, "", sStyle); //mbl_to_port1
            if (GetPosition("BY1"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_by1.ToString(), fname, fsize, "", sStyle);//mbl_by_carrier1
            if (GetPosition("TO2"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_to2.ToString(), fname, fsize, "", sStyle);//mbl_to_port2
            if (GetPosition("BY2"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_by2.ToString(), fname, fsize, "", sStyle);//mbl_by_carrier2
            if (GetPosition("TO3"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_to3.ToString(), fname, fsize, "", sStyle);
            if (GetPosition("BY3"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_by3.ToString(), fname, fsize, "", sStyle);
            if (GetPosition("CURRENCY"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_currency.ToString(), fname, fsize, "", sStyle); //MblRec.mbl_currency
            if (GetPosition("CHARGES CODE"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_frt_status.ToString().Trim() + mRow.bl_oc_status.ToString().Trim(), fname, fsize, "", sStyle);
            if (GetPosition("FRT_PPD"))
                AddXYLabel(x1, y1, h1, w1, (mRow.bl_frt_status.ToString().Trim() == "P") ? "X" : "", fname, fsize, "", sStyle);
            if (GetPosition("FRT_COLL"))
                AddXYLabel(x1, y1, h1, w1, (mRow.bl_frt_status.ToString().Trim() == "C") ? "X" : "", fname, fsize, "", sStyle);
            if (GetPosition("OTH_PPD"))
                AddXYLabel(x1, y1, h1, w1, (mRow.bl_oc_status.ToString().Trim() == "P") ? "X" : "", fname, fsize, "", sStyle);
            if (GetPosition("OTH_COLL"))
                AddXYLabel(x1, y1, h1, w1, (mRow.bl_oc_status.ToString().Trim() == "C") ? "X" : "", fname, fsize, "", sStyle);
            if (GetPosition("VALUE_CARRIAGE"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_carriage_value.ToString(), fname, fsize, "", sStyle);
            if (GetPosition("VALUE_CUSTOMS"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_customs_value.ToString(), fname, fsize, "", sStyle);

 
            if (GetPosition("AIRPORT_DESTINATION"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_pod.ToString(), fname, fsize, "", sStyle);//mbl_pod_name

            if (GetPosition("FLIGHT_DATE1"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_flight1.ToString(), fname, fsize, "", sStyle);

            
            if (GetPosition("FLIGHT_DATE2"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_flight2.ToString(), fname, fsize, "", sStyle);

            if (GetPosition("INSURANCE_AMOUNT"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_ins_amt.ToString(), fname, fsize, "", sStyle);
 

            if (GetPosition("HANDLING_INFORMATION1"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_hand_info1.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_hand_info2.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_hand_info3.ToString(), fname, fsize, "", sStyle);
            }
           
            WriteDescription();
           
            WriteOtherCharges();

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
            

            if (GetPosition("WEIGHT_CHARGE_PP"))
            {
                if (frtppAsArrngd.Trim().Length > 0)
                    Str = frtppAsArrngd;
                else
                    Str = (frtppAmt > 0) ? frtppAmt.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);
            }

            if (GetPosition("WEIGHT_CHARGE_CC"))
            {
                if (frtccAsArrngd.Trim().Length > 0)
                    Str = frtccAsArrngd;
                else
                    Str = (frtccAmt > 0) ? frtccAmt.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);
            }
       
            FindTotal();

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

            if (GetPosition("OTHER_CHARGES_AGENT_PP"))
            {
                if (OthppAsArrngd.Trim().Length > 0)
                    Str = OthppAsArrngd;
                else
                    Str = (Agent_Tot_PP > 0) ? Agent_Tot_PP.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);
            }


            if (GetPosition("OTHER_CHARGES_AGENT_CC"))
            {
                if (OthccAsArrngd.Trim().Length > 0)
                    Str = OthccAsArrngd;
                else
                    Str = (Agent_Tot_CC > 0) ? Agent_Tot_CC.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);
            }

            R1++;

            if (GetPosition("OTHER_CHARGES_CARRIER_PP"))
            {
                if (OthppAsArrngd.Trim().Length > 0)
                    Str = OthppAsArrngd;
                else
                    Str = (Carrier_Tot_PP > 0) ? Carrier_Tot_PP.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);
            }

            if (GetPosition("OTHER_CHARGES_CARRIER_CC"))
            {
                if (OthccAsArrngd.Trim().Length > 0)
                    Str = OthccAsArrngd;
                else
                    Str = (Carrier_Tot_CC > 0) ? Carrier_Tot_CC.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);
            }


            if (GetPosition("AGENT_FOR_SHIPPER1"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_by1_agent.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_by2_agent.ToString(), fname, fsize, "", sStyle);
            }
            R1++;
            R1++;
            R1++;
            if (GetPosition("TOTAL_PP"))
            {
                if (OthppAsArrngd.Trim().Length > 0 || frtppAsArrngd.Trim().Length > 0)
                    Str = "AS ARRANGED";
                else
                    Str = (Agent_Tot_PP + Carrier_Tot_PP + frtppAmt > 0) ? (Agent_Tot_PP + Carrier_Tot_PP + frtppAmt).ToString() : "";
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);
            }
            if (GetPosition("TOTAL_CC"))
            {
                if (OthccAsArrngd.Trim().Length > 0 || frtccAsArrngd.Trim().Length > 0)
                    Str = "AS ARRANGED";
                else
                    Str = (Agent_Tot_CC + Carrier_Tot_CC + frtccAmt > 0) ? (Agent_Tot_CC + Carrier_Tot_CC + frtccAmt).ToString() : "";
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);
            }
            if (GetPosition("AGENT_FOR_CARRIER1"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_by1_carrier.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_by2_carrier.ToString(), fname, fsize, "", sStyle);
            }
            R1++;
            Str = mRow.bl_issued_date_print;
            if (GetPosition("ISSUED_DATE"))
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);

            if (GetPosition("ISSUED_PLACE"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_issued_place.ToString(), fname, fsize, "", sStyle);

            if (GetPosition("ISSUED_BY_BOTTOM"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_issued_by.ToString(), fname, fsize, "", sStyle);

            if (GetPosition("MAWB_NO_BOTTOM"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_mbl_no.ToString(), fname, fsize, "", sStyle, R1, 71, 0, 16, 0);
            if (GetPosition("HAWB_NO_BOTTOM"))
                AddXYLabel(x1, y1, h1, w1, mRow.hbl_bl_no.ToString(), fname, fsize, "", sStyle, R1, 71, 0, 16, 0);

            Str = FooterNote;
            //Row += ROW_HT + 5; R1++; 
            if (GetPosition("NO_OF_COPIES"))
                AddXYLabel(x1, y1, h1, w1, Str, fname, fsize, "", sStyle);//, R1, 71, 0, 16, 0
        }
        private void WriteDescription()
        {
            string str = "";
      
            if (GetPosition("PCS"))
            {
                str = (Lib.Conv2Integer(mRow.bl_pcs.ToString()) > 0) ? mRow.bl_pcs.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
      
            }

            if (GetPosition("GR_WEIGHT"))
            {
                str = (Lib.Conv2Decimal(mRow.bl_grwt.ToString()) > 0) ? mRow.bl_grwt.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
            }
            if (GetPosition("GR_WEIGHT_UNIT"))
            {
                str = (mRow.bl_wt_unit.ToString().Trim().Length > 0) ? mRow.bl_wt_unit.ToString().Trim() : "";//.Substring(0, 1)
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);//9
                                                                          //  s3 = str;
            }

           
            if (GetPosition("CLASS"))
            {
                str = (mRow.bl_class.ToString().Trim().Length > 0) ? mRow.bl_class.ToString().Trim() : "";//.Substring(0, 1)
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
               
            }
            if (GetPosition("COMMODITY"))
            {
                AddXYLabel(x1, y1, h1, w1, mRow.bl_commodity.ToString(), fname, fsize, "", sStyle);
              

            }
             
            if (GetPosition("CH_WEIGHT"))
            {
                str = (Lib.Conv2Integer(mRow.bl_chwt.ToString()) > 0) ? mRow.bl_chwt.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
                 
            }
            if (GetPosition("CH_WEIGHT_UNIT"))
            {
                str = (mRow.bl_wt_unit.ToString().Trim().Length > 0) ? mRow.bl_wt_unit.ToString().Trim() : "";//.Substring(0, 1)
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
                
            }

            if ((mRow.bl_asarranged_shipper.ToString().Trim() == "Y") || (mRow.bl_asarranged_consignee.ToString().Trim() == "Y"))
            {
                str = "AS ARRANGED";
                if (GetPosition("RATE"))
                    AddXYLabel(x1, y1, h1, w1, "", fname, fsize, "", sStyle);
                if (GetPosition("TOTAL"))
                    AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
                
            }
            else
            {
                if (GetPosition("RATE"))
                    AddXYLabel(x1, y1, h1, w1, (Lib.Conv2Decimal(mRow.bl_rate.ToString()) > 0) ? mRow.bl_rate.ToString() : "", fname, fsize, "", sStyle);
                if (GetPosition("TOTAL"))
                    AddXYLabel(x1, y1, h1, w1, (Lib.Conv2Decimal(mRow.bl_total.ToString()) > 0) ? mRow.bl_total.ToString() : "", fname, fsize, "", sStyle);
              
            }
            
            if (GetPosition("TOTAL_PCS"))
            {
                str = (Lib.Conv2Integer(mRow.bl_pcs.ToString()) > 0) ? mRow.bl_pcs.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
                
            }
            if (GetPosition("TOTAL_GR_WEIGHT"))
            {
                str = (Lib.Conv2Decimal(mRow.bl_grwt.ToString()) > 0) ? mRow.bl_grwt.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
               
            }
            if (GetPosition("TOTAL_GR_WEIGHT_UNIT"))
            {
                str = (mRow.bl_wt_unit.ToString().Trim().Length > 0) ? mRow.bl_wt_unit.ToString().Trim() : "";//.Substring(0, 1)
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
                 
            }

            if (GetPosition("TOTAL_CH_WEIGHT"))
            {
                str = (Lib.Conv2Decimal(mRow.bl_chwt.ToString()) > 0) ? mRow.bl_chwt.ToString() : "";
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
                
            }
            if (GetPosition("TOTAL_CH_WEIGHT_UNIT"))
            {
                str = (mRow.bl_wt_unit.ToString().Trim().Length > 0) ? mRow.bl_wt_unit.ToString().Trim() : "";//.Substring(0, 1)
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
                
            }

            if ((mRow.bl_asarranged_shipper.ToString().Trim() == "Y") || (mRow.bl_asarranged_consignee.ToString().Trim() == "Y"))
            {
                str = "AS ARRANGED";
                if (GetPosition("TOTAL_RATE"))
                    AddXYLabel(x1, y1, h1, w1, "", fname, fsize, "", sStyle);
                if (GetPosition("GRANDTOTAL"))
                    AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
                
            }
            else
            {
                if (GetPosition("TOTAL_RATE"))
                    AddXYLabel(x1, y1, h1, w1, (Lib.Conv2Decimal(mRow.bl_rate.ToString()) > 0) ? mRow.bl_rate.ToString() : "", fname, fsize, "", sStyle);
                if (GetPosition("GRANDTOTAL"))
                    AddXYLabel(x1, y1, h1, w1, (Lib.Conv2Decimal(mRow.bl_total.ToString()) > 0) ? mRow.bl_total.ToString() : "", fname, fsize, "", sStyle);
         
            }

            if (GetPosition("MARKS_NO1"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark1 != null ? mRow.bl_mark1.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark2 != null ? mRow.bl_mark2.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark3 != null ? mRow.bl_mark3.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark4 != null ? mRow.bl_mark4.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark5 != null ? mRow.bl_mark5.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark6 != null ? mRow.bl_mark6.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark7 != null ? mRow.bl_mark7.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark8 != null ? mRow.bl_mark8.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark9 != null ? mRow.bl_mark9.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark10 != null ? mRow.bl_mark10.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark11 != null ? mRow.bl_mark11.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark12 != null ? mRow.bl_mark12.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark13 != null ? mRow.bl_mark13.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark14 != null ? mRow.bl_mark14.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark15 != null ? mRow.bl_mark15.ToString() : "", fname, fsize, "", sStyle);
            }
            if (GetPosition("NATURE_OF_GOODS1"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc1 != null ? mRow.bl_desc1.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc2 != null ? mRow.bl_desc2.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc3 != null ? mRow.bl_desc3.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc4 != null ? mRow.bl_desc4.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc5 != null ? mRow.bl_desc5.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc6 != null ? mRow.bl_desc6.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc7 != null ? mRow.bl_desc7.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc8 != null ? mRow.bl_desc8.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc9 != null ? mRow.bl_desc9.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc10 != null ? mRow.bl_desc10.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc11 != null ? mRow.bl_desc11.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc12 != null ? mRow.bl_desc12.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc13 != null ? mRow.bl_desc13.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc14 != null ? mRow.bl_desc14.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc15 != null ? mRow.bl_desc15.ToString() : "", fname, fsize, "", sStyle);
            }


        }


        private void WriteOtherCharges()
        {
            //OthChrgStartRow += 3;
             float OthChrgHCOL = HCOL1 + 300;
            int PrntRowCount = 0;
            string Chrgs = "";
            int c1 = 0;
            if (GetPosition("OTHER_CHARGES_CARRIER1"))
            {
                Row = y1;
                c1 = 0;

                if (GetOthData(mRow.bl_charges1_carrier))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges Carrier
                        Row += h1;
                        PrntRowCount++;
                        CH1[c1] = Chrgs; c1++;
                    }
                if (GetOthData(mRow.bl_charges2_carrier))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges Carrier
                        Row += h1;
                        PrntRowCount++;
                        CH1[c1] = Chrgs; c1++;
                    }
                if (GetOthData(mRow.bl_charges3_carrier))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges Carrier
                        Row += h1;
                        PrntRowCount++;
                        CH1[c1] = Chrgs; c1++;
                    }
                if (GetOthData(mRow.bl_charges4_carrier))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges Carrier
                        Row += h1;
                        PrntRowCount++;
                        CH1[c1] = Chrgs; c1++;
                    }
                if (GetOthData(mRow.bl_charges5_carrier))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges Carrier
                        Row += h1;
                        PrntRowCount++;
                        CH1[c1] = Chrgs; c1++;
                    }
                if (GetOthData(mRow.bl_charges6_carrier))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges Carrier
                        Row += h1;
                        PrntRowCount++;
                        CH1[c1] = Chrgs; c1++;
                    }
                
            }


            if (GetPosition("OTHER_CHARGES_AGENT1"))
            {
                Row = y1;
                c1 = 0;

                if (GetOthData(mRow.bl_charges1_agent))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges Agent
                        Row += h1;
                        PrntRowCount++;
                        CH2[c1] = Chrgs; c1++;
                    }
                if (GetOthData(mRow.bl_charges2_agent))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges Ahgent        
                        Row += h1;
                        PrntRowCount++;
                        CH2[c1] = Chrgs; c1++;
                    }
                if (GetOthData(mRow.bl_charges3_agent))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges agent
                        Row += h1;
                        PrntRowCount++;
                        CH2[c1] = Chrgs; c1++;
                    }
                if (GetOthData(mRow.bl_charges4_agent))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges agent
                        Row += h1;
                        PrntRowCount++;
                        CH2[c1] = Chrgs; c1++;
                    }

                if (GetOthData(mRow.bl_charges5_agent))
                    if ((InvokeType == "SHPR" && OthPrintS.Trim() == "Y" && OthCharges.Trim().Length > 0) ||
                      (InvokeType == "CNEE" && OthPrintC.Trim() == "Y" && OthCharges.Trim().Length > 0) || InvokeType == "OTHR")
                    {
                        Chrgs = GetFormatOtherChrgs(OthCharges, OthRate, OthTotal);
                        AddXYLabel(x1, Row, h1, w1, Chrgs, fname, fsize, "", sStyle);//other charges agent
                        Row += h1;
                        PrntRowCount++;
                        CH2[c1] = Chrgs; c1++;
                    }
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
        private string getCopy(string sCopy)
        {
            int i = Lib.Conv2Integer(sCopy);
            if (i == 0)
                return "ZERO(0)";
            else if (i == 1)
                return "ONE(1)";
            else if (i == 2)
                return "TWO(2)";
            else if (i == 3)
                return "THREE(3)";
            else if (i == 4)
                return "FOUR(4)";
            else if (i == 5)
                return "FIVE(5)";
            else
                return "THREE(3)";
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

                    bRet = true;
                }
            }
            catch (Exception)
            {
                bRet = false;
            }
            return bRet;
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
