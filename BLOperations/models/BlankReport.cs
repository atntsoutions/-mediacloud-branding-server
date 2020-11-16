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
    public class BlankReport : BaseReport
    {

        public Bl mRow = null;
        public string RootPath = "";
        public DataTable Dt_COLPOS = new DataTable();

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
        private float Xtolrnce = 0;
        private string sError = "";
        private int x1 = 0, y1 = 0, h1 = 0, w1 = 0, fsize = 0;
        private string fname = "", sStyle = "";

        public BlankReport()
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

            addList("XLCOLUMN", "100", "90", "20", "100", "100", "150", "30", "20", "30", "20", "60");
            BeginReport(1100, 800);
            WriteData();
            EndReport();
        }
        private void WriteData()
        {
            Row = 10;

            AddPage(1100, 800);
            FillData();

            if (mRow.AttachList.Count > 0)
            {                AddPage(1100, 800);
                FillAttachedData();
            }
        }
        private void FillData()
        {
            string str = "";

            if (GetPosition("BL_BL_NO"))
                AddXYLabel(x1, y1, h1, w1, mRow.hbl_bl_no.ToString(), fname, fsize, "", sStyle);

            if (GetPosition("BL_BL_NO2"))
                AddXYLabel(x1, y1, h1, w1, mRow.hbl_bl_no.ToString(), fname, fsize, "", sStyle);
            if (GetPosition("BL_SHIPPER_NAME"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_shipper_name.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_shipper_add1.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_shipper_add2.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_shipper_add3.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_shipper_add4.ToString(), fname, fsize, "", sStyle);
            }


            if (GetPosition("BL_CONSIGNEE_NAME"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_consignee_name.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_consignee_add1.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_consignee_add2.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_consignee_add3.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_consignee_add4.ToString(), fname, fsize, "", sStyle);
            }


            if (GetPosition("BL_NOTIFY_NAME"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_notify_name.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_notify_add1.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_notify_add2.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_notify_add3.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_notify_add4.ToString(), fname, fsize, "", sStyle);
            }


            if (GetPosition("BL_DELIVERY_CONTACT1"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_delivery_contact1.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_delivery_contact2.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_delivery_contact3.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_delivery_contact4.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_delivery_contact5.ToString(), fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_delivery_contact6.ToString(), fname, fsize, "", sStyle);
            }

            if (GetPosition("BL_PLACE_RECEIPT"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_place_receipt.ToString(), fname, fsize, "", sStyle);

            if (GetPosition("BL_DATE_RECEIPT"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_date_receipt_print.ToString(), fname, fsize, "", sStyle);


            if (GetPosition("BL_POL"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_pol.ToString(), fname, fsize, "", sStyle);



            if (GetPosition("BL_POD"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_pod.ToString(), fname, fsize, "", sStyle);

            if (GetPosition("BL_PLACE_DELIVERY"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_place_delivery.ToString(), fname, fsize, "", sStyle);



            if (GetPosition("BL_VSL_VOYNO"))
            {
                str = mRow.bl_vsl_name.ToString() + " " + mRow.bl_vsl_voy_no.ToString();
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
            }

            if (GetPosition("BL_PERIOD_DELIVERY"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_period_delivery.ToString(), fname, fsize, "", sStyle);
            if (GetPosition("BL_MOVE_TYPE"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_move_type.ToString(), fname, fsize, "", sStyle);
            if (GetPosition("BL_PLACE_TRANSHIPMENT"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_place_transhipment.ToString(), fname, fsize, "", sStyle);

            //if (GetPosition("BL_PKGS"))
            //    AddXYLabel(x1, y1, h1, w1, mRow.bl_KGS.ToString(), fname, fsize, "", sStyle);

            if (GetPosition("BL_GRWT_KGS"))
            {
                str = mRow.bl_grwt_caption != null ? mRow.bl_grwt_caption.ToString() : "";
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, str, fname, fsize, "", sStyle);
                Row += h1;
                str = "";
                if (Lib.Conv2Decimal(mRow.bl_grwt.ToString()) > 0)
                {
                    // str = Common.NumericFormat(DR["BL_GR_WT"].ToString(), 3);
                    //str = Lib.Conv2Decimal(mRow.bl_grwt.ToString()).ToString("0.00#");
                    str = Lib.Conv2Decimal(mRow.bl_grwt.ToString()).ToString("0.000"); //Request by baiju delsea for seabridge format
                }
                AddXYLabel(x1, Row, h1, w1, str, fname, fsize, "", sStyle);
            }

            if (GetPosition("BL_NETWT_KGS"))
            {
                str = mRow.bl_ntwt_caption != null ? mRow.bl_ntwt_caption.ToString() : "";
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, str, fname, fsize, "", sStyle);
                Row += h1;
                str = "";
                if (Lib.Conv2Decimal(mRow.bl_ntwt.ToString()) > 0)
                {
                    //str = Common.NumericFormat(DR["BL_NET_WT"].ToString(), 3);
                    // str = Lib.Conv2Decimal(mRow.bl_ntwt.ToString()).ToString("0.00#");
                    str = Lib.Conv2Decimal(mRow.bl_ntwt.ToString()).ToString("0.000");
                }
                AddXYLabel(x1, Row, h1, w1, str, fname, fsize, str, sStyle);
            }

            if (GetPosition("BL_CBM"))
            {
                if (Lib.Conv2Decimal(mRow.bl_cbm.ToString()) > 0)
                {
                    str = mRow.bl_cbm_caption != null ? mRow.bl_cbm_caption.ToString() : "";
                    Row = y1;
                    AddXYLabel(x1, Row, h1, w1, str, fname, fsize, "", sStyle);
                    Row += h1;
                    // str = Common.NumericFormat(DR["BL_CBM"].ToString(), 3);
                    //str = Lib.Conv2Decimal(mRow.bl_cbm.ToString()).ToString("0.00#");
                    str = Lib.Conv2Decimal(mRow.bl_cbm.ToString()).ToString("0.000");//Request by baiju delsea for seabridge format
                    AddXYLabel(x1, Row, h1, w1, str, fname, fsize, str, sStyle);
                }
            }

            if (GetPosition("BL_QTY"))
            {
                str = "";
                if (Lib.Conv2Decimal(mRow.bl_pcs.ToString()) > 0)
                {
                    str = mRow.bl_pcs_caption != null ? mRow.bl_pcs_caption.ToString() : "";
                    str += " / " + mRow.bl_pcs_unit.ToString();
                }
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, str, fname, fsize, "", sStyle);
                Row += h1;
                str = "";
                if (Lib.Conv2Decimal(mRow.bl_pcs.ToString()) > 0)
                    str = Lib.NumFormat(mRow.bl_pcs.ToString(), 0);
                AddXYLabel(x1, Row, h1, w1, str, fname, fsize, str, sStyle);
            }


            if (GetPosition("BL_MARKSNO_START"))
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
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark16 != null ? mRow.bl_mark16.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark17 != null ? mRow.bl_mark17.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark18 != null ? mRow.bl_mark18.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark19 != null ? mRow.bl_mark19.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark20 != null ? mRow.bl_mark20.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark21 != null ? mRow.bl_mark21.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark22 != null ? mRow.bl_mark22.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark23 != null ? mRow.bl_mark23.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_mark24 != null ? mRow.bl_mark24.ToString() : "", fname, fsize, "", sStyle);
            }
            if (GetPosition("BL_DESCRIPTION_START"))
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
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc16 != null ? mRow.bl_desc16.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc17 != null ? mRow.bl_desc17.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_desc18 != null ? mRow.bl_desc18.ToString() : "", fname, fsize, "", sStyle);
            }
            if (GetPosition("BL_DESCRIPTION_2START"))
            {
                Row = y1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_2desc7 != null ? mRow.bl_2desc7.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_2desc8 != null ? mRow.bl_2desc8.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_2desc9 != null ? mRow.bl_2desc9.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_2desc10 != null ? mRow.bl_2desc10.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_2desc11 != null ? mRow.bl_2desc11.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_2desc12 != null ? mRow.bl_2desc12.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_2desc13 != null ? mRow.bl_2desc13.ToString() : "", fname, fsize, "", sStyle);
                Row += h1;
                AddXYLabel(x1, Row, h1, w1, mRow.bl_2desc14 != null ? mRow.bl_2desc14.ToString() : "", fname, fsize, "", sStyle);

            }
            if (GetPosition("BL_FRT_AMOUNT"))
            {
                str = "";
                if (Lib.Conv2Decimal(mRow.bl_frt_amount.ToString()) != 0)
                    str = "BL_FRT_AMOUNT";
                AddXYLabel(x1, Row, h1, w1, str, fname, fsize, str, sStyle);
            }

            if (GetPosition("BL_FRT_PAY_AT"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_frt_pay_at.ToString(), fname, fsize, "", sStyle);

            if (GetPosition("BL_NO_COPIES"))
                AddXYLabel(x1, y1, h1, w1, getCopy(mRow.bl_no_copies.ToString()), fname, fsize, "", sStyle);


            if (GetPosition("BL_ISSUED_PLACE_DATE"))
            {
                str = mRow.bl_issued_place.ToString() + " " + mRow.bl_issued_date_print.ToString();
                AddXYLabel(x1, y1, h1, w1, str, fname, fsize, "", sStyle);
            }

            if (GetPosition("BL_FREIGHT_STATUS"))
                AddXYLabel(x1, y1, h1, w1, mRow.bl_frt_status.ToString(), fname, fsize, "", sStyle);
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

        private void FillAttachedData()
        {
            ifontSize = 9;
            
            ifontName = "Arial";

            HCOL1 = 30;
            HCOL2 = HCOL1 + 120;
            HCOL3 = HCOL2 + 180;
            HCOL4 = HCOL3 + 150;
            HCOL5 = HCOL4 + 280;

            bool IsMarksNoExist = false;
            bool IsDescExist = false;

            Row = 100;
            ROW_HT = 16;
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "ATTACHED SHEET", ifontName, ifontSize, "", "CB", R1, XL_COLA, 0, 20);

            Row += ROW_HT;
            Row += ROW_HT; R1++;

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "BL NO", ifontName, ifontSize, "", "B", R1, XL_COLA, 2, 16);
            AddXYLabel(HCOL2 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, mRow.hbl_bl_no.ToString(), ifontName, ifontSize, "", "", R1, XL_COLD, 0, 16);
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "BL DATE", ifontName, ifontSize, "", "B", R1, XL_COLE, 0, 16);
            AddXYLabel(HCOL4 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            //AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, Lib.DatetoStringDisplayformat(mRow.bl_issued_date), ifontName, ifontSize, "", "B", R1, XL_COLF, 5, 16);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, mRow.bl_issued_date_print, ifontName, ifontSize, "", "B", R1, XL_COLF, 5, 16);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "VESSEL", ifontName, ifontSize, "", "B", R1, XL_COLA, 2, 16);
            AddXYLabel(HCOL2 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL5 - HCOL2, mRow.bl_vsl_name.ToString(), ifontName, ifontSize, "", "", R1, XL_COLD, 0, 16);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "VOYAGE", ifontName, ifontSize, "", "B", R1, XL_COLA, 2, 16);
            AddXYLabel(HCOL2 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL5 - HCOL2, mRow.bl_vsl_voy_no.ToString(), ifontName, ifontSize, "", "", R1, XL_COLD, 0, 16);

            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;
            IsMarksNoExist = false; IsDescExist = false;

            foreach (Bldesc Rec in mRow.AttachList)
            {
                if (Rec.bl_marks != null)
                    if (Rec.bl_marks.Trim().Length > 0)
                        IsMarksNoExist = true;

                if (Rec.bl_desc != null)
                    if (Rec.bl_desc.Trim().Length > 0)
                        IsDescExist = true;
                if (IsMarksNoExist && IsDescExist)
                    break;
            }
            foreach (Bldesc Rec in mRow.AttachList.OrderBy(x => x.bl_desc_ctr))
            {
                if (IsMarksNoExist && IsDescExist)
                {
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, Rec.bl_marks != null ? Rec.bl_marks.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 2, 16);
                    AddXYLabel(HCOL3, Row, ROW_HT, HCOL5 - HCOL3, Rec.bl_desc != null ? Rec.bl_desc.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 0, 16);
                }
                else if (IsMarksNoExist)
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, Rec.bl_marks != null ? Rec.bl_marks.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 2, 16);
                else
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, Rec.bl_desc != null ? Rec.bl_desc.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 2, 16);
                Row += ROW_HT; R1++;
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
    }
}
