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
    public class FCRReport : BaseReport
    {
        public string dColr = "2";
        public Bl mRow = null;
        public string InvokeType = "";
        public Boolean Chk_BL_Original = false;
        public string RootPath = "";
         

        private int ifontSizesm = 7;
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
        private float DescStartRow = 0;
        
        private string BsideStyle = "";
        private char sColSplit = '~';
        private char sStyleSplit = '#';

        public FCRReport()
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

            if (mRow.hbl_fcr_no.ToString().Trim() == "")
                Lib.AddError(ref str, " FCR NO. not Found ");

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

            CanWrite = true;
            AddPage(1100, 800);
            WriteBackSide();
            CanWrite = true;
            if (mRow.AttachList != null)
                if (mRow.AttachList.Count > 0)
                {
                    AddPage(1100, 800);
                    FillAttachedData();
                }
        }
        private void FillData()
        {
            string str = "";
            ifontSize = 8;
            ifontSizesm = 7;
            ifontName = "Arial";

            HCOL1 = 20;
            HCOL2 = HCOL1 + 120;
            HCOL3 = HCOL2 + 80;
            HCOL4 = HCOL3 + 30;
            HCOL5 = HCOL4 + 160;
            HCOL6 = HCOL5 + 188;
            HCOL7 = HCOL6 + 188;


            Row = 40;
            ROW_HT = 15; R1++;

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "SHIPPER", ifontName, ifontSizesm, "LT", "B", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize + 2, "TLR", "CB", R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL5, Row, ROW_HT * 3, HCOL7 - HCOL5, "FORWARDER'S CARGO RECEIPT", "Arialnarrow", ifontSize + 7, "", "CB", R1, XL_COLF, 0, 16, 0, 0, 0, 20);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_name.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize + 2, "LR", "CB", R1, XL_COLF, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_add1.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize + 2, "BLR", "CB", R1, XL_COLF, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_add2.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "DOCK RECEIPT NO.", ifontName, ifontSizesm, "LR", "B", R1, XL_COLF, 4, 16, 0, 0, Xtolrnce);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_add3.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "BLR", "", R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL5, Row - 8, ROW_HT, HCOL7 - HCOL5, mRow.hbl_fcr_no.ToString(), ifontName, ifontSize, "", "CB", R1, XL_COLF, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_add4.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16);


            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "CONSIGNEE", ifontName, ifontSizesm, "LT", "B", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);

            float sRow = Row, sRow_HT = 20;
            AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, "CARGOMAR", "Times New Roman", ifontSize + 8, "", "BC", R1, XL_COLF, 5, 16, 0, 0, 0, 20);//LR2
            sRow += sRow_HT; sRow_HT = 14;

            LoadImage(RootPath + "\\CARGOMAR.jpg", HCOL6 - 23, Row + 19, 48, 48);

            sRow += sRow_HT * 3;
            sRow += 12;
            AddXYLabel(HCOL5 + 83, sRow, sRow_HT, 210, "THIS IS NOT A DOCUMENT OF TITLE", ifontName, ifontSize, "TB", "CB", R1, XL_COLF, 5, 16, 0, 0, 0, 9);//LR2
            sRow += sRow_HT;
            sRow += 8;

            AddXYLabel(HCOL5, sRow, sRow_HT + 2, HCOL7 - HCOL5, "CARGOMAR PRIVATE LIMITED", ifontName, ifontSize + 4, "", "CB", R1, XL_COLF, 5, 16, 0, 0, 0, 15);//LR2

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_name.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);

            DrawVLine(HCOL5, Row, ROW_HT * 10, "");
            DrawVLine(HCOL7, Row, ROW_HT * 10, "");

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_add1.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_add2.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_add3.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_add4.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);


            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "NOTIFY", ifontName, ifontSizesm, "LT", "B", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_name.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            // AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);


            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_add1.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            // AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);

            sRow = Row; sRow_HT = 13;

            str = "RECEIVED BY CARGOMAR PRIVATE LIMITED IN APPARENT GOOD";
            AddXYLabel(HCOL5 + 10, sRow, sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "ORDER AND CONDITION FROM THE SHIPPER, THE PACKAGE(S) LISTED";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 1), sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "BELOW SAID TO CONTAIN THE GOODS HEREINAFTER DESCRIBED. THIS";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 2), sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "RECEIPT IS NOT VALID UNLESS VERIFIED AND SIGNED BY AN";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 3), sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "AUTHORISED SIGNATORY OF CARGOMAR PRIVATE LIMITED. THIS";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 4), sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "FORWARDER'S CARGO RECEIPT IS TO BE ISSUED UPON";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 5), sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "PRESENTATION OF THE CORRESPONDING DOCK RECEIPT, CARGO AS";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 6), sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "NOTED BELOW WILL BE CONTAINERIZED AND SHIPPED UNDER";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 7), sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "OCEAN BILL(S) OF LADING, TO BE ISSUED AND SIGNED BY CARRIER.";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 8), sRow_HT, HCOL7 - (HCOL5 + 25), str, ifontName, ifontSizesm, "", "", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_add2.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            // AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_add3.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16, 0, 0, Xtolrnce);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_add4.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT + 5, (HCOL4 - 40) - HCOL1, "VESSEL VOYAGE (INTENDED)", ifontName, ifontSizesm, "LT", "B", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL4 - 40, Row, ROW_HT + 5, HCOL5 - (HCOL4 - 40), "DATE OF RECEIPT", ifontName, ifontSizesm, "LT", "B", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT + 5, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16);

            Row += ROW_HT + 5; R1++;
            str = (mRow.bl_vsl_name != null ? mRow.bl_vsl_name.ToString() : "") + " " + (mRow.bl_vsl_voy_no != null ? mRow.bl_vsl_voy_no.ToString() : "");
            AddXYLabel(HCOL1, Row - 2, ROW_HT + 5, (HCOL4 - 40) - HCOL1, str, ifontName, ifontSize, "L", "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL4 - 40, Row - 2, ROW_HT + 5, HCOL5 - (HCOL4 - 40), (mRow.bl_date_receipt_print != null ? mRow.bl_date_receipt_print.ToString() : ""), ifontName, ifontSize, "L", "", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT + 5, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16);


            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT + 5, (HCOL4 - 40) - HCOL1, "PLACE OF RECEIPT", ifontName, ifontSizesm, "LT", "B", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL4 - 40, Row, ROW_HT + 5, HCOL5 - (HCOL4 - 40), "PORT OF LOADING", ifontName, ifontSizesm, "LT", "B", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT + 5, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16);

            Row += ROW_HT + 5; R1++;
            AddXYLabel(HCOL1, Row - 2, ROW_HT + 5, (HCOL4 - 40) - HCOL1, (mRow.bl_place_receipt != null ? mRow.bl_place_receipt.ToString() : ""), ifontName, ifontSize, "L", "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL4 - 40, Row - 2, ROW_HT + 5, HCOL5 - (HCOL4 - 40), mRow.bl_pol.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16); ;

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT + 5, (HCOL4 - 40) - HCOL1, "PORT OF DISCHARGE", ifontName, ifontSizesm, "LT", "B", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL4 - 40, Row, ROW_HT + 5, HCOL5 - (HCOL4 - 40), "PLACE OF DELIVERY", ifontName, ifontSizesm, "LRT", "B", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT + 5, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16); ;

            Row += ROW_HT + 5; R1++;
            AddXYLabel(HCOL1, Row - 2, ROW_HT + 5, (HCOL4 - 40) - HCOL1, mRow.bl_pod.ToString(), ifontName, ifontSize, "L", "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL4 - 40, Row - 2, ROW_HT + 5, HCOL5 - (HCOL4 - 40), (mRow.bl_place_delivery != null ? mRow.bl_place_delivery.ToString() : ""), ifontName, ifontSize, "L", "", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16); ;


            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT + 5, (HCOL4 - 40) - HCOL1, "MARKS & NO.", ifontName, ifontSizesm, "LTB", "B", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL4 - 40, Row, ROW_HT + 5, (HCOL5 - 50) - (HCOL4 - 40), "PACKAGES", ifontName, ifontSizesm, "LTB", "B", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5 - 50, Row, ROW_HT + 5, HCOL6 - (HCOL5 - 50), "DESCRIPTION OF CARGO", ifontName, ifontSizesm, "TB", "B", R1, XL_COLF, 5, 16); ;
            AddXYLabel(HCOL6, Row, ROW_HT + 5, (HCOL6 + 100) - HCOL6, "GROSS WEIGHT", ifontName, ifontSizesm, "LTB", "B", R1, XL_COLG, 2, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT + 5, HCOL7 - (HCOL6 + 100), "MEASUREMENT", ifontName, ifontSizesm, "LRTB", "LB", R1, XL_COLJ, 1, 16, 0, 0, Xtolrnce);

            int vlineHgt = 0;
            vlineHgt = 23;
            Row += 5;

            Row += ROW_HT; R1++;
            DescStartRow = Row;
            DrawVLine(HCOL1, Row, ROW_HT * vlineHgt, "");
            DrawVLine(HCOL7, Row, ROW_HT * vlineHgt, "");

            sRow = Row;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark1 != null ? mRow.bl_mark1.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc1 != null ? mRow.bl_desc1.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark2 != null ? mRow.bl_mark2.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc2 != null ? mRow.bl_desc2.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark3 != null ? mRow.bl_mark3.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc3 != null ? mRow.bl_desc3.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, "GR.WT / KGS", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);
            str = "";
            if (Lib.Conv2Decimal(mRow.bl_cbm.ToString()) > 0)
                str = "CBM";
            AddXYLabel(HCOL6 + 100, sRow, ROW_HT, HCOL7 - (HCOL6 + 100), str, ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark4 != null ? mRow.bl_mark4.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc4 != null ? mRow.bl_desc4.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            str = "";
            if (Lib.Conv2Decimal(mRow.bl_grwt.ToString()) > 0)
                str = Lib.NumFormat(mRow.bl_grwt.ToString(), 3);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);
            str = "";
            if (Lib.Conv2Decimal(mRow.bl_cbm.ToString()) > 0)
                str = Lib.NumFormat(mRow.bl_cbm.ToString(), 3);
            AddXYLabel(HCOL6 + 100, sRow, ROW_HT, HCOL7 - (HCOL6 + 100), str, ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark5 != null ? mRow.bl_mark5.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc5 != null ? mRow.bl_desc5.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, "NET.WT / KGS", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);
            str = "";
            if (Lib.Conv2Decimal(mRow.bl_pcs.ToString()) > 0)
                str = "QTY / " + (mRow.bl_pcs_unit != null ? mRow.bl_pcs_unit.ToString() : "PCS");
            AddXYLabel(HCOL6 + 100, sRow, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark6 != null ? mRow.bl_mark6.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc6 != null ? mRow.bl_desc6.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            str = "";
            if (Lib.Conv2Decimal(mRow.bl_ntwt.ToString()) > 0)
                str = Lib.NumFormat(mRow.bl_ntwt.ToString(), 3);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - (HCOL6 + 100), str, ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);
            str = "";
            if (Lib.Conv2Decimal(mRow.bl_pcs.ToString()) > 0)
                str = Lib.NumFormat(mRow.bl_pcs.ToString(), 0);
            AddXYLabel(HCOL6 + 100, sRow, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark7 != null ? mRow.bl_mark7.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc7 != null ? mRow.bl_desc7.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_2desc7 != null ? mRow.bl_2desc7.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark8 != null ? mRow.bl_mark8.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc8 != null ? mRow.bl_desc8.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_2desc8 != null ? mRow.bl_2desc8.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark9 != null ? mRow.bl_mark9.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc9 != null ? mRow.bl_desc9.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_2desc9 != null ? mRow.bl_2desc9.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark10 != null ? mRow.bl_mark10.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc10 != null ? mRow.bl_desc10.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_2desc10 != null ? mRow.bl_2desc10.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark11 != null ? mRow.bl_mark11.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc11 != null ? mRow.bl_desc11.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_2desc11 != null ? mRow.bl_2desc11.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark12 != null ? mRow.bl_mark12.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc12 != null ? mRow.bl_desc12.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_2desc12 != null ? mRow.bl_2desc12.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark13 != null ? mRow.bl_mark13.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc13 != null ? mRow.bl_desc13.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_2desc13 != null ? mRow.bl_2desc13.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark14 != null ? mRow.bl_mark14.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc14 != null ? mRow.bl_desc14.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_2desc14 != null ? mRow.bl_2desc14.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);


            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark15 != null ? mRow.bl_mark15.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc15 != null ? mRow.bl_desc15.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark16 != null ? mRow.bl_mark16.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc16 != null ? mRow.bl_desc16.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark17 != null ? mRow.bl_mark17.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc17 != null ? mRow.bl_desc17.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark18 != null ? mRow.bl_mark18.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc18 != null ? mRow.bl_desc18.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark19 != null ? mRow.bl_mark19.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark20 != null ? mRow.bl_mark20.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark21 != null ? mRow.bl_mark21.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark22 != null ? mRow.bl_mark22.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark23 != null ? mRow.bl_mark23.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark24 != null ? mRow.bl_mark24.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);

            Row = DescStartRow + ROW_HT * 23;
            sRow = DescStartRow + ROW_HT * 25; sRow_HT = 13;

            str = "ALL TRANSACTIONS ARE SUBJECT TO THE COMPANY'S STANDARD";
            AddXYLabel(HCOL5 + 10, sRow, sRow_HT, HCOL7 - (HCOL5 + 23), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "TRADING CONDITIONS (COPIES AVAILABLE ON REQUEST FROM THE";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 1), sRow_HT, HCOL7 - (HCOL5 + 23), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "COMPANY) AND WHICH IN CERTAIN CASES, EXCLUSIVE ON LIMIT THE";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 2), sRow_HT, HCOL7 - (HCOL5 + 23), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "COMPANY'S LIABILITY AND INCLUDE CERTAIN INDEMNITIES BENEFITING";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 3), sRow_HT, HCOL7 - (HCOL5 + 23), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "THE COMPANY. CARGOMAR PRIVATE LIMITED IS ACTING AS AGENT IN ";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 4), sRow_HT, HCOL7 - (HCOL5 + 23), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "WITNESS WHEREOF(             ) FORWARDER'S CARGO RECEIPT(S). ALL";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 5), sRow_HT, HCOL7 - (HCOL5 + 23), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "OF THIS TENDER AND DATE HAVE BEEN SIGNED ONE OR WHICH BEING";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 6), sRow_HT, HCOL7 - (HCOL5 + 23), str, ifontName, ifontSizesm, "", "J", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);
            str = "ACCOMPLISHED. THE OTHERS TO STAND VOID.";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 7), sRow_HT, HCOL7 - (HCOL5 + 23), str, ifontName, ifontSizesm, "", "", R1, XL_COLF, 5, 16, 0, 0, 0, ifontSizesm + 1);

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "", ifontName, ifontSize, "LBR", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, Xtolrnce, ifontSize + 2);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "LB", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - (HCOL1), "ACCORDANCE WITH INSTRUCTIONS FROM THE BUYER ORIGINAL COPY I", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "HAVE RECEIVED THE FOLLOWING DOCUMENTS", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_fcr_doc1.ToString().Trim(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_fcr_doc2.ToString().Trim(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_fcr_doc3.ToString().Trim(), ifontName, ifontSize, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "FOR CARGOMAR PRIVATE LIMITED", ifontName, ifontSize, "LR", "BC", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "THE ABOVE DOCUMENTS PLUS BILLS OF LADING WILL BE DISPATCHED TO", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "CONSIGNEE OR OTHER DESIGNATED PARTIES", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "L", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "AUTHORISED SIGNATORY", ifontName, ifontSizesm, "LR", "C", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "TL", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, (HCOL4 - 40) - HCOL1, "FORWARDER'S CARGO RECEIPT NO.:", ifontName, ifontSizesm, "L", "B", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            str = "";
            if (mRow.hbl_fcr_no.ToString().Trim().Length > 0)
            {
                str = mRow.hbl_fcr_no.ToString().Trim();
                str += "  DT: " + GetFCRissuedDate(mRow.bl_issued_date_print);
            }

            AddXYLabel(HCOL4 - 45, Row, ROW_HT, HCOL5 - (HCOL4 - 50), str, ifontName, ifontSize, "", "", R1, XL_COLA, 4, 16, 0, 0, 0);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "BL", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LR", "", R1, XL_COLF, 5, 16);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSizesm, "LB", "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "LBR", "", R1, XL_COLF, 5, 16);
        }

         
        private int NewFontsize(int ColWidth, string ColStr, string ifname, int ifsize, out int pdfFsize)
        {
            int Newfsz = 0;
            pdfFsize = 0;

            //Label label1 = new Label();
            //label1.Width = ColWidth - 15;
            //label1.Text = ColStr;
            //label1.Font = new Font(ifname, ifsize);
            //while (label1.Width < System.Windows.Forms.TextRenderer.MeasureText(label1.Text,
            //    new Font(label1.Font.FontFamily, label1.Font.Size, label1.Font.Style)).Width)
            //{
            //    label1.Font = new Font(label1.Font.FontFamily, label1.Font.Size - 0.5f, label1.Font.Style);
            //    if (label1.Font.Size <= 6)
            //        break;
            //}

            //if (label1.Font.Size <= 6)
            //    label1.Font = new Font(ifname, 6);

            //Newfsz = (int)label1.Font.Size;
            //pdfFsize = (int)(Newfsz * 1.25);

            
            Newfsz = ifsize;
            pdfFsize = (int)(Newfsz * 1.25);

            return Newfsz;
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

        private void WriteBackSide()
        {
            Row = 250;
            ROW_HT = 18; ifontName = "Arial"; ifontSize = 10;

            HCOL1 = 120;
            HCOL2 = 650;

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT + 5, HCOL2 - HCOL1, "CARGO RECEIPT", ifontName, ifontSize + 4, "", "UCB", R1, XL_COLA, 4, 16, 0, 0, 0, ifontSize + 7);
            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "Received the goods described on the reverse side hereof in apparent good order", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "and condition except as noted, to be held and transported subject to the terms", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "and conditions contained in the regular form of Bill of Lading of the carrier, which", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "terms and conditions are incorporated herein and shall be considered as part", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "hereof with the same force and effect as if set fort herein in full. The goods are", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "received subject to delay or carriers's inability to carry due to accumulation of", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "goods, lack of conveyances, space or facilities of any sort, labour disputes,", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "strikes, lockouts, riots, war, government authority or any condition whatsoever", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "beyond the control of the carrier. Nothing in this cargo receipt shall operate to", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "limit or deprive the carrier of any statutory protection or of any exemption or", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "limitation of liability.", ifontName, ifontSize, "", "", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "This document is issued only to aid the shipper in seeking of the relevant letter of", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "credit. This document does not grant any title to the goods described on the", ifontName, ifontSize, "", "J", R1, XL_COLA, 4, 16, 0, 0, 0, 12);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "reverse side hereof.", ifontName, ifontSize, "", "", R1, XL_COLA, 4, 16, 0, 0, 0, 12);

            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "(CONTINUED ON REVERSE SIDE)", ifontName, ifontSize, "", "", R1, XL_COLA, 4, 16, 0, 0, 0, ifontSize + 3);
        }

        private string GetFormatLine(string sLine)
        {
            string[] sData = null;
            BsideStyle = "";
            sLine = sLine.Trim();
            sData = sLine.Split(sStyleSplit);
            if (sData.Length > 1)
            {
                BsideStyle = sData[0].Trim().Replace("2", dColr);
                sLine = sData[1].Trim();
            }

            return sLine;
        }

        private void FillAttachedData()
        {
            ifontSize = 9;
            ifontSizesm = 7;
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

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "FCR NO", ifontName, ifontSize, "", "B", R1, XL_COLA, 2, 16);
            AddXYLabel(HCOL2 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, mRow.hbl_fcr_no.ToString(), ifontName, ifontSize, "", "", R1, XL_COLD, 0, 16);
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "DATE", ifontName, ifontSize, "", "B", R1, XL_COLE, 0, 16);//BL DATE
            AddXYLabel(HCOL4 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");//":"
            //AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, Lib.DatetoStringDisplayformat(mRow.bl_issued_date), ifontName, ifontSize, "", "B", R1, XL_COLF, 5, 16);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4,  mRow.bl_issued_date_print, ifontName, ifontSize, "", "B", R1, XL_COLF, 5, 16);

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
        private string GetFCRissuedDate(object sDate)
        {
            string str = "";
            if (sDate != null)
            {
                str = sDate.ToString();
                if(str.Contains("-"))
                {
                    string[] sdata = str.Split('-');
                    int d = Lib.Conv2Integer(sdata[0]);
                    int m = Lib.Conv2Integer(sdata[1]);
                    int y= Lib.Conv2Integer(sdata[2]);
                    str = new DateTime(y, m, d).ToString("dd/MMM/yyyy").ToUpper();
                }
            }
            return str;
        }
        
    }
}
