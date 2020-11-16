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
    public class BLReport : BaseReport
    {
        public string dColr = "2";
        public Bl mRow = null;
        public string InvokeType = "";
        public Boolean Chk_BL_Original = false;
        public string RootPath = "";

        private int RowCount = 1;
        private int RowsPerPage = 60;
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

        public BLReport()
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

            //if (mRow.hbl_bl_no.ToString().Trim() == "" && InvokeType != "DRAFT")
            //    Lib.AddError(ref str, " MTD REG NO. not Found ");

            return str;
        }
        private void PrintData()
        {
            Row = 10;
            R1 = 0;
            addList("XLCOLUMN", "100", "90", "20", "100", "100", "150", "30", "20", "30", "20", "60");
            BeginReport(1100, 800);
            //if (Chk_BL_Original)
            //{
            //    WriteData();
            //    WriteData();
            //    WriteData();
            //}
            //else
                WriteData();
            EndReport();
        }
        private void WriteData()
        {
            Row = 10;

            AddPage(1100, 800);
            FillData();

            //if (Chk_BL_Original)
            //{
            //    CanWrite = true;
            //    AddPage(1100, 800);
            //    WriteBackSide();
            //    CanWrite = true;
            //}

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

            //LoadImage(RootPath + "\\wmLogo.png", 50, 180, 720, 720);//Water mark Image

            if (Chk_BL_Original == false)
            {
                //if (InvokeType == "DRAFT")
                //    str = "D R A F T";
                //else
                //    str = "Non-Negotiable";
                //AddXYLabel(370, 700, 0, 0, str, ifontName, 66, "", "W" + dColr, R1, XL_COLH, 3, 16, 0, 0, 0, 92); //water Mark

                //  AddXYLabel(750, 1350, 0, 0, "COPY", ifontName, 66, "", "W" + dColr, R1, XL_COLH, 3, 16, 0, 0, 0, 92); //water Mark diagonal
            }
            //if (Chk_BL_Original == false)
            //    HCOL1 = -165;
            //else
                HCOL1 = 20; // if(with out water mark)
            HCOL2 = HCOL1 + 120;
            HCOL3 = HCOL2 + 80;
            HCOL4 = HCOL3 + 30;
            HCOL5 = HCOL4 + 160;
            HCOL6 = HCOL5 + 188;
            HCOL7 = HCOL6 + 188;

            // Row = 20; if(with out water mark then row height is 20)
            //if (Chk_BL_Original == false)
            //    Row = -330;
            //else
                Row = 20;

            ROW_HT = 16; R1++;
            Row += ROW_HT + 5; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT + ROW_HT, HCOL7 - HCOL1, "ARRIVAL NOTICE", ifontName, 18, "LTR" + dColr, "BC" + dColr, R1, XL_COLA, 0, 20, 16, 0, 0, 22);//GlobalConstants.Address_Line1
            Row += ROW_HT; R1++;
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "Shipper", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSizesm, "TLR" + dColr, "", R1, XL_COLF, 5, 16);

            float sRow = Row + 16, sRow_HT = ROW_HT;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_issued_by1.ToString(), "Times New Roman", ifontSize + 8, "LR" + dColr, "BC" + dColr, R1, XL_COLF, 5, 16, 0, 0, 0, 20);
            //sRow += sRow_HT; sRow_HT = 14;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_issued_by2.ToString(), "Times New Roman", ifontSizesm + 1, "LR" + dColr, "C" + dColr, R1, XL_COLF, 5, 16, 0, 0, 0, 9);
            //sRow += sRow_HT;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_issued_by3.ToString(), "Times New Roman", ifontSizesm + 1, "LR" + dColr, "C" + dColr, R1, XL_COLF, 5, 16, 0, 0, 0, 9);
            //sRow += sRow_HT;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_issued_by4.ToString(), "Times New Roman", ifontSizesm + 1, "LR" + dColr, "C" + dColr, R1, XL_COLF, 5, 16, 0, 0, 0, 9);

            AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_issued_by1.ToString(), "Times New Roman", ifontSize + 4, "LR" + dColr, "BC" + dColr, R1, XL_COLF, 5, 16, 0, 0, 0, 20);
            sRow += sRow_HT; sRow_HT = ROW_HT;
            AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_issued_by2.ToString(), "Times New Roman", ifontSizesm + 1, "LR" + dColr, "C" + dColr, R1, XL_COLF, 5, 16, 0, 0, 0, 9);
            sRow += sRow_HT;
            AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_issued_by3.ToString(), "Times New Roman", ifontSizesm + 1, "LR" + dColr, "C" + dColr, R1, XL_COLF, 5, 16, 0, 0, 0, 9);
            sRow += sRow_HT;
            AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_issued_by4.ToString(), "Times New Roman", ifontSizesm + 1, "LR" + dColr, "C" + dColr, R1, XL_COLF, 5, 16, 0, 0, 0, 9);
            
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_name.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            
            
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_add1.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "AMS # :", ifontName, ifontSizesm, "TLR" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            //AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "ISF # :", ifontName, ifontSizesm, "TLR" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_add2.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "", ifontName, ifontSizesm, "LRB" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            //AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "", ifontName, ifontSizesm, "LRB" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);




            // LoadImage(RootPath + "\\Logo.gif", HCOL6 - 23, Row + 5, 53, 55);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_add3.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLF, 5, 16);
           

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_shipper_add4.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLF, 5, 16);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "MBL # :", ifontName, ifontSizesm, "TLR" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "HBL # :", ifontName, ifontSizesm, "TLR" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "Consignee", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, mRow.bl_mbl_no, ifontName, ifontSize, "LRB" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, mRow.bl_bl_no, ifontName, ifontSize, "LRB" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);
 
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_name.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "AMS # :", ifontName, ifontSizesm, "TLR" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "ISF # :", ifontName, ifontSizesm, "TLR" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            DrawVLine(HCOL5, Row, ROW_HT * 10, "" + dColr);
            DrawVLine(HCOL7, Row, ROW_HT * 10, "" + dColr);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_add1.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "", ifontName, ifontSizesm, "LRB" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "", ifontName, ifontSizesm, "LRB" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_add2.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "Vessel/Voyage", ifontName, ifontSizesm, "L" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "", ifontName, ifontSizesm, "R" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            str = mRow.bl_vsl_name.ToString() + " " + mRow.bl_vsl_voy_no.ToString();
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_add3.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, str, ifontName, ifontSize, "LB" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "", ifontName, ifontSizesm, "RB" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);
            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_consignee_add4.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6+100 - HCOL5, "Port of Loading", ifontName, ifontSizesm, "L" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6+100, Row, ROW_HT, HCOL7 - HCOL6, "ETD", ifontName, ifontSizesm, "R" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            sRow = Row + 2; sRow_HT = 10;

            /*str = "Taken in charge in apparently good condition herein at the place of receipt for transport and";
            AddXYLabel(HCOL5 + 10, sRow, sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "J" + dColr, R1, XL_COLF, 5, 16);
            str = "delivery as mentioned above, unless otherwise stated. The MTO in accordance with the";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 1), sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "J" + dColr, R1, XL_COLF, 5, 16);
            str = "provisions contained in the MTD undertakes to perform or to procure the performance of the";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 2), sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "J" + dColr, R1, XL_COLF, 5, 16);
            str = "multimodal tranport from the place at which the goods are taken in charge, to the place";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 3), sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "J" + dColr, R1, XL_COLF, 5, 16);
            str = "designated for delivery and assumes responsibility for such transport.";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 4), sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "" + dColr, R1, XL_COLF, 5, 16);
            str = "";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 5), sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "J" + dColr, R1, XL_COLF, 5, 16);
            str = "One of the MTD(s) must be surrendered, duly endorsed in exchange for the goods. in witness";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 6), sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "J" + dColr, R1, XL_COLF, 5, 16);
            str = "where of the original MTD all of this tenor and date have been signed in the number indicated";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 7), sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "J" + dColr, R1, XL_COLF, 5, 16);
            str = "below one of which being accomplished the other(s) to be void.";
            AddXYLabel(HCOL5 + 10, sRow + (sRow_HT * 8), sRow_HT, HCOL7 - (HCOL5 + 20), str, ifontName, 6, "", "" + dColr, R1, XL_COLF, 5, 16);
            */

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "", ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 + 100 - HCOL5, mRow.bl_pol.ToString(), ifontName, ifontSize, "LB" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - HCOL6 + 100, mRow.bl_pol_etd.ToString(), ifontName, ifontSize, "R" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "Notify Address", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLA, 4, 16, 0, 0, Xtolrnce);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 + 100 - HCOL5, "Port of Discharge :", ifontName, ifontSizesm, "L" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - HCOL6, "ETA", ifontName, ifontSizesm, "R" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_name.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 + 100 - HCOL5, mRow.bl_pod.ToString(), ifontName, ifontSize, "LB" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - (HCOL6+100), mRow.bl_pod_eta.ToString(), ifontName, ifontSize, "RB" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_add1.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 + 100 - HCOL5, "Place of Acceptance :", ifontName, ifontSizesm, "L" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - HCOL6 + 100, "", ifontName, ifontSizesm, "R" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_add2.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 + 100 - HCOL5, mRow.bl_place_receipt.ToString(), ifontName, ifontSize, "LB" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - (HCOL6 + 100), "", ifontName, ifontSizesm, "RB" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_add3.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 + 100 - HCOL5, "Place of Delivery :", ifontName, ifontSizesm, "L" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - (HCOL6 + 100), "", ifontName, ifontSizesm, "R" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, mRow.bl_notify_add4.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 4, 16, 0, 0, Xtolrnce, ifontSize + 2);
            AddXYLabel(HCOL5, Row, ROW_HT, (HCOL6 + 100) - HCOL5, mRow.bl_place_delivery.ToString(), ifontName, ifontSize, "LB" + dColr, "" + dColr, R1, XL_COLF, 0, 16, 0, 0, 0, 11);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - (HCOL6 + 100), "", ifontName, ifontSizesm, "RB" + dColr, "" + dColr, R1, XL_COLG, 0, 16, 0, 0, 0, 11);

            //sRow = Row - 1; sRow_HT = ROW_HT - 1;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_delivery_contact1.ToString(), ifontName, ifontSize, "", "", R1, XL_COLF, 5, 16, 0, 0, Xtolrnce, ifontSize + 2);//Delivery Contact1
            //sRow += sRow_HT;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_delivery_contact2.ToString(), ifontName, ifontSize, "", "", R1, XL_COLF, 5, 16, 0, 0, Xtolrnce, ifontSize + 2);//Delivery Contact2
            //sRow += sRow_HT;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_delivery_contact3.ToString(), ifontName, ifontSize, "", "", R1, XL_COLF, 5, 16, 0, 0, Xtolrnce, ifontSize + 2);//Delivery Contact3
            //sRow += sRow_HT;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_delivery_contact4.ToString(), ifontName, ifontSize, "", "", R1, XL_COLF, 5, 16, 0, 0, Xtolrnce, ifontSize + 2);//Delivery Contact4
            //sRow += sRow_HT;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_delivery_contact5.ToString(), ifontName, ifontSize, "", "", R1, XL_COLF, 5, 16, 0, 0, Xtolrnce, ifontSize + 2);//Delivery Contact5
            //sRow += sRow_HT;
            //AddXYLabel(HCOL5, sRow, sRow_HT, HCOL7 - HCOL5, mRow.bl_delivery_contact6.ToString(), ifontName, ifontSize, "", "", R1, XL_COLF, 5, 16, 0, 0, Xtolrnce, ifontSize + 2);//Delivery Contact5

            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT + 5, HCOL2 - HCOL1, "Place of Acceptance", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL2, Row, ROW_HT + 5, HCOL4 - HCOL2, "Date of Acceptance", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLD, 0, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL4, Row, ROW_HT + 5, HCOL5 - HCOL4, "Port of Loading", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLE, 0, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL5, Row, ROW_HT + 5, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLF, 5, 16);


            //Row += ROW_HT + 5; R1++;

            //int NewfszPdf = ifontSize + 2;
            //int Newfsz = NewFontsize((int)(HCOL2 - HCOL1), mRow.bl_place_receipt.ToString(), ifontName, ifontSize, out NewfszPdf);
            //AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, mRow.bl_place_receipt.ToString(), ifontName, Newfsz, "L" + dColr, "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce, NewfszPdf);
            ////AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, Lib.DatetoStringDisplayformat(mRow.bl_date_receipt), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, mRow.bl_date_receipt_print, ifontName, ifontSize, "L" + dColr, "", R1, XL_COLD, 0, 16, 0, 0, Xtolrnce, ifontSize + 2);

            //NewfszPdf = ifontSize + 2;
            //Newfsz = NewFontsize((int)(HCOL5 - HCOL4), mRow.bl_pol.ToString(), ifontName, ifontSize, out NewfszPdf);
            //AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, mRow.bl_pol.ToString(), ifontName, Newfsz, "L" + dColr, "", R1, XL_COLE, 0, 16, 0, 0, Xtolrnce, NewfszPdf);
            //AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLF, 5, 16);


            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT + 5, (HCOL3 - 25) - HCOL1, "Port of Discharge", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL3 - 25, Row, ROW_HT + 5, HCOL5 - (HCOL3 - 25), "Place of Delivery", ifontName, ifontSizesm, "TL" + dColr, "" + dColr, R1, XL_COLD, 1, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL5, Row, ROW_HT + 5, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLF, 5, 16);

            //Row += ROW_HT + 5; R1++;
            //NewfszPdf = ifontSize + 2;
            //Newfsz = NewFontsize((int)((HCOL3 - 25) - HCOL1), mRow.bl_pod.ToString(), ifontName, ifontSize, out NewfszPdf);
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL3 - 25) - HCOL1, mRow.bl_pod.ToString(), ifontName, Newfsz, "L" + dColr, "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce, NewfszPdf);

            //NewfszPdf = ifontSize + 2;
            //Newfsz = NewFontsize((int)(HCOL5 - (HCOL3 - 25)), mRow.bl_place_delivery.ToString(), ifontName, ifontSize, out NewfszPdf);
            //AddXYLabel(HCOL3 - 25, Row, ROW_HT, HCOL5 - (HCOL3 - 25), mRow.bl_place_delivery.ToString(), ifontName, Newfsz, "L" + dColr, "", R1, XL_COLD, 1, 16, 0, 0, Xtolrnce, NewfszPdf);
            //AddXYLabel(HCOL5, Row, ROW_HT, HCOL7 - HCOL5, "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLF, 5, 16); ;

            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT + 5, (HCOL4 + 25) - HCOL1, "Vessel Voyage No.", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL4 + 25, Row, ROW_HT + 5, HCOL5 - (HCOL4 + 25), "Date of Period of Delivery", ifontName, ifontSizesm, "TL" + dColr, "" + dColr, R1, XL_COLD, 1, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL5, Row, ROW_HT + 5, HCOL6 - HCOL5, "Modes / Means of Transport", ifontName, ifontSizesm, "TL" + dColr, "" + dColr, R1, XL_COLF, 1, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL6, Row, ROW_HT + 5, HCOL7 - HCOL6, "Route / Place of Transhipment (if any)", ifontName, ifontSizesm, "TLR" + dColr, "" + dColr, R1, XL_COLH, 3, 16, 0, 0, Xtolrnce);
            
            //Row += ROW_HT + 5; R1++;
            //str = mRow.bl_vsl_name.ToString() + " " + mRow.bl_vsl_voy_no.ToString();
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL4 + 25) - HCOL1, str, ifontName, ifontSize, "LB" + dColr, "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL4 + 25, Row, ROW_HT, HCOL5 - (HCOL4 + 25), mRow.bl_period_delivery.ToString(), ifontName, ifontSize, "LB" + dColr, "", R1, XL_COLD, 1, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, mRow.bl_move_type.ToString(), ifontName, ifontSize, "LB" + dColr, "", R1, XL_COLF, 1, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, mRow.bl_place_transhipment.ToString(), ifontName, ifontSize, "LBR" + dColr, "", R1, XL_COLH, 3, 16, 0, 0, Xtolrnce, ifontSize + 2);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "Container No(s)", ifontName, ifontSizesm, "LT" + dColr, "C" + dColr, R1, XL_COLA, 0, 16);
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, "Marks and Numbers", ifontName, ifontSizesm, "T", "C" + dColr, R1, XL_COLB, 1, 16);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL6 - HCOL4, "Number of packages, kinds of packages, general", ifontName, ifontSizesm, "T", "C" + dColr, R1, XL_COLD, 2, 16);
            AddXYLabel(HCOL6, Row, ROW_HT, (HCOL6 + 100) - HCOL6, "Gross Weight", ifontName, ifontSizesm, "", "TL" + dColr, R1, XL_COLG, 2, 16);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - (HCOL6 + 100), "Measurement", ifontName, ifontSizesm, "TR" + dColr, "L" + dColr, R1, XL_COLJ, 1, 16);

            Row += ROW_HT; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "", ifontName, ifontSizesm, "B" + dColr, "C");
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSizesm, "L" + dColr, "C", R1, XL_COLA, 0, 16);
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL4 - HCOL2, "", ifontName, ifontSizesm, "", "C", R1, XL_COLB, 1, 16);
            AddXYLabel(HCOL4, Row - 5, ROW_HT, HCOL6 - HCOL4, "description of goods (said to contain)", ifontName, ifontSizesm, "", "C" + dColr, R1, XL_COLD, 2, 16);
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL6 + 100 - HCOL6, "", ifontName, ifontSizesm, "", "C", R1, XL_COLG, 2, 16);
            AddXYLabel(HCOL6 + 100, Row, ROW_HT, HCOL7 - (HCOL6 + 100), "", ifontName, ifontSizesm, "R" + dColr, "C", R1, XL_COLJ, 1, 16);


            Row += ROW_HT; R1++;
            DescStartRow = Row;
            DrawVLine(HCOL1, Row, ROW_HT * 17, "" + dColr);
            DrawVLine(HCOL7, Row, ROW_HT * 17, "" + dColr);

            //sRow = Row + ROW_HT * 2;

            sRow = Row;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark1 != null ? mRow.bl_mark1.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize+1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc1 != null ? mRow.bl_desc1.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize +1);
            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark2 != null ? mRow.bl_mark2.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc2 != null ? mRow.bl_desc2.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);

            sRow = sRow + ROW_HT;
            AddXYLabel(HCOL1, sRow, ROW_HT, HCOL7 - HCOL1, mRow.bl_mark3 != null ? mRow.bl_mark3.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, 0, 16, 0, 0, Xtolrnce, ifontSize + 1);
            AddXYLabel(HCOL4, sRow, ROW_HT, HCOL7 - HCOL4, mRow.bl_desc3 != null ? mRow.bl_desc3.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLD, 2, 16, 0, 0, 0, ifontSize + 1);
           // AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, "GR.WT / KGS", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_grwt_caption != null ? mRow.bl_grwt_caption.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);
            //str = "";
            //if (Lib.Conv2Decimal(mRow.bl_cbm.ToString()) > 0)
            //    str = "CBM";
            str = "";
            if (Lib.Conv2Decimal(mRow.bl_cbm.ToString()) > 0)
                str = (mRow.bl_cbm_caption != null ? mRow.bl_cbm_caption.ToString() : "");
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
          //  AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, "NET.WT / KGS", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);
            AddXYLabel(HCOL6, sRow, ROW_HT, HCOL7 - HCOL6, mRow.bl_ntwt_caption != null ?mRow.bl_ntwt_caption.ToString() : "", ifontName, ifontSize, "", "", R1, XL_COLA, XL_COL_TOT, 16, 0, 0, 0, ifontSize + 2);
            //str = "";
            //if (Lib.Conv2Decimal(mRow.bl_pcs.ToString()) > 0)
            //    str = "QTY / " + mRow.bl_pcs_unit.ToString();
            str = "";
            if (Lib.Conv2Decimal(mRow.bl_pcs.ToString()) > 0)
                str =(mRow.bl_pcs_caption != null ? mRow.bl_pcs_caption.ToString() : "")+ " / " + mRow.bl_pcs_unit.ToString();
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

            Row = DescStartRow + ROW_HT * 17; R1++;
            DrawVLine(HCOL1, Row, ROW_HT * 7, "" + dColr);
            DrawVLine(HCOL7, Row, ROW_HT * 7, "" + dColr);
            if (mRow.bl_brazil_declaration)
            {
                sRow = (Row - 3) + ROW_HT * 5; sRow_HT = 10;
                str = "Notwithstanding any other provision contained in this Bill of Lading, the Merchant agrees that the Goods are considered formally delivered after effective";
                AddXYLabel(HCOL1 + 5, sRow, sRow_HT, HCOL7 - (HCOL1), str, ifontName, 7, "", "B0", R1, XL_COLA, XL_COL_TOT, 16);
                str = "discharge at a port in Brazil to the custody of Brazilian Customs at such port, and all liability of the Carrier whatsoever in connection with the goods (including,";
                AddXYLabel(HCOL1 + 5, sRow + (sRow_HT * 1), sRow_HT, HCOL7 - (HCOL1), str, ifontName, 7, "", "B0", R1, XL_COLA, XL_COL_TOT, 16);
                str = "without limitation, for misdelivery) shall cease at that time. Furthermore, the Merchant acknowledges and agrees that the Carrier shall not be responsible under";
                AddXYLabel(HCOL1 + 5, sRow + (sRow_HT * 2), sRow_HT, HCOL7 - (HCOL1), str, ifontName, 7, "", "B0", R1, XL_COLA, XL_COL_TOT, 16);
                str = "any circumstance for delivery of Goods without presentation of this original Bill of Lading in accordance with Brazilian Customs regulations";
                AddXYLabel(HCOL1 + 5, sRow + (sRow_HT * 3), sRow_HT, HCOL7 - (HCOL1), str, ifontName, 7, "", "B0", R1, XL_COLA, XL_COL_TOT, 16);
            }

            Row += ROW_HT * 7; R1++;
            AddXYLabel(HCOL1, Row, ROW_HT + 4, HCOL7 - HCOL1, "", ifontName, ifontSizesm, "LBR" + dColr, "C" + dColr, R1, XL_COLA, XL_COL_TOT, 16);
            //AddXYLabel(HCOL1, Row, ROW_HT + 4, HCOL7 - HCOL1, "Particulars above furnished by consignor / consignee", ifontName, ifontSizesm, "LBR" + dColr, "C" + dColr, R1, XL_COLA, XL_COL_TOT, 16);
            //if (InvokeType == "DRAFT")
            //    str = "";
            //else if (Chk_BL_Original == true)
            //    str = "ORIGINAL";
            //else
            //    str = "COPY";
            //str = "";
            //AddXYLabel(HCOL1, Row, ROW_HT + 4, HCOL7 - (HCOL1 + 10), str, "Times New Roman", ifontSize + 6, "", "BR" + dColr, R1, XL_COLA, 0, 16, 0, 0, 0, 20);

            //Row += ROW_HT + 4; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "Freight Amount", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL3, Row, ROW_HT, (HCOL5 + 45) - HCOL3, "Freight payable at", ifontName, ifontSizesm, "TL" + dColr, "" + dColr, R1, XL_COLD, 1, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL5 + 45, Row, ROW_HT, HCOL6 - (HCOL5 + 45), "", ifontName, ifontSizesm, "TL" + dColr, "" + dColr, R1, XL_COLF, 1, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "Place & date of issue", ifontName, ifontSizesm, "TLR" + dColr, "" + dColr, R1, XL_COLH, 3, 16, 0, 0, Xtolrnce);
            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "", ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 2, 16);
            //AddXYLabel(HCOL3, Row, ROW_HT, (HCOL5 + 45) - HCOL3, "", ifontName, ifontSize, "L" + dColr, "", R1, XL_COLD, 1, 16);
            //AddXYLabel(HCOL5 + 45, Row, ROW_HT, HCOL6 - (HCOL5 + 45), "", ifontName, ifontSizesm, "L" + dColr, "", R1, XL_COLF, 1, 16);
            //AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLH, 3, 16);
            //Row += ROW_HT; R1++;
            //str = "";
            //if (Lib.Conv2Decimal(mRow.bl_frt_amount.ToString()) != 0)
            //    str = Lib.NumFormat(mRow.bl_frt_amount.ToString(), 2);
            //AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, str, ifontName, ifontSize, "L" + dColr, "C", R1, XL_COLA, 2, 16, 0, 0, 0, ifontSize + 2);
            //AddXYLabel(HCOL3, Row, ROW_HT, (HCOL5 + 45) - HCOL3, mRow.bl_frt_pay_at.ToString(), ifontName, ifontSize, "L" + dColr, "C", R1, XL_COLD, 1, 16, 0, 0, 0, ifontSize + 2);


            //str = "";
            //AddXYLabel(HCOL5 + 45, Row, ROW_HT, HCOL6 - (HCOL5 + 45), str, ifontName, ifontSize, "L" + dColr, "C", R1, XL_COLF, 1, 16, 0, 0, 0, ifontSize + 2);
            ////str = mRow.bl_issued_place.ToString() + " " + Lib.DatetoStringDisplayformat(mRow.bl_issued_date);
            //str = mRow.bl_issued_place.ToString() + " " +  mRow.bl_issued_date_print;

            //AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "LR" + dColr, "C", R1, XL_COLH, 3, 16, 0, 0, 0, ifontSize + 2);

            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "", ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 2, 16);
            //AddXYLabel(HCOL3, Row, ROW_HT, (HCOL5 + 45) - HCOL3, "", ifontName, ifontSize, "L" + dColr, "", R1, XL_COLD, 1, 16);
            //AddXYLabel(HCOL5 + 45, Row, ROW_HT, HCOL6 - (HCOL5 + 45), "", ifontName, ifontSizesm, "L" + dColr, "", R1, XL_COLF, 1, 16);
            //AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLH, 3, 16);
            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL6 - 100) - HCOL1, "Other Particulars (if any)", ifontName, ifontSizesm, "LT" + dColr, "" + dColr, R1, XL_COLA, 2, 16, 0, 0, Xtolrnce);
            //AddXYLabel(HCOL6 - 100, Row, ROW_HT, HCOL7 - (HCOL6 - 100), "", "Times New Roman", ifontSize + 2, "TLR" + dColr, "BR" + dColr, R1, XL_COLH, 3, 16);
            //AddXYLabel(HCOL6 - 100, Row + 3, ROW_HT, HCOL7 - (HCOL6 - 90), "", "Times New Roman", ifontSize + 2, "", "BR" + dColr, R1, XL_COLH, 3, 16, 0, 0, 0, 12);
            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL6 - 100) - HCOL1, mRow.bl_remarks1.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL6 - 100, Row, ROW_HT, HCOL7 - (HCOL6 - 100), "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLH, 3, 16);
            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL6 - 100) - HCOL1, mRow.bl_remarks2.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL6 - 100, Row, ROW_HT, HCOL7 - (HCOL6 - 100), "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLH, 3, 16);
            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL6 - 100) - HCOL1, mRow.bl_remarks3.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL6 - 100, Row, ROW_HT, HCOL7 - (HCOL6 - 100), "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLH, 3, 16);
            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL6 - 100) - HCOL1, mRow.bl_remarks4.ToString(), ifontName, ifontSize, "L" + dColr, "", R1, XL_COLA, 2, 16, 0, 0, Xtolrnce, ifontSize + 2);
            //AddXYLabel(HCOL6 - 100, Row, ROW_HT, HCOL7 - (HCOL6 - 100), "", ifontName, ifontSize, "LR" + dColr, "", R1, XL_COLH, 3, 16);
            //Row += ROW_HT; R1++;
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL6 - 100) - HCOL1, "", ifontName, ifontSizesm, "LB" + dColr, "R", R1, XL_COLA, 2, 16);
            //AddXYLabel(HCOL1, Row, ROW_HT, (HCOL6 - 110) - HCOL1, "Weight and measurement of container not to be included", ifontName, ifontSizesm, "", "R" + dColr, R1, XL_COLA, 2, 16);
            //AddXYLabel(HCOL6 - 100, Row, ROW_HT, HCOL7 - (HCOL6 - 100), "", ifontName, ifontSizesm, "BLR" + dColr, "R", R1, XL_COLH, 3, 16);
            //AddXYLabel(HCOL6 - 100, Row, ROW_HT, HCOL7 - (HCOL6 - 90), "AUTHORISED SIGNATORY", ifontName, ifontSizesm, "", "R" + dColr, R1, XL_COLH, 3, 16);
            /*
            if (Chk_BL_Original == false)
            {
                if (InvokeType == "DRAFT")
                    str = "D R A F T";
                else
                    str = "Non-Negotiable";
                AddXYLabel(370, 700, 0, 0, str, ifontName, 66, "", "W" + dColr, R1, XL_COLH, 3, 16, 0, 0, 0, 92); //water Mark

                //  AddXYLabel(750, 1350, 0, 0, "COPY", ifontName, 66, "", "W" + dColr, R1, XL_COLH, 3, 16, 0, 0, 0, 92); //water Mark diagonal
            }
            */
        }

         
        private int NewFontsize(int ColWidth, string ColStr, string ifname, int ifsize, out int pdfFsize)
        {
            int Newfsz = 0;
            pdfFsize = 0;
            float fsize = ifsize;
            float StrWidth = Lib.GetWordWidth(ColStr, ifname, fsize, System.Drawing.FontStyle.Regular);
            ColWidth = ColWidth - 15;
            while (StrWidth > ColWidth)
            {
                fsize = fsize - 0.5f;
                StrWidth = Lib.GetWordWidth(ColStr, ifname, fsize, System.Drawing.FontStyle.Regular);
                if (fsize <= 6)
                    break;
            }
            Newfsz = (int)fsize;
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
            string Flname = "";
            string sline = null;
            float HCOL11 = 0;
            float HCOL12 = 0;
            
            HCOL1 = 20;
            HCOL2 = HCOL1 + 12;
            HCOL3 = HCOL2 + 12;
            HCOL4 = HCOL3 + 222;

            HCOL5 = HCOL4 + 10;

            HCOL6 = HCOL5 + 12;
            HCOL7 = HCOL6 + 12;
            HCOL8 = HCOL7 + 222;

            HCOL9 = HCOL8 + 10;

            HCOL10 = HCOL9 + 12;
            HCOL11 = HCOL10 + 12;
            HCOL12 = HCOL11 + 222;

            Row = 20;
            ROW_HT = 8; ifontName = "Arial"; ifontSize = 5;
            int bfontSize = 5;
            Flname = RootPath + "\\MTD-BACKSIDE1.TXT";
            if (!File.Exists(Flname))
                throw new Exception("BACK SIDE FILE NOT FOUND");

            StreamReader reader = new StreamReader(Flname);
            string[] ColLines = null;
            Boolean IsColumn1 = false;
            Boolean IsColumn2 = false;
            Boolean IsColumn3 = false;
            float Col1_StartRow = 0;
            while ((sline = reader.ReadLine()) != null)
            {
                sline = sline.Trim();

                if (sline == "{COLUMN 1}")
                {
                    ROW_HT = 7;
                    Col1_StartRow = Row;
                    IsColumn1 = true;
                    IsColumn2 = false;
                    IsColumn3 = false;
                }
                else if (sline == "{COLUMN 2}")
                {
                    ROW_HT = 7;
                    Row = Col1_StartRow;
                    IsColumn1 = false;
                    IsColumn2 = true;
                    IsColumn3 = false;
                }
                else if (sline == "{COLUMN 3}")
                {
                    ROW_HT = 7;
                    Row = Col1_StartRow;
                    IsColumn1 = false;
                    IsColumn2 = false;
                    IsColumn3 = true;
                }
                else if (sline == "{HLINE}")
                    DrawHLine(HCOL1, Row, HCOL8 - HCOL1);
                else
                {
                    ColLines = sline.Split(sColSplit);
                    if (IsColumn1)
                    {

                        if (ColLines.Length == 1)
                            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, GetFormatLine(sline), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
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
                    else if (IsColumn2)
                    {
                        if (ColLines.Length == 1)
                            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, GetFormatLine(sline), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
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
                    else if (IsColumn3)
                    {
                        if (ColLines.Length == 1)
                            AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, GetFormatLine(sline), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        else if (ColLines.Length == 2)
                        {
                            AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, GetFormatLine(ColLines[0]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL10, Row, ROW_HT, HCOL12 - HCOL10, GetFormatLine(ColLines[1]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        }
                        else if (ColLines.Length == 3)
                        {
                            AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, GetFormatLine(ColLines[0]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL10, Row, ROW_HT, HCOL11 - HCOL10, GetFormatLine(ColLines[1]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                            AddXYLabel(HCOL11, Row, ROW_HT, HCOL12 - HCOL11, GetFormatLine(ColLines[2]), ifontName, (BsideStyle.IndexOf("B") >= 0) ? bfontSize : ifontSize, "", BsideStyle);
                        }
                    }
                    else
                    {

                        if (ColLines.Length == 1)
                            AddXYLabel(HCOL1, Row, ROW_HT + 4, HCOL12 - HCOL1, GetFormatLine(sline), ifontName, (BsideStyle.IndexOf("B") >= 0) ? 8 : ifontSize, "", BsideStyle, 0, 0, 0, 16, 0, 0, 0, (BsideStyle.IndexOf("B") >= 0) ? 9 : ifontSize + 1);
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

            Row = 50;
            ROW_HT = 16;
            GetNextRow(); R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL5 - HCOL1, "ATTACHED SHEET", ifontName, ifontSize, "", "CB", R1, XL_COLA, 0, 20);

            GetNextRow();
            GetNextRow(); R1++;

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "BL NO", ifontName, ifontSize, "", "B", R1, XL_COLA, 2, 16);
            AddXYLabel(HCOL2 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, mRow.hbl_bl_no.ToString(), ifontName, ifontSize, "", "", R1, XL_COLD, 0, 16);

            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "BL DATE", ifontName, ifontSize, "", "B", R1, XL_COLE, 0, 16);
            AddXYLabel(HCOL4 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            //AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, Lib.DatetoStringDisplayformat(mRow.bl_issued_date), ifontName, ifontSize, "", "B", R1, XL_COLF, 5, 16);
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4,  mRow.bl_issued_date_print, ifontName, ifontSize, "", "B", R1, XL_COLF, 5, 16);

            GetNextRow(); R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "VESSEL", ifontName, ifontSize, "", "B", R1, XL_COLA, 2, 16);
            AddXYLabel(HCOL2 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL5 - HCOL2, mRow.bl_vsl_name.ToString(), ifontName, ifontSize, "", "", R1, XL_COLD, 0, 16);

            GetNextRow(); R1++;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "VOYAGE", ifontName, ifontSize, "", "B", R1, XL_COLA, 2, 16);
            AddXYLabel(HCOL2 - 10, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "B");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL5 - HCOL2, mRow.bl_vsl_voy_no.ToString(), ifontName, ifontSize, "", "", R1, XL_COLD, 0, 16);

            GetNextRow();  R1++;
            GetNextRow(); R1++;
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
                GetNextRow(); R1++;
            }
        }

        private void GetNextRow()
        {
            if (RowCount <= RowsPerPage)
            {
                RowCount++;
                Row += ROW_HT;
            }
            else
            {  
                GetNewPage();
                RowCount++;
                Row += ROW_HT;
            }
        }

        private void GetNewPage()
        {
            AddPage(1100, 800);
            ifontSize = 9;
            ifontSizesm = 7;
            ifontName = "Arial";
            ROW_HT = 16;
            Row = 50; //Start Row
            RowCount = 1; //Start RowIndex
        }
    }
}
