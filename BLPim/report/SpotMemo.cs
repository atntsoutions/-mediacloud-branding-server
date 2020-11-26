using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using DataBase;
using DataBase.Connections;

namespace BLPim
{
    public class SpotMemo : BaseReport
    {
        public string slno = "";
        public string Report_Caption = "";

        public string pkid = "";


        public string emails_to = "";
        public string emails_cc = "";

        private DataTable dt_record1;
        private DataTable dt_record;
        private DataRow Dr = null;

        public string imagefolder = "";
        public string comp_code = "";
        public string user_code = "";

        private string fname = "";


        Single x1 = 0;
        Single y1 = 0;
        Single x2 = 0;
        Single y2 = 0;
        Single x3 = 0;
        Single y3 = 0;

        private int R1 = 0;
        private const int XL_COLA = 1;
        private const int XL_COLB = 2;
        private const int XL_COLC = 3;
        private const int XL_COLD = 4;
        private const int XL_COLE = 5;
        private const int XL_COLF = 6;
        private const int XL_COLG = 7;
        private const int XL_COL_TOT = 6;


        private int iWidth = 595;
        private int iHeight = 500;

        private string ImagePath = "";

        private string str = "";

        public SpotMemo()
        {

        }

        public void process()
        {
            ROW_HT = 18;
            ReadData();
            PrintData();
        }


        private void ReadData()
        {
            DataBase.Connections.DBConnection con = new DBConnection();
            try
            {

                DataTable Dt_Temp = new DataTable();

                string sql = "";

                sql = "";
                sql += " select  spot_pkid, spot_slno, spot_date, spot_store_id, spot_vendor_id, spot_recce_id, a.rec_created_by, ";
                sql += " store.comp_name as spot_store_name, store.comp_location as spot_store_location, store.comp_mobile as spot_store_mobile, ";
                sql += " store.comp_district as spot_store_district, store.comp_state as spot_store_state,";
                sql += " store.comp_logo_name as spot_store_logo_name, store.comp_image_name as spot_store_image_name, ";
                sql += " comp.comp_pkid, comp.comp_name as spot_comp_name, comp.comp_logo_name as spot_comp_logo_name, comp.comp_image_name as spot_comp_image_name, ";
                sql += " spot_executive_name, spot_store_contact_name, spot_store_contact_tel ";
                sql += " ";
                sql += " from pim_spotm a  ";
                sql += " left join companym store on a.spot_store_id = store.comp_pkid ";
                sql += " left join companym comp on store.comp_parent_id = comp.comp_pkid ";
                sql += " where spot_pkid = '" + pkid + "'";
                dt_record = con.ExecuteQuery(sql);

                if (dt_record.Rows.Count > 0)
                {
                    Dr = dt_record.Rows[0];
                    sql = "select comp_type,COMP_EMAIL from COMPANYM  where comp_pkid in('" + Dr["spot_store_id"].ToString() + "','" + Dr["spot_vendor_id"].ToString() + "')";
                    Dt_Temp = con.ExecuteQuery(sql);
                    foreach ( DataRow Dr  in Dt_Temp.Rows)
                    {
                        if ( Dr["comp_type"].ToString() == "V")
                        {
                            emails_cc = Lib.getEmail(emails_cc, Dr["comp_email"].ToString());
                        }
                    }

                    sql = "select a.USER_EMAIL as m1, b.USER_EMAIL as m2 from userm  a left join USERM b on a.user_parent_id = b.user_pkid where a.rec_company_code ='" + comp_code +  "' and a.USER_CODE in('" + Dr["rec_created_by"].ToString() + "')";
                    Dt_Temp = con.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_Temp.Rows)
                    {
                        emails_to = Lib.getEmail(emails_to, Dr["m1"].ToString());
                        emails_to = Lib.getEmail(emails_to, Dr["m2"].ToString());
                    }

                    sql = "select a.USER_EMAIL as m1 FROM  userm  a  where  a.rec_company_code ='" + comp_code + "' and  a.user_pkid = '" + Dr["spot_recce_id"].ToString() + "'";
                    Dt_Temp = con.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_Temp.Rows)
                    {
                        emails_cc = Lib.getEmail(emails_cc, Dr["m1"].ToString());
                    }
                }
                else
                    throw new Exception("No Record Found");

                sql = "";
                sql += " select  spotd_pkid, spotd_name ,spotd_slno, spotd_uom, spotd_wd, spotd_ht, ";
                sql += " spotd_artwork_id, artwork.param_name as spotd_artwork_name, artwork.param_slno as spotd_artwork_slno,artwork.param_file_name as spotd_artwork_file_name,";
                sql += " spotd_product_id, product.param_name as spotd_product_name, ";
                sql += " spotd_close_view, spotd_long_view, spotd_final_view ";
                sql += " from pim_spotd  a ";
                sql += " left join param artwork on a.spotd_artwork_id = artwork.param_pkid ";
                sql += " left join param product on a.spotd_product_id = product.param_pkid ";
                sql += " where spotd_parent_id = '" + pkid + "'";
                sql += " order by spotd_slno";
                dt_record1 = con.ExecuteQuery(sql);



                con.CloseConnection();
            } 
            catch ( Exception ex )
            {
                con.CloseConnection();
                throw ex;
            }
        }


        private void PrintData()
        {
            R1 = 1;

            addList("XLCOLUMN", "60", "10", "150", "180", "100", "10", "250");
            addList("SCALE", "79", "N");
            addList("MARGIN", "0.2", "0.2", "0.2", "0.2", "0.2", "0.2");


            //iHeight = 500; iWidth = 595;

            iHeight = 595; iWidth = 842;
            HCOL_MAX_WIDTH = 842;
            HCOL_MAX_HEIGHT = iHeight;


            BeginReport(iHeight, iWidth);

            AddPage(0, 0);

            Row = 20;
            ifontSize = 14;
            
            WriteDetail();

            
            foreach (DataRow Dr1 in dt_record1.Rows)
            {
                WriteSpotDetails(Dr1);
            }
            WriteFooter();
            
            EndReport();
        }

        private void WriteDetail()
        {

            HCOL1 = 100;
            HCOL2 = HCOL_MAX_WIDTH / 2;
            HCOL3 = HCOL_MAX_WIDTH - 100;

            Row = 10;
            slno = Dr["spot_slno"].ToString();

            R1++; Row += 0;
            AddXYLabel( HCOL_START , Row, ROW_HT * 2 , HCOL_MAX_WIDTH, "RECCE PROJECT FOR", "Arial", 20, "", "BC", "ORANGE");

            R1++; Row += ROW_HT;
            R1++; Row += ROW_HT;
            R1++; Row += ROW_HT;

            ImagePath = Lib.getUploadedPath(comp_code, Dr["comp_pkid"].ToString(), "logo\\" + Dr["spot_comp_logo_name"].ToString(), false);
            LoadImage("", HCOL2 - 50, Row, ImagePath, 100, 100);

            R1++; Row += (ROW_HT * 7);

            AddXYLabel( HCOL_START, Row, ROW_HT, HCOL_MAX_WIDTH, "AT", "Arial", ifontSize, "", "BC");
            
            R1++; Row += ROW_HT + 5;

            AddXYLabel( HCOL_START, Row, ROW_HT, HCOL_MAX_WIDTH, Dr["spot_store_name"].ToString() , "Arial", ifontSize + 6, "", "BC");
            R1++; Row += ROW_HT + 5;

            str = Dr["spot_store_location"].ToString();
            if (str.Length > 0 && Dr["spot_store_mobile"].ToString().Length > 0)
                str += ",";
            if ( Dr["spot_store_mobile"].ToString().Length > 0)
                str += "+" + Dr["spot_store_mobile"].ToString();

            AddXYLabel( HCOL_START, Row, ROW_HT, HCOL_MAX_WIDTH, str, "Arial", ifontSize -2, "", "BC");
            R1++; Row += ROW_HT;
            R1++; Row += ROW_HT;

            AddXYLabel( HCOL1 , Row, ROW_HT, HCOL2 -  HCOL1, Dr["spot_store_state"].ToString() , "Arial", ifontSize, "A", "C");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr["spot_store_district"].ToString(), "Arial", 12, "A", "C");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "RECCE DATE", "Arial", ifontSize, "A", "C");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr["spot_date"].ToString(), "Arial", 12, "A", "C");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "INTIMATED BY", "Arial", ifontSize , "A", "C");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr["spot_executive_name"].ToString(), "Arial", 12, "A", "C");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "PROJECT INFO", "Arial", ifontSize, "A", "C");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr["spot_slno"].ToString(), "Arial", 12, "A", "C");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "STORE CONTACT", "Arial", ifontSize , "A", "C");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr["spot_store_contact_name"].ToString(), "Arial", 12, "A", "C");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "TEL/MOBILE", "Arial", ifontSize , "A", "C");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr["spot_store_contact_tel"].ToString(), "Arial", 12, "A", "C");


            R1++; Row += ROW_HT;

            ImagePath = Lib.getUploadedPath(comp_code, "repo", "recce_pdf_footer1\\mediacloudlogo.png", false);
            SetFillRectangle(0, HCOL_MAX_HEIGHT - 35, 30, HCOL_MAX_WIDTH, 1, "GRAY", "GRAY");
            LoadImage("", HCOL_MAX_WIDTH - 130, HCOL_MAX_HEIGHT - 36, ImagePath, 32, 100);
            //polygon
            x1 = 0;
            y1 = HCOL_MAX_HEIGHT - 150;
            x2 = 0;
            y2 = HCOL_MAX_HEIGHT;
            x3 = 150;
            y3 = HCOL_MAX_HEIGHT;
            DrawPolygon("ORANGE", 3, x1.ToString() + "," + y1.ToString() + "," + x2.ToString() + "," + y2.ToString() + "," + x3.ToString() + "," + y3.ToString());

        }


        private void WriteSpotDetails(DataRow Dr1)
        {
            AddPage(0, 0);
            Row = 10;
            ifontSize = 12;

            HCOL1 = 5;
            HCOL2 = 150;
            HCOL3 = 430;
            HCOL4 = 435;

            R1++; Row += 0;
            AddXYLabel( HCOL_START, Row, ROW_HT * 2, HCOL_MAX_WIDTH, "SITE PHOTO", "Arial", 20, "", "BC", "ORANGE");

            R1++; Row += ROW_HT;
            R1++; Row += ROW_HT;
            //R1++; Row += ROW_HT;

            AddXYLabel( HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "INSIDE LEFT SIDE TOP", "Arial", ifontSize, "", "L");

            AddXYLabel( HCOL4, Row, ROW_HT, HCOL_MAX_WIDTH - HCOL4, "CLOSE VIEW", "Arial", ifontSize, "", "L");

            R1++; Row += ROW_HT;

            ImagePath = Lib.getUploadedPath(comp_code, Dr1["spotd_pkid"].ToString(), "closeview\\" + Dr1["spotd_close_view"].ToString(), false);
            LoadImage("", HCOL4 , Row, ImagePath, 230, HCOL_MAX_WIDTH - HCOL4);

            Row += 230;

            AddXYLabel( HCOL4, Row, ROW_HT, HCOL_MAX_WIDTH - HCOL4, "LONG VIEW", "Arial", ifontSize, "", "L");
            R1++; Row += ROW_HT;
            ImagePath = Lib.getUploadedPath(comp_code, Dr1["spotd_pkid"].ToString(), "longview\\" + Dr1["spotd_long_view"].ToString(), false);
            LoadImage("", HCOL4 , Row, ImagePath, 245, 120 );

            
            ifontSize = 10;

            R1++; Row += ROW_HT;
            R1++; Row += ROW_HT;
            R1++; Row += ROW_HT;
            R1++; Row += ROW_HT;
            R1++; Row += ROW_HT;

            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "RECCE PHOTO REF NO", "Arial", ifontSize, "A", "L");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr1["spotd_slno"].ToString(), "Arial", 12, "A", "L");

            R1++; Row += ROW_HT;
            str = "(W) " + Dr1["spotd_wd"].ToString() + " (H) " + Dr1["spotd_ht"].ToString();
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "SIZE IN " + Dr1["spotd_uom"].ToString(), "Arial", ifontSize, "A", "L");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, str, "Arial", 12, "A", "L");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "ART WORK", "Arial", ifontSize, "A", "L");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr1["spotd_artwork_name"].ToString(), "Arial", 12, "A", "L");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "PRODUCT", "Arial", ifontSize , "A", "L");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Dr1["spotd_product_name"].ToString(), "Arial", 12, "A", "L");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "CREATIVE NEEDED", "Arial", ifontSize , "A", "L");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", "Arial", 12, "A", "L");

            R1++; Row += ROW_HT;
            AddXYLabel( HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "REMARKS", "Arial", ifontSize, "A", "L");
            AddXYLabel( HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", "Arial", 12, "A", "L");


            R1++; Row += ROW_HT;



            ImagePath = Lib.getUploadedPath(comp_code, "repo", "recce_pdf_footer1\\mediacloudlogo.png", false);
            SetFillRectangle(0, HCOL_MAX_HEIGHT - 35, 30, HCOL_MAX_WIDTH, 1, "GRAY", "GRAY");
            LoadImage("", 30, HCOL_MAX_HEIGHT - 36, ImagePath, 32, 100);
            // Polygon
            x1 = HCOL_MAX_WIDTH;
            y1 = HCOL_MAX_HEIGHT - 150;
            x2 = HCOL_MAX_WIDTH;
            y2 = HCOL_MAX_HEIGHT;
            x3 = HCOL_MAX_WIDTH - 150;
            y3 = HCOL_MAX_HEIGHT;
            DrawPolygon("ORANGE", 3, x1.ToString() + "," + y1.ToString() + "," + x2.ToString() + "," + y2.ToString() + "," + x3.ToString() + "," + y3.ToString());



        }

        private void WriteFooter()
        {

            AddPage(0, 0);
            Row = 0;
            ifontSize = 12;

            HCOL1 = 5; HCOL2 = 150; HCOL3 = 330;

            
            AddXYLabel( HCOL_START, Row, HCOL_MAX_HEIGHT - 140 , HCOL_MAX_WIDTH, "THANK YOU", "Arial", 20, "", "BC", "ORANGE");
            AddXYLabel(HCOL_START, HCOL_MAX_HEIGHT - 30 , HCOL_MAX_HEIGHT - 30, HCOL_MAX_WIDTH, "", "Arial", 20, "", "BC", "ORANGE");

            Row = HCOL_MAX_HEIGHT - 120;

            AddXYLabel(HCOL_START, Row, ROW_HT, HCOL_MAX_WIDTH, "MEDIA CLOUD STUDIO PVT LTD", "Arial", 20, "", "BC");
            Row += ROW_HT + 5;
            AddXYLabel( HCOL_START, Row, ROW_HT, HCOL_MAX_WIDTH, "74/2315, River Landing Lane, Correya Road, Pachalam, Cochin, 682012, Kerala, India.", "Arial", 14, "", "BC");
            Row += ROW_HT;
            AddXYLabel( HCOL_START, Row, ROW_HT, HCOL_MAX_WIDTH, "www.modiacloud.studio | support@mediacloud,studio", "Arial", 14, "", "BC");
            Row += ROW_HT;
            AddXYLabel( HCOL_START, Row, ROW_HT, HCOL_MAX_WIDTH, "80860 40000 | 80860 90000", "Arial", 14, "", "BC");

        }


    }
}