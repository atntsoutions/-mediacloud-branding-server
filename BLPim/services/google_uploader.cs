using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



using DataBase;
using DataBase.Connections;
using System.Net;

namespace BLPim
{
    public class google_uploader : BL_Base
    {

        public int iTot = 0;

        public string comp_code ="";
        public string user_id = "";

        public string user_is_admin = "";


        public Boolean bSingle = true;

        string table_id = "";
        string table_name = "";

        string google_url = "";

        string doc_id = "";

        string msql = "";

        string tab_store = "";

        DataTable Dt_TablesData = new DataTable();
        DataTable Dt_Campaign = new DataTable();

        string cols = "";
        string from = " ";
        string where = "";

        public string Process(string _table_id, string type = "id")
        {
            if (type == "id")
                table_id = _table_id;
            if (type == "name")
                table_name = _table_id;

            google_url = Lib.getSettings(comp_code, "GOOGLE-BS", "name");

            ReadData();
            if (Dt_Campaign.Rows.Count <= 0)
                return "";
            DataRow Dr = Dt_Campaign.Rows[0];
            buildQuery("store", Dr["cam_store"].ToString(), "");
            buildQuery("product", Dr["cam_product_name"].ToString(), Dr["cam_product_name_values"].ToString());
            buildQuery("size", Dr["cam_size"].ToString(), Dr["cam_size_values"].ToString());
            buildQuery("aep", Dr["cam_aep"].ToString(), Dr["cam_aep_values"].ToString());
            buildQuery("output", Dr["cam_output"].ToString(), Dr["cam_output_values"].ToString());
            buildQuery("approver", Dr["cam_approver"].ToString(), "");
            buildQuery("receiver", Dr["cam_receiver"].ToString(), "");
            buildQuery("logo", Dr["cam_logo"].ToString(), "");
            buildQuery("image1", Dr["cam_image1"].ToString(), "");
            buildQuery("image2", Dr["cam_image2"].ToString(), "");
            buildQuery("image3", Dr["cam_image3"].ToString(), "");
            buildQuery("image4", Dr["cam_image4"].ToString(), "");
            buildQuery("image5", Dr["cam_image5"].ToString(), "");
            buildQuery("text1", Dr["cam_text1"].ToString(), Dr["cam_text1_values"].ToString());
            buildQuery("text2", Dr["cam_text2"].ToString(), Dr["cam_text2_values"].ToString());
            buildQuery("text3", Dr["cam_text3"].ToString(), Dr["cam_text3_values"].ToString());
            buildQuery("text4", Dr["cam_text4"].ToString(), Dr["cam_text4_values"].ToString());
            buildQuery("text5", Dr["cam_text5"].ToString(), Dr["cam_text5_values"].ToString());
            msql = cols + from + where; ;
            return msql;
        }

        public void ReadData()
        {
            Con_Oracle = new DBConnection();
            sql = " select tab_pkid, tab_name, tab_table_name, tab_store, tabd_col_name, tabd_col_type";
            sql += " from tablesm a inner join tablesd b on a.tab_pkid = tabd_parent_id ";
            sql += " where ";
            if ( table_id != "")
                sql += " a.tab_pkid = '" + table_id + "'";
            if (table_name != "")
                sql += " a.tab_table_name = '" + table_name + "'";
            sql += " order by tabd_col_order ";
            Dt_TablesData = Con_Oracle.ExecuteQuery(sql);

            tab_store = "";
            if ( Dt_TablesData.Rows.Count > 0 )
            {
                tab_store = Dt_TablesData.Rows[0]["tab_store"].ToString();
            }
            

            sql = " select * from campaign where cam_tab_id = '" + Dt_TablesData.Rows[0]["tab_pkid"].ToString()  + "'";
            Dt_Campaign = Con_Oracle.ExecuteQuery(sql);

            Con_Oracle.CloseConnection();

            cols = "select doc_pkid, doc_slno, a.rec_company_code as doc_comp_code, a.doc_table_name";
            from = " from pim_docm a ";
            from += " left join {TABLE} x on a.doc_pkid = x.doc_parent_id";
            from += " left join companym store on a.doc_store_id = store.comp_pkid";

            if (bSingle == false && tab_store != ""  && user_is_admin == "N" ) 
            {
                from += " inner join userd e on e.rec_type = 'S' and a.doc_store_id = e.user_branch_id and e.user_id ='" + user_id + "'";
            }
            

            from = from.Replace("{TABLE}", Dt_TablesData.Rows[0]["tab_table_name"].ToString());

        }

        private DataRow getRow(string col_value)
        {
            DataRow DrRow = null;
            foreach (DataRow Dr in Dt_TablesData.Rows)
            {
                if (Dr["tabd_col_name"].ToString() == col_value)
                    DrRow = Dr;
            }
            return DrRow;
        }

        private void buildQuery(string col_name, string col_value, string col_default)
        {
            DataRow Dr;
            if (col_default != "")
                cols += ",'" + col_default + "' as " + col_name;
            else if (col_value == "STORE_NAME")
                cols += ",store.comp_name as " + col_name;
            else if (col_value == "STORE_APPROVER")
                cols += ",store.comp_approver_email as " + col_name;
            else if (col_value == "STORE_RECEIVER")
                cols += ",store.comp_receiver_email as " + col_name;
            else if (col_value == "LOGO_DEFAULT")
                cols += ",doc_file_name as " + col_name;
            else if (col_value == "PRODUCT")
                cols += ",doc_name as " + col_name;
            else if (col_value == "")
                cols += ",'' as " + col_name;
            else
            {
                if (col_value != "")
                {
                    Dr = getRow(col_value);
                    if (Dr["tabd_col_type"].ToString() == "TEXT" || Dr["tabd_col_type"].ToString() == "FILE")
                        cols += ",COL_" + Dr["tabd_col_name"].ToString() + "  as " + col_name;
                    if (Dr["tabd_col_type"].ToString() == "LIST")
                    {
                        cols += "," + Dr["tabd_col_name"].ToString() + ".param_name as " + col_name;
                        from += " left join param " + Dr["tabd_col_name"].ToString() + " on col_" + Dr["tabd_col_name"].ToString() + " = " + Dr["tabd_col_name"].ToString() + ".param_pkid ";
                    }
                }
            }

        }

        public void CreateColumn(string col)
        {
            try
            {
                sql = "alter table pim_docm add " + col + " char(1) ";
                Con_Oracle = new DBConnection();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                Con_Oracle.CloseConnection();
                
            }
        }


        public void ResetCampaign(string id)
        {
            try
            {
                sql = "";

                if (bSingle)
                    return;

                sql = msql + " where doc_table_name = '" + table_name + "' and isnull(" + id + ",'N') = 'Y' order by doc_slno";

                Con_Oracle = new DBConnection();
                DataTable Dt_Test = new DataTable();
                Dt_Test = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Test.Rows)
                {
                    sql = "update pim_docm set " + id + " = 'N' where doc_pkid = '" + Dr["doc_pkid"].ToString() + "'";

                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    this.iTot++;
                }
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                Con_Oracle.CloseConnection();
                throw Ex;

            }
        }



        public void UploadData(string id)
        {
            Boolean bTrans = false;
            try
            {

                doc_id = id;
                if (bSingle)
                    sql = msql + " where doc_pkid = '" + id + "'";
                else
                    sql = msql + " where doc_table_name = '" + table_name + "' and isnull(" + id + ",'N') = 'N' order by doc_slno";

                Con_Oracle = new DBConnection();
                DataTable Dt_Test = new DataTable();
                Dt_Test = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Test.Rows)
                {
                    UploadRecord(Dr);
                    if (!bSingle)
                    {
                        sql = "update pim_docm set " + id + " = 'Y' where doc_pkid = '" + Dr["doc_pkid"].ToString() + "'";

                        Con_Oracle.BeginTransaction();
                        bTrans = true;
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                        bTrans = false;

                        iTot++;
                    }
                }
            }
            catch ( Exception Ex)
            {
                if ( bTrans) 
                    Con_Oracle.RollbackTransaction();
                throw Ex;
            }

            Con_Oracle.CloseConnection();
        }

        public void UploadRecord(DataRow Dr)
        {
            string folder = "";

            StringBuilder sb = new StringBuilder();

            csv item = new csv();

            folder = Path.Combine(Dr["doc_comp_code"].ToString(), Dr["doc_table_name"].ToString(), Dr["doc_slno"].ToString()); 
            item.company = comp_code;
            item.orderno = Dr["doc_slno"].ToString();
            item.rowcount = "1";
            item.slno = "1";

            item.orderid = comp_code +  Dr["doc_slno"].ToString() + "-" + item.slno;
            if (Dr["size"].ToString().Length > 0)
                item.orderid = comp_code + Dr["doc_slno"].ToString() + "-" + item.slno + "-" + Dr["size"].ToString() ;

            item.customer_store = Dr["store"].ToString();
            item.approver = Dr["approver"].ToString();
            item.receiver = Dr["receiver"].ToString();

            
            if (item.approver != "")
                item.email = item.orderid + "!" + item.approver + "#" + item.receiver;
            else
                item.email = item.orderid + "#" + item.receiver;

            item.output = Dr["output"].ToString() + "\\" + item.email;

            item.project = Dr["aep"].ToString();

            if (Dr["size"].ToString().Length >0 )
                item.product_name = Dr["size"].ToString();
            else
                item.product_name = Dr["product"].ToString();

            item.upload_logo_location2 = Dr["logo"].ToString();
            if (Dr["logo"].ToString().Trim().Length > 0 )
                item.upload_logo_location2 = folder + "\\"+ Dr["logo"].ToString();

            item.upload_image1_location2 = Dr["image1"].ToString();
            if (Dr["image1"].ToString().Trim().Length > 0)
                item.upload_image1_location2 = folder + "\\" + Dr["image1"].ToString();

            item.upload_image2_location2 = Dr["image2"].ToString();
            if (Dr["image2"].ToString().Trim().Length > 0)
                item.upload_image2_location2 = folder + "\\" + Dr["image2"].ToString();

            item.upload_image3_location2 = Dr["image3"].ToString();
            if (Dr["image3"].ToString().Trim().Length > 0)
                item.upload_image3_location2 = folder + "\\" + Dr["image3"].ToString();

            item.upload_image4_location2 = Dr["image4"].ToString();
            if (Dr["image4"].ToString().Trim().Length > 0)
                item.upload_image4_location2 = folder + "\\" + Dr["image4"].ToString();

            item.upload_image5_location2 = Dr["image5"].ToString();
            if (Dr["image5"].ToString().Trim().Length > 0)
                item.upload_image5_location2 = folder + "\\" + Dr["image5"].ToString();

            item.text1 = Dr["text1"].ToString();
            item.text2 = Dr["text2"].ToString();
            item.text3 = Dr["text3"].ToString();
            item.text4 = Dr["text4"].ToString();
            item.text5 = Dr["text5"].ToString();


            string json = System.Text.Json.JsonSerializer.Serialize(item);

            postdata(json, Dr["doc_slno"].ToString(), "1");

        }

        private void postdata(string json, string orderno, string slno)
        {
            try
            {

                var cli = new WebClient();
                cli.Headers[HttpRequestHeader.ContentType] = "application/json";

                

                //joycok
                //google_url = "https://script.google.com/macros/s/AKfycbzSbUWJ5KSbeuS-Q8hx9Mjs6Pi902kt1cZQSBQRZF109Rmyk0_O/exec";

                // zissvideo
                //google_url ="";
                //mediacloud - austin
                //google_url = "https://script.google.com/macros/s/AKfycbykEG3mc4fGeQfpW_phRXrfLSa9H4dgujIP6sk5pp0LWOBjlPE/exec";
                

                string str = cli.UploadString(new Uri(google_url), "POST", json);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



    }
}
