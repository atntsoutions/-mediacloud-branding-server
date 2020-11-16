using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase.Connections;

namespace BLPim
{
    public class CampaignService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Campaign> mList = new List<Campaign>();
            Campaign mRow;

            string type = SearchData["type"].ToString();
            string comp_code = SearchData["comp_code"].ToString().ToUpper();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  a.rec_company_code ='" + comp_code + "' " ;
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.cam_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total from campaign a ";
                    sql += sWhere;
                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        page_rowcount = Lib.Conv2Integer(Dt_Temp.Rows[0]["total"].ToString());
                        page_count = Lib.Conv2Integer(Dt_Temp.Rows[0]["page_total"].ToString());
                    }
                    page_current = 1;
                }
                else
                {
                    if (type == "FIRST")
                        page_current = 1;
                    if (type == "PREV" && page_current > 1)
                        page_current--;
                    if (type == "NEXT" && page_current < page_count)
                        page_current++;
                    if (type == "LAST")
                        page_current = page_count;
                }

                startrow = (page_current - 1) * page_rows + 1;
                endrow = (startrow + page_rows) - 1;


                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select * from ( ";
                sql += "  select  cam_pkid, cam_slno, cam_name, tab_name,tab_campaign_table, row_number() over(order by tab_name) rn ";
                sql += "  from campaign a ";
                sql += "  left join tablesm b on a.cam_tab_id = tab_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by cam_name ";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Campaign();
                    mRow.cam_pkid = Dr["cam_pkid"].ToString();
                    mRow.cam_slno = Lib.Conv2Integer(Dr["cam_slno"].ToString());
                    mRow.cam_name = Dr["cam_name"].ToString();
                    mRow.cam_tab_name = Dr["tab_name"].ToString();


                    mRow.tab_campaign_table = Dr["tab_campaign_table"].ToString() == "Y" ? true : false;

                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);

            return RetData;
        }

      



        public Dictionary<string, object>  GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Campaign mRow =new Campaign();
            
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  cam_pkid,cam_slno, cam_name, cam_tab_id, tab_name,tab_table_name, tab_campaign_table, cam_product_name, cam_product_name_values, ";
                sql += " cam_store,cam_size,cam_aep,cam_output,cam_approver,cam_receiver,cam_logo,cam_image1,cam_image2,cam_image3,cam_image4,cam_image5,";
                sql += "cam_text1,cam_text2,cam_text3,cam_text4,cam_text5, ";
                sql += "cam_size_values,cam_aep_values,cam_output_values,";
                sql += "cam_text1_values,cam_text2_values,cam_text3_values,cam_text4_values,cam_text5_values";
                
                sql += " from campaign a  ";
                sql += " left join tablesm b on a.cam_tab_id = b.tab_pkid";
                sql += " where  a.cam_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Campaign();
                    mRow.cam_pkid = Dr["cam_pkid"].ToString();
                    mRow.cam_slno = Lib.Conv2Integer(Dr["cam_slno"].ToString());
                    mRow.cam_name = Dr["cam_name"].ToString();
                    mRow.cam_tab_id = Dr["cam_tab_id"].ToString();

                    mRow.cam_tab_name = Dr["tab_name"].ToString();
                    mRow.cam_table_name = Dr["tab_table_name"].ToString();

                    mRow.cam_store = Dr["cam_store"].ToString();

                    mRow.tab_campaign_table = Dr["tab_campaign_table"].ToString() == "Y" ? true : false;


                    mRow.cam_product_name = Dr["cam_product_name"].ToString();
                    mRow.cam_product_name_values = Dr["cam_product_name_values"].ToString();

                    mRow.cam_size = Dr["cam_size"].ToString();
                    mRow.cam_size_values = Dr["cam_size_values"].ToString();
                    mRow.cam_aep = Dr["cam_aep"].ToString();
                    mRow.cam_aep_values = Dr["cam_aep_values"].ToString();
                    mRow.cam_output = Dr["cam_output"].ToString();
                    mRow.cam_output_values = Dr["cam_output_values"].ToString();

                    mRow.cam_approver = Dr["cam_approver"].ToString();
                    mRow.cam_receiver = Dr["cam_receiver"].ToString();
                    mRow.cam_logo = Dr["cam_logo"].ToString();
                    mRow.cam_image1 = Dr["cam_image1"].ToString();
                    mRow.cam_image2 = Dr["cam_image2"].ToString();
                    mRow.cam_image3 = Dr["cam_image3"].ToString();
                    mRow.cam_image4 = Dr["cam_image4"].ToString();
                    mRow.cam_image5 = Dr["cam_image5"].ToString();

                    mRow.cam_text1 = Dr["cam_text1"].ToString();
                    mRow.cam_text1_values = Dr["cam_text1_values"].ToString();

                    mRow.cam_text2 = Dr["cam_text2"].ToString();
                    mRow.cam_text2_values = Dr["cam_text2_values"].ToString();

                    mRow.cam_text3 = Dr["cam_text3"].ToString();
                    mRow.cam_text3_values = Dr["cam_text3_values"].ToString();

                    mRow.cam_text4 = Dr["cam_text4"].ToString();
                    mRow.cam_text4_values = Dr["cam_text4_values"].ToString();

                    mRow.cam_text5 = Dr["cam_text5"].ToString();
                    mRow.cam_text5_values = Dr["cam_text5_values"].ToString();
                    break;
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            return RetData;
        }


        public string AllValid(Campaign Record)
        {
            string str = "";
            try
            {
                sql = "select cam_pkid from (";
                sql += "select cam_pkid  from campaign a where rec_company_code ='" + Record._globalvariables.comp_code + "' ";
                sql += " and (a.cam_name = '{NAME}')  ";
                sql += ") a where cam_pkid <> '{PKID}'";

                sql = sql.Replace("{NAME}", Record.cam_name);
                sql = sql.Replace("{PKID}", Record.cam_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Name Exists";

                string tname = Lib.getSettings(Record._globalvariables.comp_code, "IMAGE-TABLE", "NAME");
                if ( tname != Record.cam_tab_name)
                {
                    sql = "select cam_pkid from (";
                    sql += "select cam_pkid from campaign a where rec_company_code ='" + Record._globalvariables.comp_code + "' ";
                    sql += " and (a.cam_tab_id = '{ID}')  ";
                    sql += ") a where cam_pkid <> '{PKID}'";
                    sql = sql.Replace("{ID}", Record.cam_tab_id);
                    sql = sql.Replace("{PKID}", Record.cam_pkid);
                    if (Con_Oracle.IsRowExists(sql))
                        str = "Name Exists";
                }
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }



        public Dictionary<string, object> RunCampaign(Campaign Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Boolean retvalue = false;
            int iTot = 0;
            try
            {


                google_uploader g = new google_uploader();

                g.bSingle = false;

                g.user_is_admin = Record.user_is_admin;

                g.CreateColumn("doc_" + Record.cam_slno.ToString());

                g.comp_code = Record._globalvariables.comp_code;
                g.user_id = Record._globalvariables.user_pkid;
                string str = g.Process(Record.cam_table_name, "name");
                if (str != "")
                    g.UploadData("doc_" + Record.cam_slno.ToString());
                iTot = g.iTot;
                

            }
            catch (Exception Ex)
            {
                retvalue = false;
                throw Ex;
            }

            RetData.Add("total", iTot);
            RetData.Add("retvalue", retvalue);

            return RetData;
        }

        public Dictionary<string, object> ResetCampaign(Campaign Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Boolean retvalue = false;
            int iTot = 0;
            try
            {


                google_uploader g = new google_uploader();
                g.bSingle = false;
                g.user_is_admin = Record.user_is_admin;
                g.comp_code = Record._globalvariables.comp_code;
                g.user_id = Record._globalvariables.user_pkid;
                string str = g.Process(Record.cam_table_name, "name");
                if (str != "")
                    g.ResetCampaign("doc_" + Record.cam_slno.ToString());
                iTot = g.iTot;


            }
            catch (Exception Ex)
            {
                retvalue = false;
                throw Ex;
            }

            RetData.Add("total", iTot);
            RetData.Add("retvalue", retvalue);

            return RetData;
        }



        public Dictionary<string, object> Save(Campaign Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if (Record.cam_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();

                Rec.CreateRow("campaign", Record.rec_mode, "cam_pkid", Record.cam_pkid);
                Rec.InsertString("cam_name", Record.cam_name);
                Rec.InsertString("cam_tab_id", Record.cam_tab_id);
                Rec.InsertString("cam_store", Record.cam_store, "Z");
                Rec.InsertString("cam_product_name", Record.cam_product_name, "Z");
                Rec.InsertString("cam_product_name_values", Record.cam_product_name_values, "Z");
                Rec.InsertString("cam_size", Record.cam_size, "Z");
                Rec.InsertString("cam_size_values", Record.cam_size_values, "Z");
                Rec.InsertString("cam_aep", Record.cam_aep, "Z");
                Rec.InsertString("cam_aep_values", Record.cam_aep_values, "Z");
                Rec.InsertString("cam_output", Record.cam_output, "Z");
                Rec.InsertString("cam_output_values", Record.cam_output_values, "Z");
                Rec.InsertString("cam_approver", Record.cam_approver, "Z");
                Rec.InsertString("cam_receiver", Record.cam_receiver, "Z");
                Rec.InsertString("cam_logo", Record.cam_logo, "Z");
                Rec.InsertString("cam_image1", Record.cam_image1, "Z");
                Rec.InsertString("cam_image2", Record.cam_image2, "Z");
                Rec.InsertString("cam_image3", Record.cam_image3, "Z");
                Rec.InsertString("cam_image4", Record.cam_image4, "Z");
                Rec.InsertString("cam_image5", Record.cam_image5, "Z");
                Rec.InsertString("cam_text1", Record.cam_text1, "Z");
                Rec.InsertString("cam_text1_values", Record.cam_text1_values, "Z");
                Rec.InsertString("cam_text2", Record.cam_text2, "Z");
                Rec.InsertString("cam_text2_values", Record.cam_text2_values, "Z");
                Rec.InsertString("cam_text3", Record.cam_text3, "Z");
                Rec.InsertString("cam_text3_values", Record.cam_text3_values, "Z");
                Rec.InsertString("cam_text4", Record.cam_text4, "Z");
                Rec.InsertString("cam_text4_values", Record.cam_text4_values, "Z");
                Rec.InsertString("cam_text5", Record.cam_text5, "Z");
                Rec.InsertString("cam_text5_values", Record.cam_text5_values, "Z");

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_created_date", "GETDATE()");
                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_edited_date", "GETDATE()");
                }

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

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
            Con_Oracle.CloseConnection();
            return RetData;
        }

        public IDictionary<string, object> Delete(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();
            string pkid = SearchData["pkid"].ToString();

            RemoveColumn(pkid);

            try
            {
                Con_Oracle.BeginTransaction();
                sql = "delete from campaign where cam_pkid = '" + pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            return RetData;
        }


        private void RemoveColumn(string pkid)
        {
            sql = "select cam_slno from campaign where cam_pkid = '" + pkid + "'";
            object id = Con_Oracle.ExecuteScalar(sql);
            try
            {
                sql = "alter table pim_docm drop column doc_" + id.ToString();
                Con_Oracle = new DBConnection();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CloseConnection();
            }
            catch (Exception)
            {

            }
        }

       


    }
}
