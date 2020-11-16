using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;
using System.Drawing;
using BLOperations.models;

namespace BLOperations
{
    public class HouseListService : BL_Base
    {
        private string File_Name = "";
        private string File_Type = "";
        private string File_Display_Name = "myreport.pdf";
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Bl> mList = new List<Bl>();
            Bl mRow;
         

            string sWhere = "";
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string company_code = SearchData["company_code"].ToString();

            string branch_code = "";
            //string branch_code = SearchData["branch_code"].ToString();

            string type = SearchData["type"].ToString();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            string masterno = "";
            if (SearchData.ContainsKey("masterno"))
                masterno = SearchData["masterno"].ToString();
            string houseno = "";
            if (SearchData.ContainsKey("houseno"))
                houseno = SearchData["houseno"].ToString();

            try
            {
                sWhere = "";
                sWhere = " where a.bl_type ='HBL' and a.rec_company_code = '{COMPCODE}'  ";
                if (masterno.Trim() != "")
                    sWhere += " and a.bl_mbl_no like '%" + masterno + "%'";
                if (houseno.Trim() != "")
                    sWhere += " and a.bl_bl_no like '%" + houseno + "%'";

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM  bl a ";
                    sql += sWhere;

                    sql = sql.Replace("{COMPCODE}", company_code);
                   
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
                sql += " select bl_pkid,bl_shipper_name,bl_consignee_name,bl_notify_name,";
                sql += " bl_place_receipt,bl_pol,bl_pod,bl_place_delivery,bl_bl_no,";
                sql += " bl_vsl_name,bl_vsl_voy_no,bl_issued_place,bl_issued_date,bl_agent_name,";
                sql += " bl_carrier_name,bl_mbl_no,bl_pol_etd,bl_pod_eta,";
                sql += " row_number() over(order by bl_bl_no ) rn";
                sql += " from bl a ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by bl_pod_eta desc, bl_bl_no ";
 
               
                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());
               
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Bl();
         
                    mRow.bl_pkid = Dr["bl_pkid"].ToString();
                    mRow.bl_shipper_name = Dr["bl_shipper_name"].ToString();
                    mRow.bl_consignee_name = Dr["bl_consignee_name"].ToString();
                    mRow.bl_notify_name = Dr["bl_notify_name"].ToString();
                    mRow.bl_place_receipt = Dr["bl_place_receipt"].ToString();
                    mRow.bl_pol = Dr["bl_pol"].ToString();
                    mRow.bl_pod = Dr["bl_pod"].ToString();
                    mRow.bl_place_delivery = Dr["bl_place_delivery"].ToString();
                    mRow.bl_bl_no = Dr["bl_bl_no"].ToString();
                    mRow.bl_vsl_name = Dr["bl_vsl_name"].ToString();
                    mRow.bl_vsl_voy_no = Dr["bl_vsl_voy_no"].ToString();
                    mRow.bl_issued_place = Dr["bl_issued_place"].ToString();
                    mRow.bl_issued_date = Lib.DatetoString(Dr["bl_issued_date"]);
                    mRow.bl_agent_name = Dr["bl_agent_name"].ToString();
                    mRow.bl_carrier_name = Dr["bl_carrier_name"].ToString();
                    mRow.bl_mbl_no = Dr["bl_mbl_no"].ToString();
                    mRow.bl_pol_etd = Lib.DatetoString(Dr["bl_pol_etd"]);
                    mRow.bl_pod_eta = Lib.DatetoString(Dr["bl_pod_eta"]);


                    mList.Add(mRow);
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                    {
                        //PrintOrderList(mList, branch_code, file_pkid);

                    }
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("list", mList);
            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rows", page_rows);
            RetData.Add("page_rowcount", page_rowcount);

            //RetData.Add("filename", File_Name);
            //RetData.Add("filetype", File_Type);
            //RetData.Add("filedisplayname", File_Display_Name);
            return RetData;
        }


        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string id = SearchData["pkid"].ToString();
            string report_folder = "";
            string folderid = "";
            string type ="";
            string comp_code = "";
            if (SearchData.ContainsKey("type"))
                type = SearchData["type"].ToString();
            if (SearchData.ContainsKey("report_folder"))
                report_folder = SearchData["report_folder"].ToString();
            if (SearchData.ContainsKey("folderid"))
                folderid = SearchData["folderid"].ToString();
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            Bl mRow = new Bl();
            bool bOk = false;
            int Ctr = 0;

            try
            {

                Con_Oracle = new DBConnection();

                sql = "select  bl_pkid,bl_shipper_id,bl_shipper_br_id,shpr.cust_code as bl_shipper_code,shpraddr.add_branch_slno as  bl_shipper_br_no,bl_shipper_name ";
                sql += " ,bl_shipper_add1 ,bl_shipper_add2,bl_shipper_add3,bl_shipper_add4  ";
                sql += " ,bl_consignee_id,cngeaddr.add_branch_slno as bl_consignee_br_no,bl_consignee_br_id,cnge.cust_code as bl_consignee_code,bl_consignee_name,bl_consignee_add1,bl_consignee_add2,bl_consignee_add3,bl_consignee_add4 ";
                sql += " ,bl_issued_by1,bl_issued_by2,bl_issued_by3,bl_issued_by4,bl_issued_by5 ";
                sql += " ,bl_notify_id,nfyaddr.add_branch_slno as bl_notify_br_no ,bl_notify_br_id,nfy.cust_code as  bl_notify_code,bl_notify_name,bl_notify_add1,bl_notify_add2,bl_notify_add3,bl_notify_add4    ";

                sql += " ,bl_place_receipt,bl_date_receipt,bl_pol_id, bl_pol,pol.param_code as bl_pol_code ,bl_pod_id,bl_pod,pod.param_code as bl_pod_code,bl_place_delivery";
                sql += " ,bl_delivery_contact1,bl_delivery_contact2,bl_delivery_contact3,bl_delivery_contact4,bl_delivery_contact5    ";
                sql += " ,bl_delivery_contact6,bl_reg_no,bl_fcr_doc1,bl_fcr_doc2,bl_fcr_doc3,bl_vsl_name    ";
                sql += " ,bl_vsl_voy_no,bl_period_delivery,bl_move_type,bl_place_transhipment  ";
                sql += " ,bl_grwt,bl_cbm,bl_ntwt,bl_pcs,bl_pcs_unit ";
                sql += " ,bl_frt_amount,bl_frt_pay_at,bl_issued_place,bl_issued_date,bl_no_copies";
                sql += " ,bl_remarks1, bl_remarks2, bl_remarks3, bl_remarks4,bl_is_original ";
                sql += " ,'' as hbl_blno_generated,bl_fcr_no as hbl_fcr_no,bl_bl_no as hbl_bl_no,bl_bl_date as hbl_date,'' as hbl_seq_format_id,bl_brazil_declaration";
                sql += " ,bl_print_format_id,bl_print,bl_original_print,a.rec_category,bl_iata_carrier ";
                sql += " ,bl_itm_desc,bl_itm_po ";
                sql += " ,agnt.cust_code as bl_agent_code,agntaddr.add_branch_slno as  bl_agent_br_no,bl_agent_name as bl_delivery_contact1 ";
                sql += " ,agntaddr.add_line1 as bl_delivery_contact2,agntaddr.add_line2 as bl_delivery_contact3,agntaddr.add_line3 as bl_delivery_contact4,agntaddr.add_line4 as bl_delivery_contact5 ";
                sql += " ,a.bl_mbl_no,a.bl_pol_etd,a.bl_pod_eta ";
                sql += " from bl a  ";
               // sql += " left join hblm b on a.bl_pkid = b.hbl_pkid ";
                sql += " left join customerm shpr on a.bl_shipper_id = shpr.cust_pkid ";
                sql += " left join addressm shpraddr on a.bl_shipper_br_id = shpraddr.add_pkid ";
                sql += " left join customerm cnge on a.bl_consignee_id = cnge.cust_pkid ";
                sql += " left join addressm cngeaddr on a.bl_consignee_br_id = cngeaddr.add_pkid ";
                sql += " left join customerm nfy on a.bl_notify_id = nfy.cust_pkid ";
                sql += " left join addressm nfyaddr on a.bl_notify_br_id = nfyaddr.add_pkid ";
                sql += " left join param pol on a.bl_pol_id = pol.param_pkid ";
                sql += " left join param pod on a.bl_pod_id = pod.param_pkid ";
                sql += " left join customerm agnt on a.bl_agent_id = agnt.cust_pkid ";
                sql += " left join addressm agntaddr on a.bl_agent_id = agntaddr.add_pkid ";

                sql += " where  a.bl_pkid ='" + id + "'";

                DataTable Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    bOk = true;
                    mRow = new Bl();
                    mRow.bl_pkid = Dr["bl_pkid"].ToString();
                    mRow.bl_shipper_id = Dr["bl_shipper_id"].ToString();
                    mRow.bl_shipper_br_id = Dr["bl_shipper_br_id"].ToString();
                    mRow.bl_shipper_br_no = Dr["bl_shipper_br_no"].ToString();
                    mRow.bl_shipper_code = Dr["bl_shipper_code"].ToString();
                    mRow.bl_shipper_name = Dr["bl_shipper_name"].ToString();
                    mRow.bl_shipper_add1 = Dr["bl_shipper_add1"].ToString();
                    mRow.bl_shipper_add2 = Dr["bl_shipper_add2"].ToString();
                    mRow.bl_shipper_add3 = Dr["bl_shipper_add3"].ToString();
                    mRow.bl_shipper_add4 = Dr["bl_shipper_add4"].ToString();

                    mRow.bl_consignee_id = Dr["bl_consignee_id"].ToString();
                    mRow.bl_consignee_br_id = Dr["bl_consignee_br_id"].ToString();
                    mRow.bl_consignee_br_no = Dr["bl_consignee_br_no"].ToString();
                    mRow.bl_consignee_code = Dr["bl_consignee_code"].ToString();
                    mRow.bl_consignee_name = Dr["bl_consignee_name"].ToString();
                    mRow.bl_consignee_add1 = Dr["bl_consignee_add1"].ToString();
                    mRow.bl_consignee_add2 = Dr["bl_consignee_add2"].ToString();
                    mRow.bl_consignee_add3 = Dr["bl_consignee_add3"].ToString();
                    mRow.bl_consignee_add4 = Dr["bl_consignee_add4"].ToString();
                    mRow.bl_issued_by1 = Dr["bl_issued_by1"].ToString();
                    mRow.bl_issued_by2 = Dr["bl_issued_by2"].ToString();
                    mRow.bl_issued_by3 = Dr["bl_issued_by3"].ToString();
                    mRow.bl_issued_by4 = Dr["bl_issued_by4"].ToString();
                    mRow.bl_issued_by5 = Dr["bl_issued_by5"].ToString();
                    mRow.bl_notify_id = Dr["bl_notify_id"].ToString();
                    mRow.bl_notify_br_id = Dr["bl_notify_br_id"].ToString();
                    mRow.bl_notify_br_no = Dr["bl_notify_br_no"].ToString();
                    mRow.bl_notify_code = Dr["bl_notify_code"].ToString();
                    mRow.bl_notify_name = Dr["bl_notify_name"].ToString();
                    mRow.bl_notify_add1 = Dr["bl_notify_add1"].ToString();
                    mRow.bl_notify_add2 = Dr["bl_notify_add2"].ToString();
                    mRow.bl_notify_add3 = Dr["bl_notify_add3"].ToString();
                    mRow.bl_notify_add4 = Dr["bl_notify_add4"].ToString();

                    mRow.bl_place_receipt = Dr["bl_place_receipt"].ToString();
                    mRow.bl_date_receipt = Lib.DatetoString(Dr["bl_date_receipt"]);
                    mRow.bl_date_receipt_print = Lib.DatetoStringDisplayformat(Dr["bl_date_receipt"]);
                    mRow.bl_pol_id = Dr["bl_pol_id"].ToString();
                    mRow.bl_pol = Dr["bl_pol"].ToString();
                    mRow.bl_pol_code = Dr["bl_pol_code"].ToString();
                    mRow.bl_pod_id = Dr["bl_pod_id"].ToString();
                    mRow.bl_pod = Dr["bl_pod"].ToString();
                    mRow.bl_pod_code = Dr["bl_pod_code"].ToString();
                    mRow.bl_place_delivery = Dr["bl_place_delivery"].ToString();
                    mRow.bl_delivery_contact1 = Dr["bl_delivery_contact1"].ToString();
                    mRow.bl_delivery_contact2 = Dr["bl_delivery_contact2"].ToString();
                    mRow.bl_delivery_contact3 = Dr["bl_delivery_contact3"].ToString();
                    mRow.bl_delivery_contact4 = Dr["bl_delivery_contact4"].ToString();
                    mRow.bl_delivery_contact5 = Dr["bl_delivery_contact5"].ToString();
                    mRow.bl_delivery_contact6 = Dr["bl_delivery_contact6"].ToString();
                    mRow.bl_reg_no = Dr["bl_reg_no"].ToString();

                    mRow.bl_fcr_doc1 = Dr["bl_fcr_doc1"].ToString();
                    mRow.bl_fcr_doc2 = Dr["bl_fcr_doc2"].ToString();
                    mRow.bl_fcr_doc3 = Dr["bl_fcr_doc3"].ToString();
                    mRow.bl_vsl_name = Dr["bl_vsl_name"].ToString();
                    mRow.bl_vsl_voy_no = Dr["bl_vsl_voy_no"].ToString();
                    mRow.bl_period_delivery = Dr["bl_period_delivery"].ToString();
                    mRow.bl_move_type = Dr["bl_move_type"].ToString();
                    mRow.bl_place_transhipment = Dr["bl_place_transhipment"].ToString();

                    mRow.bl_grwt_caption = "";
                    mRow.bl_ntwt_caption = "";
                    mRow.bl_cbm_caption = "";
                    mRow.bl_pcs_caption = "";
                    mRow.bl_pcs_unit_caption = "";

                    mRow.bl_grwt = Lib.Conv2Decimal(Dr["bl_grwt"].ToString());
                    mRow.bl_cbm = Lib.Conv2Decimal(Dr["bl_cbm"].ToString());
                    mRow.bl_ntwt = Lib.Conv2Decimal(Dr["bl_ntwt"].ToString());
                    mRow.bl_pcs = Lib.Conv2Decimal(Dr["bl_pcs"].ToString());
                    mRow.bl_pcs_unit = Dr["bl_pcs_unit"].ToString();

                    mRow.bl_frt_amount = Lib.Conv2Decimal(Dr["bl_frt_amount"].ToString());
                    mRow.bl_frt_pay_at = Dr["bl_frt_pay_at"].ToString();
                    mRow.bl_issued_place = Dr["bl_issued_place"].ToString();
                    mRow.bl_issued_date = Lib.DatetoString(Dr["bl_issued_date"]);
                    mRow.bl_issued_date_print = Lib.DatetoStringDisplayformat(Dr["bl_issued_date"]);
                    mRow.bl_no_copies = Lib.Conv2Integer(Dr["bl_no_copies"].ToString());
                    mRow.bl_remarks1 = Dr["bl_remarks1"].ToString();
                    mRow.bl_remarks2 = Dr["bl_remarks2"].ToString();
                    mRow.bl_remarks3 = Dr["bl_remarks3"].ToString();
                    mRow.bl_remarks4 = Dr["bl_remarks4"].ToString();

                    mRow.bl_is_original = (Dr["bl_is_original"].ToString() == "Y" ? true : false);
                    mRow.hbl_bl_no = Dr["hbl_bl_no"].ToString();
                    mRow.hbl_fcr_no = Dr["hbl_fcr_no"].ToString();
                    mRow.hbl_blno_generated = Dr["hbl_blno_generated"].ToString();
                    mRow.hbl_seq_format_id = Dr["hbl_seq_format_id"].ToString();
                    mRow.bl_brazil_declaration = (Dr["bl_brazil_declaration"].ToString() == "Y" ? true : false);
                    mRow.bl_print = Dr["bl_print"].ToString();
                    mRow.bl_original_print = Dr["bl_original_print"].ToString();
                    mRow.bl_print_format_id = Dr["bl_print_format_id"].ToString();
                    mRow.rec_category = Dr["rec_category"].ToString();
                    mRow.bl_iata_carrier = Dr["bl_iata_carrier"].ToString();
                    mRow.hbl_date = Lib.DatetoString(Dr["hbl_date"]);
                    mRow.hbl_date = mRow.bl_issued_date;
                    mRow.bl_itm_po = Dr["bl_itm_po"].ToString();
                    mRow.bl_itm_desc = Dr["bl_itm_desc"].ToString();
                    mRow.bl_bl_no = Dr["hbl_bl_no"].ToString();
                    mRow.bl_mbl_no = Dr["bl_mbl_no"].ToString();
                    mRow.bl_pol_etd = Lib.DatetoStringDisplayformat(Dr["bl_pol_etd"]);
                    mRow.bl_pod_eta = Lib.DatetoStringDisplayformat(Dr["bl_pod_eta"]);
                    // mRow.bl_frt_status = getseafrtstatus(id);
                    break;
                }

                if (bOk)
                {

                    sql = "";
                    sql = "select bl_marks,bl_desc,bl_desc2,bl_desc_ctr,bl_is_mark from BLDESC ";
                    sql += " where bl_parent_id = '" + id + "' and bl_parent_type = 'SEAEXPDESC' ";
                    sql += " order by bl_desc_ctr";
                    DataTable Dt_Desc = new DataTable();
                    Dt_Desc = Con_Oracle.ExecuteQuery(sql);

                    foreach (DataRow dr in Dt_Desc.Rows)
                    {
                        Ctr = Lib.Conv2Integer(dr["bl_desc_ctr"].ToString());
                        if (Ctr == 1)
                        {
                            mRow.bl_mark1 = dr["bl_marks"].ToString();
                            mRow.bl_desc1 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark1 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 2)
                        {
                            mRow.bl_mark2 = dr["bl_marks"].ToString();
                            mRow.bl_desc2 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark2 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 3)
                        {
                            if (dr["bl_desc2"].ToString().Length > 15)
                            {
                                mRow.bl_grwt_caption = dr["bl_desc2"].ToString().Substring(0, 15).Trim();
                                mRow.bl_cbm_caption = dr["bl_desc2"].ToString().Substring(15).Trim();
                            }
                            mRow.bl_mark3 = dr["bl_marks"].ToString();
                            mRow.bl_desc3 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark3 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 4)
                        {
                            mRow.bl_mark4 = dr["bl_marks"].ToString();
                            mRow.bl_desc4 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark4 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 5)
                        {
                            if (dr["bl_desc2"].ToString().Length > 25)
                            {
                                mRow.bl_ntwt_caption = dr["bl_desc2"].ToString().Substring(0, 15).Trim();
                                mRow.bl_pcs_caption = dr["bl_desc2"].ToString().Substring(15, 10).Trim();
                                mRow.bl_pcs_unit_caption = dr["bl_desc2"].ToString().Substring(25).Trim();
                            }
                            mRow.bl_mark5 = dr["bl_marks"].ToString();
                            mRow.bl_desc5 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark5 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 6)
                        {
                            mRow.bl_mark6 = dr["bl_marks"].ToString();
                            mRow.bl_desc6 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark6 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 7)
                        {
                            mRow.bl_mark7 = dr["bl_marks"].ToString();
                            mRow.bl_desc7 = dr["bl_desc"].ToString();
                            mRow.bl_2desc7 = dr["bl_desc2"].ToString();
                            mRow.bl_is_mark7 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 8)
                        {
                            mRow.bl_mark8 = dr["bl_marks"].ToString();
                            mRow.bl_desc8 = dr["bl_desc"].ToString();
                            mRow.bl_2desc8 = dr["bl_desc2"].ToString();
                            mRow.bl_is_mark8 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 9)
                        {
                            mRow.bl_mark9 = dr["bl_marks"].ToString();
                            mRow.bl_desc9 = dr["bl_desc"].ToString();
                            mRow.bl_2desc9 = dr["bl_desc2"].ToString();
                            mRow.bl_is_mark9 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 10)
                        {
                            mRow.bl_mark10 = dr["bl_marks"].ToString();
                            mRow.bl_desc10 = dr["bl_desc"].ToString();
                            mRow.bl_2desc10 = dr["bl_desc2"].ToString();
                            mRow.bl_is_mark10 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 11)
                        {
                            mRow.bl_mark11 = dr["bl_marks"].ToString();
                            mRow.bl_desc11 = dr["bl_desc"].ToString();
                            mRow.bl_2desc11 = dr["bl_desc2"].ToString();
                            mRow.bl_is_mark11 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 12)
                        {
                            mRow.bl_mark12 = dr["bl_marks"].ToString();
                            mRow.bl_desc12 = dr["bl_desc"].ToString();
                            mRow.bl_2desc12 = dr["bl_desc2"].ToString();
                            mRow.bl_is_mark12 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 13)
                        {
                            mRow.bl_mark13 = dr["bl_marks"].ToString();
                            mRow.bl_desc13 = dr["bl_desc"].ToString();
                            mRow.bl_2desc13 = dr["bl_desc2"].ToString();
                            mRow.bl_is_mark13 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 14)
                        {
                            mRow.bl_mark14 = dr["bl_marks"].ToString();
                            mRow.bl_desc14 = dr["bl_desc"].ToString();
                            mRow.bl_2desc14 = dr["bl_desc2"].ToString();
                            mRow.bl_is_mark14 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 15)
                        {
                            mRow.bl_mark15 = dr["bl_marks"].ToString();
                            mRow.bl_desc15 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark15 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 16)
                        {
                            mRow.bl_mark16 = dr["bl_marks"].ToString();
                            mRow.bl_desc16 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark16 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 17)
                        {
                            mRow.bl_mark17 = dr["bl_marks"].ToString();
                            mRow.bl_desc17 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark17 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 18)
                        {
                            mRow.bl_mark18 = dr["bl_marks"].ToString();
                            mRow.bl_desc18 = dr["bl_desc"].ToString();
                            mRow.bl_is_mark18 = dr["bl_is_mark"].ToString() == "Y" ? true : false;
                        }
                        else if (Ctr == 19)
                        {
                            mRow.bl_mark19 = dr["bl_marks"].ToString();
                        }
                        else if (Ctr == 20)
                        {
                            mRow.bl_mark20 = dr["bl_marks"].ToString();
                        }
                        else if (Ctr == 21)
                        {
                            mRow.bl_mark21 = dr["bl_marks"].ToString();
                        }
                        else if (Ctr == 22)
                        {
                            mRow.bl_mark22 = dr["bl_marks"].ToString();
                        }
                        else if (Ctr == 23)
                        {
                            mRow.bl_mark23 = dr["bl_marks"].ToString();
                        }
                        else if (Ctr == 24)
                        {
                            mRow.bl_mark24 = dr["bl_marks"].ToString();
                        }
                    }


                    List<Bldesc> mList = new List<Bldesc>();
                    Bldesc bRow;

                    sql = "select a.bl_pkid, a.bl_parent_id, a.bl_marks, a.bl_desc  ";
                    sql += " from bldesc a ";
                    sql += " where bl_parent_id ='{ID}' ";
                    sql += " and bl_parent_type ='ATTACHEDBL' ";
                    sql += " order by bl_desc_ctr ";

                    sql = sql.Replace("{ID}", id);

                    Dt_Rec = new DataTable();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_Rec.Rows)
                    {
                        bRow = new Bldesc();
                        bRow.bl_pkid = Dr["bl_pkid"].ToString();
                        bRow.bl_marks = Dr["bl_marks"].ToString();
                        bRow.bl_desc = Dr["bl_desc"].ToString();
                        mList.Add(bRow);
                    }
                    mRow.AttachList = mList;

                    List<Containerm> cList = new List<Containerm>();
                    Containerm cRow;
                    sql = " select cntr_pkid,cntr_no,cntr_type_id,b.param_code as cntr_type_code,b.param_name as cntr_type_name ,cntr_csealno,";
                    sql += " cntr_asealno,cntr_pkg,cntr_pkg_unit_id,";
                    sql += " cntr_pcs,cntr_ntwt,cntr_grwt,cntr_cbm,cntr_shipment_type ";
                    sql += " from imp_container a";
                    sql += " left join param b on a.cntr_type_id = b.param_pkid";
                    sql += " where cntr_parent_id='{ID}' ";
                    sql = sql.Replace("{ID}", id);

                    Dt_Rec = new DataTable();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_Rec.Rows)
                    {
                        cRow = new Containerm();
                        cRow.cntr_pkid = Dr["cntr_pkid"].ToString();
                        cRow.cntr_no = Dr["cntr_no"].ToString();
                        cRow.cntr_type_id = Dr["cntr_type_id"].ToString();
                        cRow.cntr_type_code = Dr["cntr_type_code"].ToString();
                        cRow.cntr_type_name = Dr["cntr_type_name"].ToString();
                        cRow.cntr_csealno = Dr["cntr_csealno"].ToString();
                        cRow.cntr_asealno = Dr["cntr_asealno"].ToString();
                        cRow.cntr_pkg = Lib.Conv2Integer(Lib.NumericFormat(Dr["cntr_pkg"].ToString(), 0));
                        //cRow.cntr_pkg_unit_id = Dr["cntr_pkg_unit_id"].ToString();
                        cRow.cntr_pcs = Lib.Conv2Decimal(Lib.NumericFormat(Dr["cntr_pcs"].ToString(), 3));
                        cRow.cntr_ntwt = Lib.Conv2Decimal(Lib.NumericFormat(Dr["cntr_ntwt"].ToString(), 3));
                        cRow.cntr_grwt = Lib.Conv2Decimal(Lib.NumericFormat(Dr["cntr_grwt"].ToString(), 3));
                        cRow.cntr_cbm = Lib.Conv2Decimal(Lib.NumericFormat(Dr["cntr_cbm"].ToString(), 3));
                        cRow.cntr_shipment_type = Dr["cntr_shipment_type"].ToString();
                        cList.Add(cRow);
                    }
                    mRow.CntrList = cList;

                    List<HouseOrderm> hList = new List<HouseOrderm>();
                    HouseOrderm hRow;

                    sql = " select hord_pkid,hord_hbl_id,hord_cntr_id,b.cntr_no as hord_cntrno,hord_po,";
                    sql += " hord_style,hord_color,hord_invno,hord_pkg,";
                    sql += " hord_pkg_unit,hord_grwt,hord_ntwt,hord_pcs ,";
                    sql += " hord_cbm,hord_remarks  ";
                    sql += " from blorder a";
                    sql += " left join imp_container b on a.hord_cntr_id = b.cntr_pkid";
                    sql += " where hord_hbl_id='{ID}' order by b.cntr_no,hord_po";
                    sql = sql.Replace("{ID}", id);
                    Dt_Rec = new DataTable();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_Rec.Rows)
                    {
                        hRow = new HouseOrderm();
                        hRow.hord_pkid = Dr["hord_pkid"].ToString();
                        hRow.hord_hbl_id = Dr["hord_hbl_id"].ToString();
                        hRow.hord_cntr_id = Dr["hord_cntr_id"].ToString();
                        hRow.hord_no = Dr["hord_po"].ToString();
                        hRow.hord_style = Dr["hord_style"].ToString();
                        hRow.hord_color = Dr["hord_color"].ToString();
                        hRow.hord_color = Dr["hord_color"].ToString();
                        hRow.hord_pkgs = Lib.Conv2Integer(Dr["hord_pkg"].ToString());
                        hRow.hord_pkgs_unit = Dr["hord_pkg_unit"].ToString();
                        hRow.hord_grwt = Lib.Conv2Decimal(Lib.NumericFormat( Dr["hord_grwt"].ToString(),3));
                        hRow.hord_ntwt = Lib.Conv2Decimal(Lib.NumericFormat(Dr["hord_ntwt"].ToString(),3));
                        hRow.hord_pcs = Lib.Conv2Decimal(Lib.NumericFormat(Dr["hord_pcs"].ToString(),3));
                        hRow.hord_cbm = Lib.Conv2Decimal(Lib.NumericFormat(Dr["hord_cbm"].ToString(),3));
                        hRow.hord_cntrno = Dr["hord_cntrno"].ToString();
                        hRow.hord_remarks = Dr["hord_remarks"].ToString();
                        hList.Add(hRow);
                    }
                    mRow.OrdList = hList;
                }

                if(type=="PDF")
                {
                    mRow.bl_invoke_frm = "HBL";
                    ProcessArrivalNotice(report_folder, folderid, comp_code, mRow);
                }

                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("record", mRow);
            return RetData;
        }

        private void ProcessArrivalNotice(string report_folder, string folderid,string comp_code, Bl mRow)
        {
            File_Name = "";
            File_Type = "";
            File_Display_Name = "myreport.pdf";
            string RootPath = "";
            RootPath = report_folder;

            if (mRow.hbl_bl_no.ToString().Length > 0)
                File_Display_Name = Lib.ProperFileName(mRow.hbl_bl_no.ToString()) + ".pdf";
            File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
            File_Type = "pdf";


            Dictionary<string, object> mSearchData = new Dictionary<string, object>();
            LovService mService = new LovService();
            mSearchData.Add("table", "COMP_ADDRESS");
            mSearchData.Add("comp_code", comp_code);

            DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
            if (Dt_CompAddress != null)
            {
                foreach (DataRow Dr in Dt_CompAddress.Rows)
                {
                    mRow.bl_issued_by1 = Dr["COMP_NAME"].ToString();
                    mRow.bl_issued_by2 = Dr["COMP_ADDRESS1"].ToString();
                    mRow.bl_issued_by3 = Dr["COMP_ADDRESS2"].ToString();
                    mRow.bl_issued_by4 = Dr["comp_address3"].ToString();
                    //mRow.bl_issued_by1 = Dr["COMP_FAX"].ToString();
                    //mRow.bl_issued_by1 = Dr["COMP_WEB"].ToString();
                    break;
                }
            }

            BLReport mReport = new BLReport();
            mReport.mRow = mRow;
            mReport.RootPath = RootPath;
            mReport.Chk_BL_Original = false;
            mReport.dColr = "0";
            mReport.InvokeType = "HBL";
            mReport.ProcessData();
            if (mReport.ExportList != null)
            {
                if (Lib.CreateFolder(report_folder))
                {
                    Export2Pdf mypdf = new Export2Pdf();
                    mypdf.ExportList = mReport.ExportList;
                    mypdf.FileName = File_Name;
                    mypdf.Page_Height = 1120;
                    mypdf.Page_Width = 800;
                    mypdf.Process();
                    //if (formattype == "SEABL")
                    //    if (UpdateBL(mRow.bl_pkid, mRow.bl_is_original))
                    //    {
                    //        mRow.bl_print = "Y";
                    //        if (mRow.bl_is_original)
                    //            mRow.bl_original_print = "Y";
                    //    }
                }
            }

        }

    }
}
