using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLEdi
{
    public class EdiBLService : BL_Base
    {
        string JOBORDERID = "";
        DataTable Dt_EdiHouse = null;
        DataTable Dt_EdiContainer = null;
        DataTable Dt_EdiTracking = null;
        DataTable Dt_EdiOrders = null;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string type = SearchData["type"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string User_Code = SearchData["user_code"].ToString();
            string partnerid = SearchData["partnerid"].ToString();
            string rowstatus = SearchData["rowstatus"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string hblstatus = SearchData["hblstatus"].ToString();
            string fileno = SearchData["fileno"].ToString();
            Boolean showdeleted = (Boolean) SearchData["showdeleted"];

            string houseno = "";
            string masterno = "";

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            Con_Oracle = new DBConnection();
            List<edi_hbl> mList = new List<edi_hbl>();
            edi_hbl mRow;
            string sWhere = "";


            if (SearchData.ContainsKey("masterno"))
                masterno = SearchData["masterno"].ToString();
            if (SearchData.ContainsKey("houseno"))
                houseno = SearchData["houseno"].ToString();

            try
            {
                sWhere = "";
                sWhere = "where a.rec_company_code = '{COMPCODE}'  ";

                if (partnerid != "ALL")
                    sWhere += " and a.hbl_sender = '" + partnerid + "'";

                if (hblstatus != "ALL")
                    sWhere += " and a.hbl_status = '" + hblstatus + "'";

                if ( rowstatus != "ALL")
                    sWhere += " and a.hbl_updated = '" + rowstatus + "'";

                if ( showdeleted)
                    sWhere += " and a.rec_deleted = 'Y'";
                else
                    sWhere += " and a.rec_deleted = 'N'";

                if (fileno.Length > 0)
                    sWhere += " and b.messagenumber = '" + fileno + "'";

                if (masterno.Trim().Length > 0)
                    sWhere += " and a.hbl_master_no like '%" + masterno.Trim() + "%'";

                if (houseno.Trim().Length > 0)
                    sWhere += " and a.hbl_house_no like '%" + houseno.Trim() + "%'";

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM edi_house a inner join edi_header b on a.hbl_headerid = b.headerid ";
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
                sql += " select b.slno, hbl_pkid,hbl_headerid,hbl_sender,messagedate,messagefilename,messagenumber,transfer_remarks, ";
                sql += " hbl_status,hbl_updated, hbl_pol_agent, hbl_house_no,hbl_master_no,hbl_carrier_name,hbl_pol_name,hbl_pod_name,hbl_etd,hbl_eta, ";
                sql += " hbl_freight,hbl_shipper_name,hbl_consignee_name,hbl_vessel,hbl_voyage, a.rec_deleted, ";
                sql += " row_number() over(order by slno,hbl_sender,hbl_house_no ) rn";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by slno,hbl_sender,hbl_house_no ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new edi_hbl();
                    mRow.hbl_pkid = Dr["hbl_pkid"].ToString();

                    mRow.hbl_header_id = Dr["hbl_headerid"].ToString();
                    mRow.hbl_sender = Dr["hbl_sender"].ToString();
                    mRow.hbl_message_date = Dr["messagedate"].ToString();
                    mRow.hbl_message_file_name = Dr["messagefilename"].ToString();
                    mRow.hbl_message_number = Dr["messagenumber"].ToString();

                    mRow.hbl_status = Dr["hbl_status"].ToString();
                    mRow.hbl_updated = Dr["hbl_updated"].ToString();

                    mRow.hbl_house_no = Dr["hbl_house_no"].ToString();
                    mRow.hbl_master_no = Dr["hbl_master_no"].ToString();
                    mRow.hbl_carrier_name = Dr["hbl_carrier_name"].ToString();
                    mRow.hbl_pol_name = Dr["hbl_pol_name"].ToString();
                    mRow.hbl_pod_name  = Dr["hbl_pod_name"].ToString();
                    mRow.hbl_etd = Lib.DatetoString(Dr["hbl_etd"]);
                    mRow.hbl_eta = Lib.DatetoString(Dr["hbl_eta"]);
                    mRow.hbl_freight = Dr["hbl_freight"].ToString();
                    mRow.hbl_shipper_name = Dr["hbl_shipper_name"].ToString();
                    mRow.hbl_consignee_name = Dr["hbl_consignee_name"].ToString();
                    mRow.hbl_vessel = Dr["hbl_vessel"].ToString();
                    mRow.hbl_voyage = Dr["hbl_voyage"].ToString();
 
                    mRow.rec_deleted = Dr["rec_deleted"].ToString() == "Y" ?  true : false;
                    mRow.transfer_remarks = Dr["transfer_remarks"].ToString();

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
            return RetData;
        }


        public IDictionary<string, object> Validate(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string company_code = SearchData["company_code"].ToString();
            //string partnerid = SearchData["partnerid"].ToString();

            Con_Oracle = new DBConnection();
            List<edi_missingdata> mList = new List<edi_missingdata>();
            edi_missingdata mRow;

            try
            {

                DataTable Dt_List = new DataTable();

                sql = "";
                sql += " select * from( ";

                sql += " select distinct hbl_sender as sender, 'SHIPPER' as type, source , link_target_id as target_id, link_target_name as target_name ";
                sql += " from( ";
                sql += " select hbl_sender, a.hbl_shipper_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner ";
                sql += " join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'SHIPPER' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and a.hbl_shipper_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";

                sql += " select distinct  hbl_sender,'CONSIGNEE' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,a.hbl_consignee_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'CONSIGNEE' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and a.hbl_consignee_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";

                sql += " select distinct  hbl_sender,'POL' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,a.hbl_pol_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " left ";
                sql += " join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'PORT' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and a.hbl_pol_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";

                sql += " select distinct  hbl_sender,'POD' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,a.hbl_pod_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'PORT' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and a.hbl_pod_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";


                sql += " select distinct  hbl_sender,'CONTAINER TYPE' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,cntr.cntr_size as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " inner join edi_house_container cntr on a.hbl_pkid = cntr.cntr_hbl_id";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'CONTAINER TYPE' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and cntr.cntr_size = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";


                sql += " select distinct  hbl_sender,'UNIT' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,cntr.cntr_pkgs_unit as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " inner join edi_house_container cntr on a.hbl_pkid = cntr.cntr_hbl_id";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'UNIT' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and cntr.cntr_pkgs_unit = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";

                sql += " select distinct  hbl_sender,'VESSEL' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,vsl.vsl_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " inner join edi_house_vessel vsl on a.hbl_pkid = vsl.vsl_hbl_id";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'VESSEL' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and vsl.vsl_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";

                sql += " select distinct  hbl_sender,'PORT' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,vsl.vsl_pol_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " inner join edi_house_vessel vsl on a.hbl_pkid = vsl.vsl_hbl_id";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'PORT' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and vsl.vsl_pol_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";

                sql += " select distinct  hbl_sender,'PORT' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,vsl.vsl_pod_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " inner join edi_house_vessel vsl on a.hbl_pkid = vsl.vsl_hbl_id";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'PORT' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and vsl.vsl_pod_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";

                sql += " union all ";

                sql += " select distinct  hbl_sender,'POL-AGENT' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  hbl_sender,a.hbl_pol_agent as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_house a ";
                sql += " inner join edi_header b on a.hbl_headerid = b.headerid ";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'POL-AGENT' and  a.rec_company_code = c.rec_company_code and a.hbl_sender = c.link_messagesender and a.hbl_pol_agent = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";
                sql += " ) b ";
                sql += " order by type, source ";
                sql = sql.Replace("{COMPCODE}", company_code);

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new edi_missingdata();
                    mRow.sender = Dr["sender"].ToString();
                    mRow.type = Dr["type"].ToString();
                    mRow.source = Dr["source"].ToString();
                    mRow.target_id = "";
                    mRow.target_name = "";
                    mList.Add(mRow);
                }

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("list", mList);
            return RetData;
        }

        public IDictionary<string, object> TransferBL(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string company_code = SearchData["company_code"].ToString();
            string sender = SearchData["partnerid"].ToString();
            string user_code = SearchData["user_code"].ToString();
            Con_Oracle = new DBConnection();

            string Remarks = "";
            Boolean SaveOk = false;

            int processed = 0;
            int notprocessed = 0;

            Boolean SaveMaster = false;
            string MBL_ID = "";
            DataTable Dt_temp;
            string HouseNos = ""; 
            string MsgNo = "";
            try
            {

                sql = "";
                sql = " select messagenumber,hbl_pkid as bl_pkid,hbl_master_no,hbl_master_date,a.rec_company_code, nvl(a.hbl_mode,'SEA') as rec_category,a.hbl_direct_bl";
                sql += " ,hbl_etd as bl_pol_etd,hbl_eta as bl_pod_eta, a.hbl_freight as bl_freight,hbl_movement as bl_move_type,hbl_issu_place as bl_issued_place,hbl_issu_date as bl_issued_date ";
                sql += " ,shpr_link.link_target_id as bl_shipper_id, shpr_link.link_target_name as bl_shipper_name";
                sql += " ,a.hbl_shipper_add1 as bl_shipper_add1,a.hbl_shipper_add2 as bl_shipper_add2,a.hbl_shipper_add3 as bl_shipper_add3,a.hbl_shipper_add4 as bl_shipper_add4";
                sql += " ,agent_link.link_target_id as bl_agent_id, agent_link.link_target_name as bl_agent_name";
                sql += " ,carr_link.link_target_id as bl_carrier_id, carr_link.link_target_name as bl_carrier_name";
                sql += " ,cons_link.link_target_id as bl_consignee_id, cons_link.link_target_name as bl_consignee_name";
                sql += " ,a.hbl_consignee_add1 as bl_consignee_add1,a.hbl_consignee_add2 as bl_consignee_add2,a.hbl_consignee_add3 as bl_consignee_add3,a.hbl_consignee_add4 as bl_consignee_add4";
                sql += " ,notify_link.link_target_id as bl_notify_id, notify_link.link_target_name as bl_notify_name";
                sql += " ,a.hbl_notify_add1 as bl_notify_add1,a.hbl_notify_add2 as bl_notify_add2,a.hbl_notify_add3 as bl_notify_add3,a.hbl_notify_add4 as bl_notify_add4";
                sql += " ,a.hbl_por_name as bl_place_receipt,a.hbl_pol_name as bl_pol,pol_link.link_target_id as bl_pol_id,a.hbl_pod_name as bl_pod,pod_link.link_target_id as bl_pod_id,a.hbl_delivery_name as bl_place_delivery";
                sql += " ,a.hbl_vessel as bl_vsl_name,a.hbl_voyage as bl_vsl_voy_no,a.hbl_house_no as bl_bl_no,hbl_house_date ";
                sql += " ,br.comp_code as rec_branch_code, hbl_pkg as bl_pkg, hbl_pkg_unit,unit_link.link_target_id as bl_pkg_unit_id,hbl_pcs as bl_pcs,hbl_grwt as bl_grwt,hbl_ntwt as bl_ntwt,hbl_chwt as bl_chwt,hbl_cbm as bl_cbm ";
                sql += " from edi_house a inner join edi_header b on a.hbl_headerid = b.headerid";
                sql += " left join edi_link shpr_link  on shpr_link.link_type = 'INWARD' and shpr_link.link_subcategory = 'SHIPPER' and a.rec_company_code = shpr_link.rec_company_code and a.hbl_sender = shpr_link.link_messagesender   and a.hbl_shipper_name = shpr_link.link_source_name";
                sql += " left join edi_link agent_link on agent_link.link_type = 'INWARD' and agent_link.link_subcategory = 'POL-AGENT' and a.rec_company_code = agent_link.rec_company_code and a.hbl_sender = agent_link.link_messagesender and a.hbl_pol_agent = agent_link.link_source_name";
                sql += " left join edi_link carr_link on carr_link.link_type = 'INWARD' and carr_link.link_subcategory = 'CARRIER' and a.rec_company_code = carr_link.rec_company_code and a.hbl_sender = carr_link.link_messagesender and a.hbl_carrier_name = carr_link.link_source_name";
                sql += " left join edi_link cons_link  on cons_link.link_type = 'INWARD' and cons_link.link_subcategory = 'CONSIGNEE' and a.rec_company_code = cons_link.rec_company_code and a.hbl_sender = cons_link.link_messagesender   and a.hbl_consignee_name = cons_link.link_source_name";
                sql += " left join edi_link notify_link  on notify_link.link_type = 'INWARD' and notify_link.link_subcategory = 'NOTIFY' and a.rec_company_code = notify_link.rec_company_code and a.hbl_sender = notify_link.link_messagesender   and a.hbl_notify_name = notify_link.link_source_name";
                sql += " left join edi_link pol_link  on pol_link.link_type = 'INWARD' and pol_link.link_subcategory = 'PORT' and a.rec_company_code = pol_link.rec_company_code and a.hbl_sender = pol_link.link_messagesender and a.hbl_pol_name = pol_link.link_source_name ";
                sql += " left join edi_link pod_link  on pod_link.link_type = 'INWARD' and pod_link.link_subcategory = 'PORT' and a.rec_company_code = pod_link.rec_company_code and a.hbl_sender = pod_link.link_messagesender and a.hbl_pod_name = pod_link.link_source_name ";
                sql += " left join edi_link unit_link  on unit_link.link_type = 'INWARD' and unit_link.link_subcategory = 'UNIT' and a.rec_company_code = unit_link.rec_company_code and a.hbl_sender = unit_link.link_messagesender  and a.hbl_pkg_unit = unit_link.link_source_name";
                sql += " left join edi_link branch_link  on branch_link.link_type = 'INWARD' and branch_link.link_subcategory = 'BRANCH' and a.rec_company_code = branch_link.rec_company_code and a.hbl_sender = branch_link.link_messagesender and a.hbl_pod_name = branch_link.link_source_name ";
                sql += " left join companym br on branch_link.link_target_id = br.comp_pkid and br.comp_type='B' ";
                sql += " where a.rec_company_code = '{COMPCODE}'  ";
                if (sender != "ALL")
                    sql += " and a.hbl_sender = '{SENDER}' ";
                sql += " and a.hbl_updated = 'N' and a.rec_deleted = 'N' ";

                sql += " order by hbl_master_no ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{SENDER}", sender);
                
                Dt_EdiHouse = new DataTable();
                Dt_EdiHouse = Con_Oracle.ExecuteQuery(sql);


                sql = " select cntr_pkid,cntr_hbl_id,cntr_no,cntr_size,cntr_type,";
                sql += " cntrtype_link.link_target_id as cntr_size_id, cntrtype_link.link_target_name as cntr_size_name,";
                sql += " cntr_aseal,cntr_cseal,cntr_pkgs,cntr_pkgs_unit,";
                sql += " cntrunit_link.link_target_id as cntr_pkgs_unit_id, cntrunit_link.link_target_name as cntr_pkgs_unit_name,";
                sql += " cntr_grwt,cntr_ntwt,cntr_pcs,cntr_cbm,a.rec_company_code ";
                sql += " ,br.comp_code as rec_branch_code ";
                sql += " from edi_house a";
                sql += " inner join edi_house_container b on a.hbl_pkid = b.cntr_hbl_id";
                sql += " left join edi_link cntrtype_link  on cntrtype_link.link_type = 'INWARD' and cntrtype_link.link_subcategory = 'CONTAINER TYPE' and a.rec_company_code = cntrtype_link.rec_company_code and a.hbl_sender = cntrtype_link.link_messagesender  and b.cntr_size = cntrtype_link.link_source_name";
                sql += " left join edi_link cntrunit_link  on cntrunit_link.link_type = 'INWARD' and cntrunit_link.link_subcategory = 'UNIT' and a.rec_company_code = cntrunit_link.rec_company_code and a.hbl_sender = cntrunit_link.link_messagesender  and b.cntr_pkgs_unit = cntrunit_link.link_source_name";
                sql += " left join edi_link branch_link  on branch_link.link_type = 'INWARD' and branch_link.link_subcategory = 'BRANCH' and a.rec_company_code = branch_link.rec_company_code and a.hbl_sender = branch_link.link_messagesender and a.hbl_pod_name = branch_link.link_source_name ";
                sql += " left join companym br on branch_link.link_target_id = br.comp_pkid and br.comp_type='B' ";
                sql += " where a.rec_company_code = '{COMPCODE}'  ";
                if (sender != "ALL")
                    sql += " and a.hbl_sender = '{SENDER}' ";
                sql += " and a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " order by hbl_house_no,cntr_no ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{SENDER}", sender);

                Dt_EdiContainer = new DataTable();
                Dt_EdiContainer = Con_Oracle.ExecuteQuery(sql);


                sql = " select vsl_pkid as trk_pkid,vsl_hbl_id as trk_parent_id,";
                sql += " vsl_seq as trk_order,vsl_link.link_target_id as trk_vsl_id, vsl_link.link_target_name as trk_vsl_name,vsl_voyage as trk_voyage,";
                sql += " pol_link.link_target_id as trk_pol_id, pol_link.link_target_name as trk_pol_name, ";
                sql += " pod_link.link_target_id as trk_pod_id, pod_link.link_target_name as trk_pod_name, ";
                sql += " vsl_etd as trk_pol_etd,vsl_etd_confirm as trk_pol_etd_confirm,";
                sql += " vsl_eta as trk_pod_eta,vsl_eta_confirm as trk_pod_eta_confirm,a.rec_company_code";
                sql += " ,br.comp_code as rec_branch_code ";
                sql += " from edi_house a";
                sql += " inner join edi_house_vessel b on a.hbl_pkid = b.vsl_hbl_id";
                sql += " left join edi_link vsl_link  on vsl_link.link_type = 'INWARD' and vsl_link.link_subcategory = 'VESSEL' and a.rec_company_code = vsl_link.rec_company_code and a.hbl_sender = vsl_link.link_messagesender and b.vsl_name = vsl_link.link_source_name";
                sql += " left join edi_link pol_link  on pol_link.link_type = 'INWARD' and pol_link.link_subcategory = 'PORT' and a.rec_company_code = pol_link.rec_company_code and a.hbl_sender = pol_link.link_messagesender and b.vsl_pol_name = pol_link.link_source_name ";
                sql += " left join edi_link pod_link  on pod_link.link_type = 'INWARD' and pod_link.link_subcategory = 'PORT' and a.rec_company_code = pod_link.rec_company_code and a.hbl_sender = pod_link.link_messagesender and b.vsl_pod_name = pod_link.link_source_name ";
                sql += " left join edi_link branch_link  on branch_link.link_type = 'INWARD' and branch_link.link_subcategory = 'BRANCH' and a.rec_company_code = branch_link.rec_company_code and a.hbl_sender = branch_link.link_messagesender and a.hbl_pod_name = branch_link.link_source_name ";
                sql += " left join companym br on branch_link.link_target_id = br.comp_pkid and br.comp_type='B' ";
                sql += " where a.rec_company_code = '{COMPCODE}'  ";
                if (sender != "ALL")
                    sql += " and a.hbl_sender = '{SENDER}' ";
                sql += " and a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " order by hbl_house_no ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{SENDER}", sender);

                Dt_EdiTracking = new DataTable();
                Dt_EdiTracking = Con_Oracle.ExecuteQuery(sql);


                sql = " select cntr_pkid,cntr_hbl_id,cntr_no,cntr_size,cntr_type,";
                sql += " cntrtype_link.link_target_id as cntr_size_id, cntrtype_link.link_target_name as cntr_size_name,";
                sql += " cntr_aseal,cntr_cseal,cntr_pkgs,cntr_pkgs_unit,";
                sql += " cntrunit_link.link_target_id as cntr_pkgs_unit_id, cntrunit_link.link_target_name as cntr_pkgs_unit_name,";
                sql += " cntr_grwt,cntr_ntwt,cntr_pcs,cntr_cbm,a.rec_company_code ";
                sql += " ,br.comp_code as rec_branch_code ";
                sql += " from edi_house a";
                sql += " inner join edi_house_container b on a.hbl_pkid = b.cntr_hbl_id";
                sql += " left join edi_link cntrtype_link  on cntrtype_link.link_type = 'INWARD' and cntrtype_link.link_subcategory = 'CONTAINER TYPE' and a.rec_company_code = cntrtype_link.rec_company_code and a.hbl_sender = cntrtype_link.link_messagesender  and b.cntr_size = cntrtype_link.link_source_name";
                sql += " left join edi_link cntrunit_link  on cntrunit_link.link_type = 'INWARD' and cntrunit_link.link_subcategory = 'UNIT' and a.rec_company_code = cntrunit_link.rec_company_code and a.hbl_sender = cntrunit_link.link_messagesender  and b.cntr_pkgs_unit = cntrunit_link.link_source_name";
                sql += " left join edi_link branch_link  on branch_link.link_type = 'INWARD' and branch_link.link_subcategory = 'BRANCH' and a.rec_company_code = branch_link.rec_company_code and a.hbl_sender = branch_link.link_messagesender and a.hbl_pod_name = branch_link.link_source_name ";
                sql += " left join companym br on branch_link.link_target_id = br.comp_pkid and br.comp_type='B' ";
                sql += " where a.rec_company_code = '{COMPCODE}'  ";
                if (sender != "ALL")
                    sql += " and a.hbl_sender = '{SENDER}' ";
                sql += " and a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " order by hbl_house_no,cntr_no ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{SENDER}", sender);

                Dt_EdiContainer = new DataTable();
                Dt_EdiContainer = Con_Oracle.ExecuteQuery(sql);


                sql = " select ho_pkid,ho_hbl_id,ho_cntr_id,ho_ordno,ho_style,";
                sql += " ho_color,ho_invno,ho_pkgs,ho_pkgs_unit,";
                sql += " ho_grwt,ho_ntwt,ho_pcs ,ho_cbm,'' as ho_ord_id,'' as ho_ord_remarks,a.rec_company_code ";
                sql += " ,br.comp_code as rec_branch_code ";
                sql += " from edi_house a";
                sql += " inner join edi_house_order b on a.hbl_pkid = b.ho_hbl_id";
                sql += " left join edi_link branch_link  on branch_link.link_type = 'INWARD' and branch_link.link_subcategory = 'BRANCH' and a.rec_company_code = branch_link.rec_company_code and a.hbl_sender = branch_link.link_messagesender and a.hbl_pod_name = branch_link.link_source_name ";
                sql += " left join companym br on branch_link.link_target_id = br.comp_pkid and br.comp_type='B' ";
                sql += " where a.rec_company_code = '{COMPCODE}'  ";
                if (sender != "ALL")
                    sql += " and a.hbl_sender = '{SENDER}' ";
                sql += " and a.hbl_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " order by ho_ordno ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{SENDER}", sender);

                Dt_EdiOrders = new DataTable();
                Dt_EdiOrders = Con_Oracle.ExecuteQuery(sql);


                DataTable DistinctMBL = Dt_EdiHouse.DefaultView.ToTable(true, "hbl_master_no", "bl_agent_id");
                foreach (DataRow Dr in DistinctMBL.Rows)
                {
                    SaveMaster = false;
                    MBL_ID = Guid.NewGuid().ToString().ToUpper();
                    sql = "select bl_pkid from bl where bl_mbl_no='" + Dr["hbl_master_no"].ToString() + "' and bl_agent_id='" + Dr["bl_agent_id"].ToString() + "'";
                    Dt_temp = new DataTable();
                    Dt_temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_temp.Rows.Count > 0)
                        MBL_ID = Dt_temp.Rows[0]["bl_pkid"].ToString();

                    Con_Oracle.BeginTransaction();
                    HouseNos = "";
                    foreach (DataRow Dr_target in Dt_EdiHouse.Select("hbl_master_no='" + Dr["hbl_master_no"].ToString() + "'"))
                    {
                        Remarks = "";
                        SaveOk = false;
                        if (Remarks == "")
                            Remarks = IsValid(Dr_target);
                        if (Remarks == "")
                            Remarks = CheckDuplication(Dr_target);
                        if (Remarks == "")
                        {
                            if (!SaveMaster)
                            {
                                MsgNo = Dr_target["messagenumber"].ToString();
                                SaveMaster = true;
                                InsertMaster(Dr_target, MBL_ID, user_code);
                                InsertVesselDetails(MBL_ID, Dr_target["bl_pkid"].ToString(), user_code);
                            }

                            InsertHouse(Dr_target, MBL_ID, user_code);
                            InsertHouseContainer(Dr_target["bl_pkid"].ToString(),user_code);
                            InsertHouseOrders(Dr_target["bl_pkid"].ToString(), user_code);
                            InsertDescriptions(Dr_target["bl_pkid"].ToString(), Dr_target["rec_category"].ToString(), user_code);

                            Remarks = "TRANSFERED";
                            SaveOk = true;
                            if (HouseNos != "")
                                HouseNos += ",";
                            HouseNos += Dr_target["bl_bl_no"].ToString();
                        }
                        if (Remarks != "")
                        {
                            Remarks = Remarks.Length > 100 ? Remarks.Substring(0, 100) : Remarks;
                            sql = "update edi_house set transfer_remarks = '" + Remarks + "' ";
                            if (SaveOk)
                            {
                                sql += ",hbl_house_id = '" + Dr_target["BL_PKID"].ToString() + "'";
                                sql += ",hbl_updated = 'Y' ";
                                processed++;
                            }
                            else
                            {
                                notprocessed++;
                            }
                            sql += " where hbl_pkid = '" + Dr_target["bl_pkid"].ToString() + "'";
                            Con_Oracle.ExecuteNonQuery(sql);
                        }
                    }

                    InsertMasterContainer(MBL_ID, user_code);
                    Con_Oracle.CommitTransaction();

                    if (SaveOk)
                    {
                        string srem = "EDI FILE# " + MsgNo + " Housenos " + HouseNos;
                        Lib.AuditLog("BL", "BL", "EDI ADD", company_code, "", user_code, MBL_ID, Dr["hbl_master_no"].ToString(), srem);
                    }

                }

                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("status", "TRANSFERED : " + processed.ToString() + ",  NOT PROCESSED " + notprocessed.ToString());
            return RetData;
        }


        private void InsertMaster(DataRow dr, string masterid, string user_code)
        {
            // sql = "select bl_pkid from bl where bl_mbl_no='" + dr["hbl_master_no"].ToString() + "' and bl_agent_id='" + dr["bl_agent_id"].ToString() + "'";
            sql = "select bl_pkid from bl where bl_pkid='" + masterid + "'";
            if (Con_Oracle.IsRowExists(sql))
                return;

            DBRecord Rec = new DBRecord();
            Rec.CreateRow("BL", "ADD", "BL_PKID", masterid);
            Rec.InsertString("BL_TYPE", "MBL");
            Rec.InsertString("BL_AGENT_ID", dr["BL_AGENT_ID"].ToString());
            Rec.InsertString("BL_AGENT_NAME", dr["BL_AGENT_NAME"].ToString());
            Rec.InsertString("BL_CARRIER_ID", dr["BL_CARRIER_ID"].ToString());
            Rec.InsertString("BL_CARRIER_NAME", dr["BL_CARRIER_NAME"].ToString());
            Rec.InsertString("BL_IS_DIRECT", dr["HBL_DIRECT_BL"].ToString());
            Rec.InsertString("bl_mbl_no", dr["hbl_master_no"].ToString());
            Rec.InsertDate("bl_bl_date", dr["hbl_master_date"]);
            Rec.InsertString("bl_bl_no", dr["hbl_master_no"].ToString());
            Rec.InsertString("bl_place_receipt", dr["bl_place_receipt"].ToString());
            Rec.InsertDate("bl_pol_etd", dr["bl_pol_etd"]);
            Rec.InsertDate("bl_pod_eta", dr["bl_pod_eta"]);
            Rec.InsertString("bl_pol_id", dr["bl_pol_id"].ToString());
            Rec.InsertString("bl_pod_id", dr["bl_pod_id"].ToString());
            Rec.InsertString("bl_pol", dr["bl_pol"].ToString());
            Rec.InsertString("bl_pod", dr["bl_pod"].ToString());
            Rec.InsertString("bl_place_delivery", dr["bl_place_delivery"].ToString());
            Rec.InsertString("bl_vsl_name", dr["bl_vsl_name"].ToString());
            Rec.InsertString("bl_vsl_voy_no", dr["bl_vsl_voy_no"].ToString());
            Rec.InsertString("bl_freight", dr["bl_freight"].ToString());

            Rec.InsertString("REC_CREATED_BY", user_code);
            Rec.InsertDate("REC_CREATED_DATE", System.DateTime.Today);
            Rec.InsertString("REC_COMPANY_CODE", dr["REC_COMPANY_CODE"].ToString());
            Rec.InsertString("REC_BRANCH_CODE", dr["REC_BRANCH_CODE"].ToString());
            Rec.InsertString("REC_CATEGORY", dr["REC_CATEGORY"].ToString() + " EXPORT");

            sql = Rec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);
        }
        private void InsertHouse(DataRow dr, string masterid, string user_code)
        {
           // sql = "select bl_pkid from bl where bl_bl_no='" + dr["bl_bl_no"].ToString() + "' and bl_agent_id='" + dr["bl_agent_id"].ToString() + "'";

            sql = "select bl_pkid from bl where bl_pkid='" + dr["bl_pkid"].ToString() + "'";
            if (Con_Oracle.IsRowExists(sql))
                return;

            DBRecord Rec = new DBRecord();
            Rec.CreateRow("BL", "ADD", "BL_PKID", dr["BL_PKID"].ToString());
            Rec.InsertString("BL_TYPE", "HBL");
            Rec.InsertString("BL_MBL_ID", masterid);
            Rec.InsertString("BL_AGENT_ID", dr["BL_AGENT_ID"].ToString());
            Rec.InsertString("BL_AGENT_NAME", dr["BL_AGENT_NAME"].ToString());
            Rec.InsertString("BL_CARRIER_ID", dr["BL_CARRIER_ID"].ToString());
            Rec.InsertString("BL_CARRIER_NAME", dr["BL_CARRIER_NAME"].ToString());

            Rec.InsertString("bl_shipper_id", dr["bl_shipper_id"].ToString());
            // Rec.InsertString("bl_shipper_br_id", bl_shipper_br_id);
            Rec.InsertString("bl_shipper_name", dr["bl_shipper_name"].ToString());
            Rec.InsertString("bl_shipper_add1", dr["bl_shipper_add1"].ToString());
            Rec.InsertString("bl_shipper_add2", dr["bl_shipper_add2"].ToString());
            Rec.InsertString("bl_shipper_add3", dr["bl_shipper_add3"].ToString());
            Rec.InsertString("bl_shipper_add4", dr["bl_shipper_add4"].ToString());

            Rec.InsertString("bl_consignee_id", dr["bl_consignee_id"].ToString());
            // Rec.InsertString("bl_consignee_br_id", Record.bl_consignee_br_id);
            Rec.InsertString("bl_consignee_name", dr["bl_consignee_name"].ToString());
            Rec.InsertString("bl_consignee_add1", dr["bl_consignee_add1"].ToString());
            Rec.InsertString("bl_consignee_add2", dr["bl_consignee_add2"].ToString());
            Rec.InsertString("bl_consignee_add3", dr["bl_consignee_add3"].ToString());
            Rec.InsertString("bl_consignee_add4", dr["bl_consignee_add4"].ToString());
            //Rec.InsertString("bl_issued_by1", Record.bl_issued_by1, "P");
            //Rec.InsertString("bl_issued_by2", Record.bl_issued_by2, "P");
            //Rec.InsertString("bl_issued_by3", Record.bl_issued_by3, "P");
            //Rec.InsertString("bl_issued_by4", Record.bl_issued_by4, "P");
            //Rec.InsertString("bl_issued_by5", Record.bl_issued_by5, "P");
            Rec.InsertString("bl_notify_id", dr["bl_notify_id"].ToString());
            // Rec.InsertString("bl_notify_br_id", Record.bl_notify_br_id);
            Rec.InsertString("bl_notify_name", dr["bl_notify_name"].ToString());
            Rec.InsertString("bl_notify_add1", dr["bl_notify_add1"].ToString());
            Rec.InsertString("bl_notify_add2", dr["bl_notify_add2"].ToString());
            Rec.InsertString("bl_notify_add3", dr["bl_notify_add3"].ToString());
            Rec.InsertString("bl_notify_add4", dr["bl_notify_add4"].ToString());

            Rec.InsertDate("bl_pol_etd", dr["bl_pol_etd"]);
            Rec.InsertDate("bl_pod_eta", dr["bl_pod_eta"]);
            Rec.InsertString("bl_place_receipt", dr["bl_place_receipt"].ToString());
            //Rec.InsertDate("bl_date_receipt", Record.bl_date_receipt);
            Rec.InsertString("bl_freight", dr["bl_freight"].ToString());
            Rec.InsertString("bl_pol_id", dr["bl_pol_id"].ToString());
            Rec.InsertString("bl_pod_id", dr["bl_pod_id"].ToString());
            Rec.InsertString("bl_pol", dr["bl_pol"].ToString());
            Rec.InsertString("bl_pod", dr["bl_pod"].ToString());
            Rec.InsertString("bl_place_delivery", dr["bl_place_delivery"].ToString());
            //Rec.InsertString("bl_delivery_contact1", Record.bl_delivery_contact1);
            //Rec.InsertString("bl_delivery_contact2", Record.bl_delivery_contact2);
            //Rec.InsertString("bl_delivery_contact3", Record.bl_delivery_contact3);
            //Rec.InsertString("bl_delivery_contact4", Record.bl_delivery_contact4);
            //Rec.InsertString("bl_delivery_contact5", Record.bl_delivery_contact5);
            //Rec.InsertString("bl_delivery_contact6", Record.bl_delivery_contact6);
            //Rec.InsertString("bl_reg_no", Record.bl_reg_no);
            Rec.InsertString("bl_bl_no", dr["bl_bl_no"].ToString());
            Rec.InsertDate("bl_bl_date", dr["hbl_house_date"]);
            //Rec.InsertString("bl_fcr_no", Record.bl_fcr_no);
            //Rec.InsertString("bl_fcr_doc1", Record.bl_fcr_doc1);
            //Rec.InsertString("bl_fcr_doc2", Record.bl_fcr_doc2);
            //Rec.InsertString("bl_fcr_doc3", Record.bl_fcr_doc3);
            Rec.InsertString("bl_vsl_name", dr["bl_vsl_name"].ToString());
            Rec.InsertString("bl_vsl_voy_no", dr["bl_vsl_voy_no"].ToString());
            //Rec.InsertString("bl_period_delivery", Record.bl_period_delivery);
            Rec.InsertString("bl_move_type", dr["bl_move_type"].ToString());


            //Rec.InsertString("bl_place_transhipment", Record.bl_place_transhipment);

            Rec.InsertNumeric("bl_pkg", Lib.Conv2Decimal(dr["bl_pkg"].ToString()).ToString());
            Rec.InsertString("bl_pkg_unit_id", dr["bl_pkg_unit_id"].ToString());
            Rec.InsertNumeric("bl_grwt", Lib.Conv2Decimal(dr["bl_grwt"].ToString()).ToString());
            Rec.InsertNumeric("bl_cbm", Lib.Conv2Decimal(dr["bl_cbm"].ToString()).ToString());
            Rec.InsertNumeric("bl_ntwt", Lib.Conv2Decimal(dr["bl_ntwt"].ToString()).ToString());
            Rec.InsertNumeric("bl_chwt", Lib.Conv2Decimal(dr["bl_chwt"].ToString()).ToString());
            Rec.InsertNumeric("bl_pcs", Lib.Conv2Decimal(dr["bl_pcs"].ToString()).ToString());

            //Rec.InsertString("bl_pcs_unit", Record.bl_pcs_unit);

            //Rec.InsertNumeric("bl_frt_amount", Lib.Conv2Decimal(Record.bl_frt_amount.ToString()).ToString());
            //Rec.InsertString("bl_frt_pay_at", Record.bl_frt_pay_at);
            Rec.InsertString("bl_issued_place", dr["bl_issued_place"].ToString());
            Rec.InsertDate("bl_issued_date", dr["bl_issued_date"].ToString());
            //Rec.InsertNumeric("bl_no_copies", Lib.Conv2Integer(Record.bl_no_copies.ToString()).ToString());
            //Rec.InsertString("bl_remarks1", Record.bl_remarks1);
            //Rec.InsertString("bl_remarks2", Record.bl_remarks2);
            //Rec.InsertString("bl_remarks3", Record.bl_remarks3);
            //Rec.InsertString("bl_remarks4", Record.bl_remarks4);
            //Rec.InsertString("bl_is_original", (Record.bl_is_original == true ? "Y" : "N"));
            //Rec.InsertString("bl_brazil_declaration", (Record.bl_brazil_declaration == true ? "Y" : "N"));
            //Rec.InsertString("bl_print_format_id", Record.bl_print_format_id);
            //Rec.InsertString("bl_iata_carrier", Record.bl_iata_carrier);
            //Rec.InsertString("bl_itm_po", Record.bl_itm_po);
            //Rec.InsertString("bl_itm_desc", Record.bl_itm_desc);

            Rec.InsertString("bl_mbl_no", dr["hbl_master_no"].ToString());
            Rec.InsertString("rec_created_by", user_code);
            Rec.InsertDate("rec_created_date", System.DateTime.Today);
            Rec.InsertString("rec_company_code", dr["rec_company_code"].ToString());
            Rec.InsertString("rec_branch_code", dr["rec_branch_code"].ToString());
            Rec.InsertString("rec_category", dr["rec_category"].ToString() + " EXPORT");

            sql = Rec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);

        }

        private void InsertHouseContainer(string parentid, string user_code)
        {
            sql = "select cntr_parent_id from imp_container where cntr_parent_id ='" + parentid + "'";
            if (Con_Oracle.IsRowExists(sql))
                return;

            foreach (DataRow dr in Dt_EdiContainer.Select("cntr_hbl_id='" + parentid + "'", "cntr_no"))
            {
                InsertContainer(dr, parentid, "H", user_code);
            }
        }

        private void InsertMasterContainer(string mblid, string user_code)
        {
            sql = "select cntr_parent_id from imp_container where cntr_parent_id ='" + mblid + "'";
            if (Con_Oracle.IsRowExists(sql))
                return;

            sql = " select '' as cntr_pkid, cntr_no,max(cntr_type_id) as cntr_size_id,max(cntrtype.param_code) as cntr_size_name";
            sql += "   ,max(cntr_csealno) as cntr_cseal,max(cntr_asealno) as cntr_aseal";
            sql += "   ,max(cntr_pkg_unit_id) as cntr_pkgs_unit_id,sum(cntr_pkg) as cntr_pkgs";
            sql += "   ,sum(cntr_pcs) as cntr_pcs,sum(cntr_ntwt) as cntr_ntwt,sum(cntr_grwt) as cntr_grwt";
            sql += "   ,sum(cntr_cbm) as cntr_cbm ,max(cntr_shipment_type) as cntr_type,max(a.rec_company_code) as rec_company_code,max(a.rec_branch_code) as rec_branch_code  ";
            sql += "   from imp_container a";
            sql += "   inner join bl b on a.cntr_parent_id=b.bl_pkid";
            sql += "   left join param cntrtype on a.cntr_type_id = cntrtype.param_pkid";
            sql += "   where bl_mbl_id='" + mblid + "'";
            sql += "   group by cntr_no";
            sql += "   order by cntr_no";
            DataTable Dt_cntr = new DataTable();
            Dt_cntr = Con_Oracle.ExecuteQuery(sql);

            foreach (DataRow dr in Dt_cntr.Rows)
            {
                dr["cntr_pkid"] = Guid.NewGuid().ToString().ToUpper();
                dr.AcceptChanges();
                InsertContainer(dr, mblid, "M", user_code);
            }
        }
        private void InsertContainer(DataRow dr, string parentid, string MorH, string user_code)
        {
            decimal teu = 0;

            if (dr["cntr_size_name"].ToString().Contains("20"))
                teu = 1;
            else if (dr["cntr_size_name"].ToString().Contains("40"))
            {
                if (dr["cntr_size_name"].ToString().Contains("HC"))
                    teu = (decimal)2.25;
                else
                    teu = 2;
            }
            else if (dr["cntr_size_name"].ToString().Contains("45"))
                teu = (decimal)2.50;
            else
                teu = 0;


            DBRecord Rec = new DBRecord();
            Rec.CreateRow("imp_container", "ADD", "cntr_pkid", dr["cntr_pkid"].ToString());
            Rec.InsertString("cntr_no", dr["cntr_no"].ToString());
            Rec.InsertNumeric("cntr_teu", teu.ToString());
            Rec.InsertString("cntr_parent_id", parentid);
            Rec.InsertString("cntr_morh", MorH);
            Rec.InsertString("cntr_type_id", dr["cntr_size_id"].ToString());
            Rec.InsertString("cntr_csealno", dr["cntr_cseal"].ToString());
            Rec.InsertString("cntr_asealno", dr["cntr_aseal"].ToString());
            Rec.InsertString("cntr_shipment_type", dr["cntr_type"].ToString());
            Rec.InsertString("cntr_pkg_unit_id", dr["cntr_pkgs_unit_id"].ToString());
            Rec.InsertNumeric("cntr_pkg", Lib.NumericFormat(dr["cntr_pkgs"].ToString(), 0));
            Rec.InsertNumeric("cntr_pcs", Lib.NumericFormat(dr["cntr_pcs"].ToString(), 3));
            Rec.InsertNumeric("cntr_ntwt", Lib.NumericFormat(dr["cntr_ntwt"].ToString(), 3));
            Rec.InsertNumeric("cntr_grwt", Lib.NumericFormat(dr["cntr_grwt"].ToString(), 3));
            Rec.InsertNumeric("cntr_cbm", Lib.NumericFormat(dr["cntr_cbm"].ToString(), 3));
            Rec.InsertString("rec_company_code", dr["rec_company_code"].ToString());
            Rec.InsertString("rec_branch_code", dr["rec_branch_code"].ToString());
            Rec.InsertString("rec_created_by", user_code);
            Rec.InsertDate("rec_created_date", System.DateTime.Today);

            sql = Rec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);

        }

        private void InsertHouseOrders(string parentid, string user_code)
        {
            sql = "select hord_hbl_id from blorder where hord_hbl_id ='" + parentid + "'";
            if (Con_Oracle.IsRowExists(sql))
                return;

            foreach(DataRow dr in Dt_EdiOrders.Select("ho_hbl_id='" + parentid + "'", "ho_ordno"))
            {
                InsertOrders(dr, parentid, user_code);
            }
        }

        private void InsertOrders(DataRow dr, string parentid, string user_code)
        {
  
            DBRecord Rec = new DBRecord();
            Rec.CreateRow("blorder", "ADD", "hord_pkid", dr["ho_pkid"].ToString());
            Rec.InsertString("hord_po", dr["ho_ordno"].ToString());
            Rec.InsertString("hord_style", dr["ho_style"].ToString());
            Rec.InsertString("hord_color", dr["ho_color"].ToString());
            Rec.InsertString("hord_hbl_id", parentid);
            Rec.InsertString("hord_cntr_id", dr["ho_cntr_id"].ToString());
            Rec.InsertString("hord_ord_id", dr["ho_ord_id"].ToString());  
            Rec.InsertString("hord_remarks", dr["ho_ord_remarks"].ToString());
            Rec.InsertString("hord_invno", dr["ho_invno"].ToString());
            Rec.InsertString("hord_pkg_unit", dr["ho_pkgs_unit"].ToString());
            Rec.InsertNumeric("hord_pkg", Lib.NumericFormat(dr["ho_pkgs"].ToString(), 0));
            Rec.InsertNumeric("hord_pcs", Lib.NumericFormat(dr["ho_pcs"].ToString(), 3));
            Rec.InsertNumeric("hord_grwt", Lib.NumericFormat(dr["ho_grwt"].ToString(), 3));
            Rec.InsertNumeric("hord_ntwt", Lib.NumericFormat(dr["ho_ntwt"].ToString(), 3));
            Rec.InsertNumeric("hord_cbm", Lib.NumericFormat(dr["ho_cbm"].ToString(), 3));
            Rec.InsertString("rec_company_code", dr["rec_company_code"].ToString());
            Rec.InsertString("rec_branch_code", dr["rec_branch_code"].ToString());
            Rec.InsertString("rec_created_by", user_code);
            Rec.InsertDate("rec_created_date", System.DateTime.Today);

            sql = Rec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);

        }

        private void InsertVesselDetails(string mblid, string hblid, string user_code)
        {
            sql = "select trk_parent_id from trackingm where trk_parent_id ='" + mblid + "'";
            if (Con_Oracle.IsRowExists(sql))
                return;

            int Tot_TrkSeq = 0;
            foreach (DataRow dr in Dt_EdiTracking.Select("trk_parent_id='" + hblid + "'", "trk_order"))
            {
                Tot_TrkSeq++;
                InsertTracking(dr, mblid, user_code);
            }

            sql = "update bl set bl_tot_tracking=" + Tot_TrkSeq.ToString() + " where bl_pkid='" + mblid + "'";
            Con_Oracle.ExecuteNonQuery(sql);
        }

        private void InsertTracking(DataRow dr, string parentid, string user_code)
        {
            DBRecord Rec = new DBRecord();
            Rec.CreateRow("trackingm", "ADD", "trk_pkid", dr["trk_pkid"].ToString());
            Rec.InsertString("trk_vsl_id", dr["trk_vsl_id"].ToString());
            Rec.InsertString("trk_voyage", dr["trk_voyage"].ToString());
            Rec.InsertString("trk_pol_id", dr["trk_pol_id"].ToString());
            Rec.InsertDate("trk_pol_etd", dr["trk_pol_etd"]);
            Rec.InsertString("trk_pol_etd_confirm", dr["trk_pol_etd_confirm"].ToString());
            Rec.InsertString("trk_pod_id ", dr["trk_pod_id"].ToString());
            Rec.InsertDate("trk_pod_eta", dr["trk_pod_eta"]);
            Rec.InsertString("trk_pod_eta_confirm", dr["trk_pod_eta_confirm"].ToString());
            Rec.InsertNumeric("trk_order", Lib.Conv2Integer(dr["trk_order"].ToString()).ToString());
            Rec.InsertString("trk_parent_id", dr["trk_parent_id"].ToString());
            Rec.InsertString("rec_company_code", dr["rec_company_code"].ToString());
            Rec.InsertString("rec_branch_code", dr["rec_branch_code"].ToString());
            Rec.InsertString("rec_created_by", user_code);
            Rec.InsertDate("rec_created_date", System.DateTime.Today);

            sql = Rec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);
        }


        private void InsertDescriptions(string hblid, string Category, string user_code)
        {
            sql = "select bl_parent_id from bldesc where bl_parent_id ='" + hblid + "'";
            if (Con_Oracle.IsRowExists(sql))
                return;

            List<Bldesc> mList = new List<Bldesc>();
            Bldesc dRow;
            for (int i = 0; i < 30; i++)
            {
                dRow = new Bldesc();
                dRow.bl_pkid = Guid.NewGuid().ToString().ToUpper();
                dRow.bl_parent_id = hblid;
                dRow.bl_parent_type = Category == "SEA" ? "SEAEXPDESC" : "AIREXPDESC";
                dRow.bl_desc_ctr = i + 1;
                dRow.bl_marks = "";
                dRow.bl_desc = "";
                dRow.bl_desc2 = "";
                mList.Add(dRow);
            }

            sql = "select hd_pkid,hd_hbl_id,hd_seq,hd_desc,hd_type from edi_house_desc ";
            sql += " where hd_type in ('1','2','3') and hd_hbl_id ='" + hblid + "' order by hd_type,hd_seq";
            DataTable Dt_Desc = new DataTable();
            Dt_Desc = Con_Oracle.ExecuteQuery(sql);

            int indx = 0;
            foreach (DataRow dr in Dt_Desc.Rows)
            {
                indx = Lib.Conv2Integer(dr["hd_seq"].ToString());
                if (indx < mList.Count)
                {
                    mList[indx].bl_desc_ctr = indx;
                    mList[indx].bl_pkid = dr["hd_pkid"].ToString();
                    mList[indx].bl_parent_id = dr["hd_hbl_id"].ToString();
                    if (dr["hd_type"].ToString() == "1")
                        mList[indx].bl_marks = dr["hd_desc"].ToString();
                    else if (dr["hd_type"].ToString() == "2")
                        mList[indx].bl_desc = dr["hd_desc"].ToString();
                    else if (dr["hd_type"].ToString() == "3")
                        mList[indx].bl_desc2 = dr["hd_desc"].ToString();
                }
            }

            DBRecord Rec;
            foreach (Bldesc Row in mList)
            {
                if (Row.bl_marks != "" || Row.bl_desc != "" || Row.bl_desc2 != "")
                {
                    Rec = new DBRecord();
                    Rec.CreateRow("Bldesc", "ADD", "bl_pkid", Row.bl_pkid);
                    Rec.InsertString("bl_parent_id", Row.bl_parent_id);
                    Rec.InsertString("bl_parent_type", Row.bl_parent_type);
                    Rec.InsertString("bl_marks", Row.bl_marks);
                    Rec.InsertString("bl_desc", Row.bl_desc);
                    Rec.InsertString("bl_desc2", Row.bl_desc2);
                    Rec.InsertNumeric("bl_desc_ctr", Row.bl_desc_ctr.ToString());
                    sql = Rec.UpdateRow();
                    Con_Oracle.ExecuteNonQuery(sql);
                }
            }
        }

        private string IsValid(DataRow Dr)
        {
            string Error = "";
            if (Dr["bl_agent_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "AGENT";
            if (Dr["bl_carrier_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "CARRIER";
            if (Dr["bl_shipper_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "SHIPPER";
            if (Dr["bl_consignee_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "CONSIGNEE";
            if (Dr["rec_branch_code"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "BRANCH";
            if (Dr["bl_pkg_unit_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "UNIT";
            foreach (DataRow dr in Dt_EdiContainer.Select("cntr_hbl_id='" + Dr["bl_pkid"].ToString() + "'", "cntr_no"))
            {
                if (dr["cntr_size_id"].Equals(DBNull.Value))
                    Error += ((Error != "") ? "," : "") + "CONTAINER TYPE";
                if (dr["cntr_pkgs_unit_id"].Equals(DBNull.Value))
                    Error += ((Error != "") ? "," : "") + "UNIT";
            }
            foreach (DataRow dr in Dt_EdiTracking.Select("trk_parent_id='" + Dr["bl_pkid"].ToString() + "'", "trk_order"))
            {
                if (dr["trk_vsl_id"].Equals(DBNull.Value))
                    Error += ((Error != "") ? "," : "") + "VESSEL";
                if (dr["trk_pol_id"].Equals(DBNull.Value))
                    Error += ((Error != "") ? "," : "") + "POL";
                if (dr["trk_pod_id"].Equals(DBNull.Value))
                    Error += ((Error != "") ? "," : "") + "POD";
            }

            if (Error != "")
                Error = "MISSING "+ Error;

            return Error;
        }

        private string CheckDuplication(DataRow Dr)
        {
            string sql = "";
            string Error = "";
            try
            {
                foreach (DataRow dr in Dt_EdiOrders.Select("ho_hbl_id='" + Dr["bl_pkid"].ToString() + "'", "ho_ordno"))
                {
                    sql = "";
                    sql += " select ord_pkid from joborderm where ";
                    sql += " rec_company_code = '" + Dr["REC_COMPANY_CODE"].ToString() + "'";
                    sql += " and ord_agent_id = '" + Dr["bl_agent_id"].ToString() + "'";
                    sql += " and ord_exp_id = '" + Dr["bl_shipper_id"].ToString() + "'";
                    sql += " and ord_imp_id = '" + Dr["bl_consignee_id"].ToString() + "'";
                    sql += " and ord_po = '" + dr["ho_ordno"].ToString() + "'";

                    if (dr["ho_style"].Equals(DBNull.Value))
                        sql += " and ord_style is null";
                    else
                        sql += " and ord_style = '" + dr["ho_style"].ToString() + "'";

                    if (dr["ho_color"].Equals(DBNull.Value))
                        sql += " and ord_color is null";
                    else
                        sql += " and ord_color = '" + dr["ho_color"].ToString() + "'";

                    DataTable Dt_Test = new DataTable();
                    Dt_Test = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Test.Rows.Count > 0)
                    {
                        dr["ho_ord_remarks"] = "";
                        dr["ho_ord_id"] = Dt_Test.Rows[0]["ord_pkid"].ToString();
                    }
                    else
                    {
                        dr["ho_ord_remarks"] = "ORDER NOT FOUND";
                        dr["ho_ord_id"] = "";
                    }
                }
            }
            catch( Exception Ex  )
            {
                Error = Ex.Message.ToString();

            }
            return Error;
        }
    }
}
