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
    public class EdiOrderService : BL_Base
    {

        string JOBORDERID = "";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string type = SearchData["type"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string User_Code = SearchData["user_code"].ToString();
            string partnerid = SearchData["partnerid"].ToString();
            string rowstatus = SearchData["rowstatus"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string ordstatus = SearchData["ordstatus"].ToString();
            string fileno = SearchData["fileno"].ToString();
            Boolean showdeleted = (Boolean) SearchData["showdeleted"];

            string ord_po = "";
            
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            Con_Oracle = new DBConnection();
            List<edi_order> mList = new List<edi_order>();
            edi_order mRow;
            string sWhere = "";


            if (SearchData.ContainsKey("po"))
            {
                ord_po = SearchData["po"].ToString();
                ord_po = ord_po.Replace(" ", "");
                ord_po = ord_po.Replace(",", "','");
            }
            

            try
            {
                sWhere = "";
                sWhere = "where a.rec_company_code = '{COMPCODE}'  ";

                if (partnerid != "ALL")
                    sWhere += " and a.ord_sender = '" + partnerid + "'";

                if (ordstatus != "ALL")
                    sWhere += " and a.ord_status = '" + ordstatus + "'";

                if ( rowstatus != "ALL")
                    sWhere += " and a.ord_updated = '" + rowstatus + "'";

                if ( showdeleted)
                    sWhere += " and a.rec_deleted = 'Y'";
                else
                    sWhere += " and a.rec_deleted = 'N'";

                if (fileno.Length > 0)
                    sWhere += " and b.messagenumber = '" + fileno + "'";

                if (ord_po.Length > 0)
                    sWhere += " and a.ord_po in ('" + ord_po + "')";

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM edi_order a inner join edi_header b on a.ord_headerid = b.headerid ";
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
                sql += " select b.slno, ord_pkid,ord_headerid,ord_sender,messagedate,messagefilename,messagenumber, ";
                sql += " ord_status,ord_updated, ord_pol_agent,ord_exp_name, ord_imp_name, rec_category, ord_pol, ord_pod,ord_invno,";
                sql += " ord_po, ord_style, ord_color, ord_boarding1, ord_boarding2, ord_instock1, ord_instock2, a.rec_deleted, transfer_remarks,";
                sql += " row_number() over(order by slno,ord_sender,ord_po ) rn";
                sql += " from edi_order a ";
                sql += " inner join edi_header b on a.ord_headerid = b.headerid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by slno,ord_sender,ord_po ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new edi_order();
                    mRow.ord_pkid = Dr["ord_pkid"].ToString();

                    mRow.ord_header_id = Dr["ord_headerid"].ToString();
                    mRow.ord_sender = Dr["ord_sender"].ToString();
                    mRow.ord_message_date = Dr["messagedate"].ToString();
                    mRow.ord_message_file_name = Dr["messagefilename"].ToString();
                    mRow.ord_message_number = Dr["messagenumber"].ToString();

                    mRow.ord_status = Dr["ord_status"].ToString();
                    mRow.ord_updated = Dr["ord_updated"].ToString();

                    mRow.ord_invno = Dr["ord_invno"].ToString();
                    mRow.ord_po = Dr["ord_po"].ToString();
                    mRow.ord_style = Dr["ord_style"].ToString();
                    mRow.ord_color = Dr["ord_color"].ToString();
                    mRow.ord_pol_agent = Dr["ord_pol_agent"].ToString();
                    mRow.ord_exp_name = Dr["ord_exp_name"].ToString();
                    mRow.ord_imp_name = Dr["ord_imp_name"].ToString();
                    mRow.ord_pol = Dr["ord_pol"].ToString();
                    mRow.ord_pod = Dr["ord_pod"].ToString();
                    
                    mRow.ord_boarding1 = Lib.DatetoString(Dr["ord_boarding1"]);
                    mRow.ord_boarding2 = Lib.DatetoString(Dr["ord_boarding2"]);
                    mRow.ord_instock1 = Lib.DatetoString(Dr["ord_instock1"]);
                    mRow.ord_instock2 = Lib.DatetoString(Dr["ord_instock2"]);

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
                sql += " select distinct ord_sender as sender, 'SHIPPER' as type, source , link_target_id as target_id, link_target_name as target_name ";
                sql += " from( ";
                sql += " select ord_sender, a.ord_exp_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_order a ";
                sql += " inner ";
                sql += " join edi_header b on a.ord_headerid = b.headerid ";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'SHIPPER' and  a.rec_company_code = c.rec_company_code and a.ord_sender = c.link_messagesender and a.ord_exp_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.ord_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";
                sql += " union all ";
                sql += " select distinct  ord_sender,'CONSIGNEE' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  ord_sender,a.ord_imp_name as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_order a ";
                sql += " inner join edi_header b on a.ord_headerid = b.headerid ";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'CONSIGNEE' and  a.rec_company_code = c.rec_company_code and a.ord_sender = c.link_messagesender and a.ord_imp_name = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.ord_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";
                sql += " union all ";
                sql += " select distinct  ord_sender,'POL' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  ord_sender,a.ord_pol as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_order a ";
                sql += " inner join edi_header b on a.ord_headerid = b.headerid ";
                sql += " left ";
                sql += " join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'PORT' and  a.rec_company_code = c.rec_company_code and a.ord_sender = c.link_messagesender and a.ord_pol = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.ord_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";
                sql += " union all ";
                sql += " select distinct  ord_sender,'POD' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  ord_sender,a.ord_pod as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_order a ";
                sql += " inner join edi_header b on a.ord_headerid = b.headerid ";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'PORT' and  a.rec_company_code = c.rec_company_code and a.ord_sender = c.link_messagesender and a.ord_pod = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.ord_updated = 'N' and a.rec_deleted = 'N' ";
                sql += " ) a where link_target_id is null ";
                sql += " union all ";
                sql += " select distinct  ord_sender,'POL-AGENT' as type, source , link_target_id, link_target_name ";
                sql += " from( ";
                sql += " select  ord_sender,a.ord_pol_agent as source, c.link_target_id, c.link_target_name ";
                sql += " from edi_order a ";
                sql += " inner join edi_header b on a.ord_headerid = b.headerid ";
                sql += " left join edi_link c on c.link_type = 'INWARD' and c.link_subcategory = 'POL-AGENT' and  a.rec_company_code = c.rec_company_code and a.ord_sender = c.link_messagesender and a.ord_pol_agent = c.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}' and  a.ord_updated = 'N' and a.rec_deleted = 'N' ";
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


        public IDictionary<string, object> TransferPO(Dictionary<string, object> SearchData)
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

            DBRecord mRec = null;

            try
            {

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select ord_pkid, ord_sender, messagenumber,ord_status, a.rec_company_code, a.rec_category, ";
                sql += " a.ord_pol_agent, agent_link.link_target_id as agent_link_id, agent_link.link_target_name as agent_link_name, ";
                sql += " a.ord_exp_name , shpr_link.link_target_id as shpr_link_id, shpr_link.link_target_name as shpr_link_name, ";
                sql += " a.ord_imp_name , cons_link.link_target_id as cons_link_id, cons_link.link_target_name as cons_link_name, ";
                sql += " a.ord_pol , pol_link.link_target_id as pol_link_id, pol_link.link_target_name as pol_link_name, ";
                sql += " a.ord_pod , pod_link.link_target_id as pod_link_id, pod_link.link_target_name as pod_link_name, ";
                sql += " ord_uneco,  ord_invno,  ord_po,  ord_style,  ord_color,  ord_contractno,  ord_pkg,  ord_grwt,  ord_ntwt, ";
                sql += " ord_pcs, ord_cbm, ord_hs_code, ord_desc, ord_boarding1, ord_boarding2, ord_instock1, ord_instock2, ord_updated,  a.rec_deleted, ";
                sql += " transfer_remarks, ord_booking_date, ord_rnd_insp_date, ord_po_rel_date, ord_cargo_ready_date, ord_fcr_date, ord_insp_date, ord_stuf_date, ord_whd_date ";
                sql += " from edi_order a  inner join edi_header b on a.ord_headerid = b.headerid  ";
                sql += " left join edi_link agent_link on agent_link.link_type = 'INWARD' and agent_link.link_subcategory = 'POL-AGENT' and a.rec_company_code = agent_link.rec_company_code and a.ord_sender = agent_link.link_messagesender and a.ord_pol_agent = agent_link.link_source_name ";
                sql += " left join edi_link shpr_link  on shpr_link.link_type = 'INWARD' and shpr_link.link_subcategory = 'SHIPPER' and a.rec_company_code = shpr_link.rec_company_code and a.ord_sender = shpr_link.link_messagesender   and a.ord_exp_name = shpr_link.link_source_name ";
                sql += " left join edi_link cons_link  on cons_link.link_type = 'INWARD' and cons_link.link_subcategory = 'CONSIGNEE' and a.rec_company_code = cons_link.rec_company_code and a.ord_sender = cons_link.link_messagesender   and a.ord_imp_name = cons_link.link_source_name ";
                sql += " left join edi_link pol_link   on pol_link.link_type = 'INWARD' and pol_link.link_subcategory = 'PORT' and a.rec_company_code = pol_link.rec_company_code and a.ord_sender = pol_link.link_messagesender     and a.ord_pol = pol_link.link_source_name ";
                sql += " left join edi_link pod_link   on pod_link.link_type = 'INWARD' and pod_link.link_subcategory = 'PORT' and a.rec_company_code = pod_link.rec_company_code and a.ord_sender = pod_link.link_messagesender     and a.ord_pod = pod_link.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}'  and ord_status = 'PO' ";
                if (sender != "ALL")
                    sql += " and a.ord_sender = '{SENDER}' ";
                sql += " and a.ord_updated = 'N' and a.rec_deleted = 'N' ";

                sql += " order by slno ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{SENDER}", sender);

                Dt_List = Con_Oracle.ExecuteQuery(sql);

                foreach ( DataRow Dr in Dt_List.Rows)
                {
                    Remarks  = "";
                    SaveOk = false;
                    Con_Oracle.BeginTransaction();
                    if (Remarks == "")
                        Remarks = IsValid(Dr);
                    if (Remarks == "")
                        Remarks = CheckDuplication(Dr);
                    if (Remarks == "")
                    {
                        mRec = new DBRecord();
                        mRec.CreateRow("joborderm", "ADD", "ORD_PKID", Dr["ORD_PKID"].ToString());
                        mRec.InsertString("ORD_STATUS", "REPORTED");
                        mRec.InsertString("ORD_AGENT_ID", Dr["AGENT_LINK_ID"].ToString());
                        mRec.InsertString("ORD_AGENT_NAME", Dr["AGENT_LINK_NAME"].ToString());
                        mRec.InsertString("ORD_EXP_ID", Dr["SHPR_LINK_ID"].ToString());
                        mRec.InsertString("ORD_EXP_NAME", Dr["SHPR_LINK_NAME"].ToString());
                        mRec.InsertString("ORD_IMP_ID", Dr["CONS_LINK_ID"].ToString());
                        mRec.InsertString("ORD_IMP_NAME", Dr["CONS_LINK_NAME"].ToString());
                        mRec.InsertString("ORD_APPROVED", "N");
                        mRec.InsertString("ORD_POL_ID", Dr["POL_LINK_ID"].ToString());
                        mRec.InsertString("ORD_POL", Dr["POL_LINK_NAME"].ToString());
                        mRec.InsertString("ORD_POD_ID", Dr["POD_LINK_ID"].ToString());
                        mRec.InsertString("ORD_POD", Dr["POD_LINK_NAME"].ToString());
                        mRec.InsertString("ORD_SOURCE", "EDI");
                        mRec.InsertString("ORD_INVNO", Dr["ORD_INVNO"].ToString());
                        mRec.InsertString("ORD_UNECO", Dr["ORD_UNECO"].ToString());
                        mRec.InsertString("ORD_PO", Dr["ORD_PO"].ToString());
                        mRec.InsertString("ORD_STYLE", Dr["ORD_STYLE"].ToString());
                        mRec.InsertString("ORD_COLOR", Dr["ORD_COLOR"].ToString());

                        mRec.InsertNumeric("ORD_CBM", Dr["ORD_CBM"].ToString());
                        mRec.InsertNumeric("ORD_PCS", Dr["ORD_PCS"].ToString());
                        mRec.InsertNumeric("ORD_PKG", Dr["ORD_PKG"].ToString());
                        mRec.InsertNumeric("ORD_GRWT", Dr["ORD_GRWT"].ToString());
                        mRec.InsertNumeric("ORD_NTWT", Dr["ORD_NTWT"].ToString());

                        mRec.InsertString("ORD_HS_CODE", Dr["ORD_HS_CODE"].ToString());
                        mRec.InsertString("ORD_DESC", Dr["ORD_DESC"].ToString());

                        mRec.InsertDate("ORD_BOARDING1", Dr["ORD_BOARDING1"]);
                        mRec.InsertDate("ORD_BOARDING2", Dr["ORD_BOARDING2"]);
                        mRec.InsertDate("ORD_INSTOCK1", Dr["ORD_INSTOCK1"]);
                        mRec.InsertDate("ORD_INSTOCK2", Dr["ORD_INSTOCK2"]);
                        mRec.InsertDate("ORD_BOOKING_DATE", Dr["ORD_BOOKING_DATE"]);
                        mRec.InsertDate("ORD_RND_INSP_DATE", Dr["ORD_RND_INSP_DATE"]);
                        mRec.InsertDate("ORD_PO_REL_DATE", Dr["ORD_PO_REL_DATE"]);
                        mRec.InsertDate("ORD_CARGO_READY_DATE", Dr["ORD_CARGO_READY_DATE"]);
                        mRec.InsertDate("ORD_FCR_DATE", Dr["ORD_FCR_DATE"]);
                        mRec.InsertDate("ORD_INSP_DATE", Dr["ORD_INSP_DATE"]);
                        mRec.InsertDate("ORD_STUF_DATE", Dr["ORD_STUF_DATE"]);
                        mRec.InsertDate("ORD_WHD_DATE", Dr["ORD_WHD_DATE"]);

                        mRec.InsertString("REC_CREATED_BY", user_code);
                        mRec.InsertDate("REC_CREATED_DATE", System.DateTime.Today);
                        mRec.InsertString("REC_COMPANY_CODE", Dr["REC_COMPANY_CODE"].ToString());
                        mRec.InsertString("REC_CATEGORY", Dr["REC_CATEGORY"].ToString() + " EXPORT");

                        sql = mRec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);

                        Remarks = "TRANSFERED";
                        SaveOk = true;
                    }
                    if (Remarks != "")
                    {
                        Remarks = Remarks.Length > 100 ? Remarks.Substring(0, 100) : Remarks;
                        sql = "update edi_order set transfer_remarks = '" + Remarks + "' ";
                        if (SaveOk)
                        {
                            sql += ",ord_order_id = '" + Dr["ORD_PKID"].ToString() + "'";
                            sql += ",ord_updated = 'Y' ";
                            processed++;
                        }
                        else
                        {
                            notprocessed++;
                        }
                        sql += " where ord_pkid = '" + Dr["ord_pkid"].ToString() + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                    Con_Oracle.CommitTransaction();

                    if (SaveOk)
                    {
                        string srem = "EDI FILE# " + Dr["messagenumber"].ToString();
                        Lib.AuditLog("PO", "PO", "EDI ADD", company_code, "", user_code, Dr["ord_pkid"].ToString(), Dr["ORD_PO"].ToString(), srem);
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


        public IDictionary<string, object> TransferPOTracking(Dictionary<string, object> SearchData)
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

            DBRecord mRec = null;

            try
            {

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select ord_pkid, ord_sender, ord_status, messagenumber, a.rec_company_code, a.rec_category, ";
                sql += " a.ord_pol_agent, agent_link.link_target_id as agent_link_id, agent_link.link_target_name as agent_link_name, ";
                sql += " a.ord_exp_name , shpr_link.link_target_id as shpr_link_id, shpr_link.link_target_name as shpr_link_name, ";
                sql += " a.ord_imp_name , cons_link.link_target_id as cons_link_id, cons_link.link_target_name as cons_link_name, ";
                sql += " a.ord_pol , pol_link.link_target_id as pol_link_id, pol_link.link_target_name as pol_link_name, ";
                sql += " a.ord_pod , pod_link.link_target_id as pod_link_id, pod_link.link_target_name as pod_link_name, ";
                sql += " ord_uneco,  ord_invno,  ord_po,  ord_style,  ord_color,  ord_contractno,  ord_pkg,  ord_grwt,  ord_ntwt, ";
                sql += " ord_pcs, ord_cbm, ord_hs_code, ord_desc, ord_boarding1, ord_boarding2, ord_instock1, ord_instock2, ord_updated,  a.rec_deleted, ";
                sql += " transfer_remarks, ord_booking_date, ord_rnd_insp_date, ord_po_rel_date, ord_cargo_ready_date, ord_fcr_date, ord_insp_date, ord_stuf_date, ord_whd_date ";
                sql += " from edi_order a  inner join edi_header b on a.ord_headerid = b.headerid  ";
                sql += " left join edi_link agent_link on agent_link.link_type = 'INWARD' and agent_link.link_subcategory = 'POL-AGENT' and a.rec_company_code = agent_link.rec_company_code and a.ord_sender = agent_link.link_messagesender and a.ord_pol_agent = agent_link.link_source_name ";
                sql += " left join edi_link shpr_link  on shpr_link.link_type = 'INWARD' and shpr_link.link_subcategory = 'SHIPPER' and a.rec_company_code = shpr_link.rec_company_code and a.ord_sender = shpr_link.link_messagesender   and a.ord_exp_name = shpr_link.link_source_name ";
                sql += " left join edi_link cons_link  on cons_link.link_type = 'INWARD' and cons_link.link_subcategory = 'CONSIGNEE' and a.rec_company_code = cons_link.rec_company_code and a.ord_sender = cons_link.link_messagesender   and a.ord_imp_name = cons_link.link_source_name ";
                sql += " left join edi_link pol_link   on pol_link.link_type = 'INWARD' and pol_link.link_subcategory = 'PORT' and a.rec_company_code = pol_link.rec_company_code and a.ord_sender = pol_link.link_messagesender     and a.ord_pol = pol_link.link_source_name ";
                sql += " left join edi_link pod_link   on pod_link.link_type = 'INWARD' and pod_link.link_subcategory = 'PORT' and a.rec_company_code = pod_link.rec_company_code and a.ord_sender = pod_link.link_messagesender     and a.ord_pod = pod_link.link_source_name ";
                sql += " where a.rec_company_code = '{COMPCODE}'  and ord_status = 'PO TRACKING' ";
                if (sender != "ALL")
                    sql += " and a.ord_sender = '{SENDER}' ";
                sql += " and a.ord_updated = 'N' and a.rec_deleted = 'N' ";

                sql += " order by slno ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{SENDER}", sender);

                Dt_List = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    SaveOk = false;
                    Remarks = "";
                    Con_Oracle.BeginTransaction();
                    if (Remarks == "")
                        Remarks = IsValid(Dr);
                    if (Remarks == "")
                        Remarks = CheckDuplication(Dr);
                    if (Remarks == "PO EXISTS")
                    {
                        mRec = new DBRecord();
                        mRec.CreateRow("joborderm", "EDIT", "ORD_PKID", JOBORDERID);
                        mRec.InsertDate("ORD_BOOKING_DATE", Dr["ORD_BOOKING_DATE"]);
                        mRec.InsertDate("ORD_RND_INSP_DATE", Dr["ORD_RND_INSP_DATE"]);
                        mRec.InsertDate("ORD_PO_REL_DATE", Dr["ORD_PO_REL_DATE"]);
                        mRec.InsertDate("ORD_CARGO_READY_DATE", Dr["ORD_CARGO_READY_DATE"]);
                        mRec.InsertDate("ORD_FCR_DATE", Dr["ORD_FCR_DATE"]);
                        mRec.InsertDate("ORD_INSP_DATE", Dr["ORD_INSP_DATE"]);
                        mRec.InsertDate("ORD_STUF_DATE", Dr["ORD_STUF_DATE"]);
                        mRec.InsertDate("ORD_WHD_DATE", Dr["ORD_WHD_DATE"]);

                        mRec.InsertString("REC_EDITED_BY", user_code);
                        mRec.InsertDate("REC_EDITED_DATE", System.DateTime.Today);

                        sql = mRec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);
                        Remarks = "TRACKING UPDATED";
                        SaveOk = true;
                    }
                    if (Remarks != "")
                    {
                        Remarks = Remarks.Length > 100 ? Remarks.Substring(0, 100) : Remarks;
                        sql = "update edi_order set transfer_remarks = '" + Remarks + "' ";
                        if (SaveOk)
                        {
                            sql += ",ord_order_id = '" + JOBORDERID + "'";
                            sql += ",ord_updated = 'Y' ";
                            processed++;
                        }
                        else
                        {
                            notprocessed++;
                        }
                        sql += " where ord_pkid = '" + Dr["ord_pkid"].ToString() + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                    Con_Oracle.CommitTransaction();

                    if (SaveOk)
                    {
                        string srem = "EDI FILE# " + Dr["messagenumber"].ToString();
                        Lib.AuditLog("PO", "PO", "EDI TRACKING", company_code, "", user_code, JOBORDERID, Dr["ORD_PO"].ToString(), srem);
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


        private string IsValid(DataRow Dr)
        {
            string Error = "";
            if (Dr["agent_link_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") +  "AGENT";
            if (Dr["shpr_link_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "SHIPPER";
            if (Dr["cons_link_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "CONSIGNEE";
            if (Dr["pol_link_id"].Equals(DBNull.Value) )
                Error += ((Error != "") ? "," : "") + "POL";
            if (Dr["pod_link_id"].Equals(DBNull.Value))
                Error += ((Error != "") ? "," : "") + "POD";
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
                JOBORDERID = "";
                sql = "";
                sql += " select ord_pkid from joborderm where ";
                sql += " rec_company_code = '" + Dr["REC_COMPANY_CODE"].ToString() + "'";
                sql += " and ord_agent_id = '" + Dr["agent_link_id"].ToString() + "'";
                sql += " and ord_exp_id = '" + Dr["shpr_link_id"].ToString() + "'";
                sql += " and ord_imp_id = '" + Dr["cons_link_id"].ToString() + "'";
                sql += " and ord_po = '" + Dr["ord_po"].ToString() + "'";

                if (Dr["ord_style"].Equals(DBNull.Value))
                    sql += " and ord_style is null";
                else
                    sql += " and ord_style = '" + Dr["ord_style"].ToString() + "'";

                if (Dr["ord_color"].Equals(DBNull.Value))
                    sql += " and ord_color is null";
                else
                    sql += " and ord_color = '" + Dr["ord_color"].ToString() + "'";

                DataTable Dt_Test = new DataTable();
                Dt_Test = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Test.Rows.Count > 0)
                {
                    Error = "PO EXISTS";
                    JOBORDERID = Dt_Test.Rows[0]["ord_pkid"].ToString();
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
