using System;
using System.Data;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using System.Text;

namespace BLXml 
{
    public class MexicoOrdersCsv : BaseReport
    {
        private DataTable DT_ORDER = new DataTable();
        public Boolean IsError = false;
        public string ErrorMessage = "";
        private string sql = "";
        public string PKID = "";
        public string InvokeType = "";
        public string File_Name = "";
        public int ProcessOrdCount = 0;
        public string File_Subject = "";
        public string Ftp_updtsql = "";
        DBConnection Con_Oracle = null;
        StringBuilder sb = new StringBuilder();
        public void Generate()
        {
            try
            {
                ErrorMessage = "";
                IsError = false;
                PKID = PKID.Replace(",", "','");
                ReadData();

                if (IsError)
                {
                    return;
                }

                GenerateCsvFiles();
                WriteCsvFiles();
                DT_ORDER.Rows.Clear();
            }
            catch (Exception ex)
            {
                IsError = true;
                ErrorMessage += " |" + ex.Message.ToString();
            }
        }
        private void ReadData()
        {

            Con_Oracle = new DBConnection();

            sql = " select ord_pkid ";
            sql += " ,'' as POID";
            // sql += " ,case when nvl(ord_approved,'N')='Y' then 'APPROVED' else 'REPORTED' end as STATUS";
            sql += " ,nvl(ord_status,'REPORTED') as STATUS";
            sql += " ,ord_uneco as DIVISION";
            sql += " ,ord_style as MODEL_SKU";
            sql += " ,ord_po as PO";
            sql += " ,ord_po as SUPPLIER_PO";
            sql += " ,'' as PO_BEFORE ";
            sql += " ,imp.targetid as CONSIGNEE";
            sql += " ,cnge.cust_name as CONSIGNEE_NAME";
            sql += " ,nvl(orgcntry.param_code,'IN') as ORIGIN_COUNTRY";
            sql += " ,nvl(ord_pol,pol.param_code) as POL";
            sql += " ,nvl(ord_pod,pod.param_code) as POD";
            sql += " ,case when ord.rec_category='SEA EXPORT' then 'SEA' else 'AIR' end as TRANSPORT_WAY";
            sql += " ,exp.targetid as SUPPLIER_ID";
            sql += " ,ord_desc as CARGO_DESCRIPTION";
            sql += " ,'' as WINDOW_OF_BOARDING1";
            sql += " ,'' as WINDOW_OF_BOARDING2";
            sql += " ,'' as IN_STOCK1";
            sql += " ,'' as IN_STOCK2";
            sql += " ,'' as FACTORY";
            sql += " ,'' as SHIPPER";
            sql += " ,'' as INCOTERM";
            sql += " ,'' as IMPORT_EXECUTIVE";
            sql += " ,ord_agentref_id as AGENT_REFERENCE_NUMBER";
            sql += " from joborderm ord"; 
            sql += " left join linkm2 imp on ord.ord_imp_id = imp.sourceid and imp.sourcetable='MEXICO-TMM' and imp.sourcetype='CONSIGNEE'";
            sql += " left join linkm2 exp on ord.ord_exp_id = exp.sourceid and exp.sourcetable='MEXICO-TMM' and exp.sourcetype='SHIPPER'";
            //sql += " left join customerm imp on ord.ord_imp_id = imp.cust_pkid";
            sql += " left join jobm job on  ord.ord_parent_id =  job.job_pkid";
            sql += " left join param orgcntry on job.job_origin_country_id = orgcntry.param_pkid";
            sql += " left join param pol on job.job_pol_id = pol.param_pkid";
            sql += " left join param pod on job.job_pod_id = pod.param_pkid";
            sql += " left join customerm shpr on ord.ord_exp_id = shpr.cust_pkid";
            sql += " left join customerm cnge on ord.ord_imp_id = cnge.cust_pkid";
            if (InvokeType == "AGENTBOOKING")
                sql += " where  ord_booking_id ='" + PKID + "'";
            else
                sql += " where  ord_pkid in ('" + PKID + "')";
            sql += " and nvl(ord_status,'REPORTED') = 'REPORTED' ";
            // sql += " order by case when nvl(ord_approved,'N')='Y' then 'B' else 'A' end,shpr.cust_name,ord_po ";
            sql += " order by ord_status,shpr.cust_name,ord_po ";
            DT_ORDER = new DataTable();
            DT_ORDER = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();

            ProcessOrdCount = DT_ORDER.Rows.Count;
            if (DT_ORDER.Rows.Count <= 0)
            {
                IsError = true;
                ErrorMessage += " | Details not Found";
            }
            File_Subject = "PO-";
            string str = "";
            string ord_pkid = "";
            foreach (DataRow dr in DT_ORDER.Rows)
            {
                if (str != "")
                    str += ",";
                str += dr["PO"].ToString();

                if (ord_pkid != "")
                    ord_pkid += ",";
                ord_pkid += dr["ord_pkid"].ToString();

                if (dr["PO"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " |  PO Cannot be Blank " ;
                }
                //if (dr["SUPPLIER_ID"].ToString().Length <= 0)
                //{
                //    IsError = true;
                //    ErrorMessage += " | Please link Shipper, " + dr["SHIPPER"].ToString() + "  for PO " + dr["PO"].ToString();
                //}
                if (dr["CONSIGNEE"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Please link Consinee, "+ dr["CONSIGNEE_NAME"].ToString() + " for PO " + dr["PO"].ToString();
                }
                if (dr["POL"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | POL Cannot be Blank for PO " + dr["PO"].ToString();
                }
                if (dr["POD"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | POD Cannot be Blank for PO " + dr["PO"].ToString();
                }
                //if (dr["SUPPLIER_ID"].ToString().Length <= 0)
                //{
                //    IsError = true;
                //    ErrorMessage += " | SUPPLIER ID Cannot be Blank for PO " + dr["PO"].ToString();
                //}
                if (dr["AGENT_REFERENCE_NUMBER"].ToString().Length <= 0 || dr["AGENT_REFERENCE_NUMBER"].ToString().Length > 10)
                {
                    IsError = true;
                    ErrorMessage += " | Invalid Agent Reference Number for PO " + dr["PO"].ToString();
                }

            }
            File_Subject += str;

            Ftp_updtsql = "";
            if (ord_pkid != "")
            {
                if (ord_pkid.Contains(","))
                    ord_pkid = ord_pkid.Replace(",", "','");
                Ftp_updtsql = "update joborderm set ord_ftp_status ='{STATUS}' where ord_pkid in ('" + ord_pkid + "')";
            }
        }

        private void GenerateCsvFiles()
        {
            sb = new StringBuilder();

            sb.Append("ID PO"); sb.Append(",");
            sb.Append("STATUS"); sb.Append(",");
            sb.Append("DIVISION"); sb.Append(",");
            sb.Append("MODEL/ SKU"); sb.Append(",");
            sb.Append("PO"); sb.Append(",");
            sb.Append("SUPPLIER PO"); sb.Append(",");
            sb.Append("PO BEFORE"); sb.Append(",");
            sb.Append("CONSIGNEE"); sb.Append(",");
            sb.Append("ORIGIN COUNTRY"); sb.Append(",");
            sb.Append("POL"); sb.Append(",");
            sb.Append("POD"); sb.Append(",");
            sb.Append("TRANSPORT WAY"); sb.Append(",");
            sb.Append("SUPPLIER ID"); sb.Append(",");
            sb.Append("CARGO DESCRIPTION"); sb.Append(",");
            sb.Append("WINDOW OF BOARDING"); sb.Append(",");
            sb.Append("WINDOW OF BOARDING"); sb.Append(",");
            sb.Append("IN STOCK"); sb.Append(",");
            sb.Append("IN STOCK"); sb.Append(",");
            sb.Append("FACTORY"); sb.Append(",");
            sb.Append("SHIPPER"); sb.Append(",");
            sb.Append("INCOTERM"); sb.Append(",");
            sb.Append("IMPORT EXECUTIVE"); sb.Append(",");
            sb.Append("AGENT REFERENCE NUMBER"); 
            foreach (DataRow dr in DT_ORDER.Rows)
            {
                dr["TRANSPORT_WAY"] = "SEA";

                sb.AppendLine();
                sb.Append(Lib.GetTruncated(dr["POID"].ToString(), 11)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["STATUS"].ToString(), 10)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["DIVISION"].ToString(), 50)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["MODEL_SKU"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["PO"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["SUPPLIER_PO"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["PO_BEFORE"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["CONSIGNEE"].ToString().Replace(",", " "), 3)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["ORIGIN_COUNTRY"].ToString(), 2)); sb.Append(",");
                sb.Append(Lib.GetPortCode(dr["POL"].ToString())); sb.Append(",");
                sb.Append(Lib.GetPortCode(dr["POD"].ToString())); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["TRANSPORT_WAY"].ToString(), 3)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["SUPPLIER_ID"].ToString(), 5)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["CARGO_DESCRIPTION"].ToString().Replace(",", " "), 100)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["WINDOW_OF_BOARDING1"].ToString(), 10)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["WINDOW_OF_BOARDING2"].ToString(), 10)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["IN_STOCK1"].ToString(), 10)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["IN_STOCK2"].ToString(), 10)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["FACTORY"].ToString(), 50)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["SHIPPER"].ToString().Replace(",", " "), 50)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["INCOTERM"].ToString(), 10)); sb.Append(",");
                sb.Append(Lib.GetTruncated(dr["IMPORT_EXECUTIVE"].ToString(), 50)); sb.Append(",");
                sb.Append(dr["AGENT_REFERENCE_NUMBER"].ToString());
            }

        }



        private void WriteCsvFiles()
        {
            try
            {
                if (sb == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | Cargo Orders Not Generated.";
                    return;
                }

                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                System.IO.File.AppendAllText(File_Name,sb.ToString());
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }
    }
}
