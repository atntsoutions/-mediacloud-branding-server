using System;
using System.Data;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.MexicoAck;

namespace BLXml
{
    public class MexicoAckRpt : BaseReport
    {
        private DataTable DT_ACK = new DataTable();
        private DataTable DT_ORDER = new DataTable();
        
        public Boolean IsError = false;
        public string ErrorMessage = "";
        private MessageAck AckMessage = null;
        private string sql = "";
        public string PKID = "";
        public string File_Name = "";
        public string GenerateType = "";
        DBConnection Con_Oracle = null;
        public int ProcessCount = 1;

        public void Generate()
        {
            try
            {
                ErrorMessage = "";
                IsError = false;
                ReadData();
                //if (DT_ACK.Rows.Count <= 0)
                //{
                //    IsError = true;
                //    ErrorMessage = "Details not Found";
                //    return;
                //}

                IsError = false;
                GenerateXmlFiles();
                WriteXmlFiles();
                DT_ACK.Rows.Clear();
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

            //if (GenerateType == "CONTAINER")
            //{
            //    sql = "";
            //    sql = " Select  ";
            //    sql += "   ord_uid as poid";
            //    sql += "   ,ord_po as po";
            //    sql += "   ,shpr.cust_name as posupplier";
            //    sql += "   ,'' as oldpo";
            //    sql += "   ,ord_style as styleno";
            //    sql += "   ,ord_grwt as gw";
            //    sql += "   ,ord_pkg as packageno ";
            //    sql += "   ,ord_cbm as cbm";
            //    sql += "   ,ord_pcs as piecenumber";
            //    sql += "   ,ord_desc as Description";
            //    sql += "   ,ord_booking_date as BookingRequestDate";
            //    sql += "   ,ord_rnd_insp_date as RandomInspectionDate";
            //    sql += "   ,ord_po_rel_date as poreleasedate";
            //    sql += "   ,ord_cargo_ready_date as cargoreadydate";
            //    sql += "   ,ord_fcr_date as fcrdate";
            //    sql += "   ,ord_insp_date as inspectiondate";
            //    sql += "   ,ord_stuf_date as stuffingdate";
            //    sql += "   ,ord_whd_date as warehousedeparturedate";
            //    sql += "   from packingm pkg";
            //    sql += "   inner join joborderm ord on pkg.pack_job_id = ord.ord_parent_id";
            //    sql += "   left join customerm shpr on ord.ord_exp_id = shpr.cust_pkid";
            //    sql += "   where pkg.pack_cntr_id='" + PKID + "'";
            //    sql += "   order by shpr.cust_name,ord_po";
            //    DT_ORDER = new DataTable();
            //    DT_ORDER = Con_Oracle.ExecuteQuery(sql);
            //    ProcessCount = DT_ORDER.Rows.Count;
            //}
            DT_ACK = new DataTable();
            Con_Oracle.CloseConnection();

        }

        private void GenerateXmlFiles()
        {
            AckMessage = new MessageAck();
            AckMessage.ProcessID = XmlLib.PROCESSID;
            AckMessage.TotalProcessCount = ProcessCount.ToString();
            AckMessage.TotalFailCount = "0";
            AckMessage.FailData = Generate_FailData();
        }

        private MessageAckFailData[] Generate_FailData()
        {
            MessageAckFailData Rec = null;
            MessageAckFailData[] fList = null;

            MessageAckFailDataVSLInfo Rec2 = null;
            MessageAckFailDataVSLInfo[] vList = null;


            MessageAckFailDataOrdersOrder Rec3 = null;
            MessageAckFailDataOrdersOrder[] OList = null;

            int ArrIndex = 0;
            try
            {
                vList = new MessageAckFailDataVSLInfo[1];
                Rec2 = new MessageAckFailDataVSLInfo();
                Rec2.VSLKey = "";
                Rec2.VSLFailReason = "";
                vList[0] = Rec2;

                OList = new MessageAckFailDataOrdersOrder[1];
                Rec3 = new MessageAckFailDataOrdersOrder();
                Rec3.PO = "";
                Rec3.StyleNo = "";
                Rec3.Reason = "";
                OList[0] = Rec3;

                fList = new MessageAckFailData[1];
                Rec = new MessageAckFailData();
                Rec.VSLInfo = vList;
                Rec.Orders = OList;
                fList[0] = Rec;

            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return fList;
        }

        private void WriteXmlFiles()
        {
            try
            {
                if (AckMessage == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | Acks Not Generated.";
                    return;
                }

                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");
                XmlSerializer mySerializer = new XmlSerializer(typeof(MessageAck));
                StreamWriter writer = new StreamWriter(File_Name);
                mySerializer.Serialize(writer, AckMessage, ns);
                writer.Close();
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }

    }
}
