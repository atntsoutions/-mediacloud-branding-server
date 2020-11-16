
using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using System.Xml.Serialization;
using System.IO;

using BLEdi.DefaultPO.Schema;

namespace BLEdi 
{
    public class DefaultPOImport : BL_Base
    {
        public string Sender = "";
        public string company_code = "";
        public string FilePathName = "";
        public string HeaderID = "";
        public string doctype = "";

        private string Sql = "";
        private string Ord_pkid = "";



        OrderMessage poXml = null;
        
        public Boolean ImportData()
        {
            Boolean bRet = false;
            if (FilePathName == "")
                return false;
            ReadXmlFiles();
            ProcessFile();
            Lib.InsertMappingData_EDI_ORDER(HeaderID);
            return bRet;
        }
        private void ReadXmlFiles()
        {
            XmlSerializer mySerializer = new XmlSerializer(typeof(OrderMessage));
            StreamReader reader = new StreamReader(FilePathName);
            poXml = (OrderMessage)mySerializer.Deserialize(reader);
            reader.Close();
        }

        private Boolean ProcessFile()
        {
            string sError = "";
            object[] items = poXml.Items;
            OrderMessageOrders ordRec;
            bool bRet = false;
            try
            {
                Con_Oracle = new DBConnection();
                OrderMessageMessageInfo msgInfo = (OrderMessageMessageInfo)items[0];
                if (msgInfo.MessageNumber.Length <= 0)
                {
                    sError = "Invalid Message Number";
                    sql = "update edi_header set  messageremarks='" + sError + "' where headerid='" + HeaderID + "'";
                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    Con_Oracle.CloseConnection();
                    return false;
                }

                Con_Oracle.BeginTransaction();

                sql = "update edi_header set messagenumber='" + msgInfo.MessageNumber + "', messageremarks = '', messageprocessed = 'Y'   where headerid='" + HeaderID + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                for (int i = 1; i < items.Count(); i++)
                {

                    ordRec = (OrderMessageOrders)items[i];

                    Ord_pkid = Guid.NewGuid().ToString().ToUpper();
                    sql = InsertOrder(ordRec, "ADD", Ord_pkid, HeaderID, msgInfo.MessagePolAgent);
                    Con_Oracle.ExecuteNonQuery(sql);
                }

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
                bRet = true;
            }
            catch (Exception Ex)
            {
                bRet = false;
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.BeginTransaction();
                    sql = "update edi_header set  messageremarks='" + TruncateData(Ex.Message, 100) + "' where headerid='" + HeaderID + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    Con_Oracle.CloseConnection();
                }
            }
            return bRet;
        }


        private string TruncateData(string sdata, int slen)
        {
            try
            {
                if (sdata != null)
                    if (sdata.Length > slen)
                        sdata = sdata.Substring(0, slen);

            }
            catch (Exception)
            {
            }
            return sdata;
        }
        private string InsertOrder(OrderMessageOrders ordRec, string mode, string pkid, string mHeaderid,string MsgpolAgent)
        {
            DBRecord mRec = new DBRecord();
            mRec.CreateRow("edi_order", mode, "ord_pkid", pkid);
            mRec.InsertString("ord_headerid", mHeaderid);
            mRec.InsertString("ord_sender", Sender);
            mRec.InsertString("ord_status", doctype);

            mRec.InsertString("ord_exp_name", TruncateData(ordRec.Shipper,60));
            mRec.InsertString("ord_imp_name", TruncateData(ordRec.consignee,60));
            mRec.InsertString("rec_category", ordRec.Mode);
            mRec.InsertString("ord_pol", TruncateData(ordRec.Pol,60));
            mRec.InsertString("ord_pod", TruncateData(ordRec.Pod,60));
            mRec.InsertString("ord_uneco", ordRec.Uneco);
            mRec.InsertString("ord_invno", ordRec.InvoiceNumber);
            mRec.InsertString("ord_po", ordRec.OrderNumber);
            mRec.InsertString("ord_style", ordRec.Style);
            mRec.InsertString("ord_color", ordRec.Color);
            mRec.InsertNumeric("ord_pkg", Lib.Conv2Integer(ordRec.Pkgs).ToString());
            mRec.InsertNumeric("ord_grwt", Lib.Conv2Decimal(ordRec.GrWt).ToString());
            mRec.InsertNumeric("ord_ntwt", Lib.Conv2Decimal(ordRec.NtWt).ToString());
            mRec.InsertNumeric("ord_pcs", Lib.Conv2Decimal(ordRec.Pcs).ToString());
            mRec.InsertNumeric("ord_cbm", Lib.Conv2Decimal(ordRec.Cbm).ToString());
            mRec.InsertString("ord_hs_code", ordRec.Hscode);
            mRec.InsertString("ord_desc", ordRec.Desc);
            mRec.InsertDate("ord_boarding1", ordRec.Boarding1);
            mRec.InsertDate("ord_boarding2", ordRec.Boarding2);
            mRec.InsertDate("ord_instock1", ordRec.Instock1);
            mRec.InsertDate("ord_instock2", ordRec.Instock2);
            
            mRec.InsertString("ord_pol_agent ", MsgpolAgent);
            mRec.InsertString("rec_company_code", company_code);

            mRec.InsertString("ord_updated", "N");
            mRec.InsertString("rec_deleted", "N");

            mRec.InsertDate("ord_booking_date", ordRec.BookingRequestDate);
            mRec.InsertDate("ord_rnd_insp_date", ordRec.RandomInspectionDate);
            mRec.InsertDate("ord_po_rel_date", ordRec.POReleaseDate);
            mRec.InsertDate("ord_cargo_ready_date", ordRec.CargoReadyDate);
            mRec.InsertDate("ord_fcr_date", ordRec.FCRDate);
            mRec.InsertDate("ord_insp_date", ordRec.InspectionDate);
            mRec.InsertDate("ord_stuf_date", ordRec.StuffingDate);
            mRec.InsertDate("ord_whd_date", ordRec.WarehouseDepartureDate);
            return mRec.UpdateRow();
        }

    }
}
