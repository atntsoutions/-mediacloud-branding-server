
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

using BLEdi.DefaultBL.Schema;

namespace BLEdi
{
    public class DefaultBLImport : BL_Base
    {
        public string Sender = "";
        public string company_code = "";
        public string FilePathName = "";
        public string HeaderID = "";
        public string doctype = "";

        private string Sql = "";
        private string HBl_PKID = "";
        private string HBl_BL_NO = "";

        BLMessage blXml = null;

        public Boolean ImportData()
        {
            Boolean bRet = false;
            if (FilePathName == "")
                return false;
            ReadXmlFiles();
            ProcessFile();
            Lib.InsertMappingData_EDI_BL(HeaderID);
            return bRet;
        }
        private void ReadXmlFiles()
        {
            XmlSerializer mySerializer = new XmlSerializer(typeof(BLMessage));
            StreamReader reader = new StreamReader(FilePathName);
            blXml = (BLMessage)mySerializer.Deserialize(reader);
            reader.Close();
        }

        private Boolean ProcessFile()
        {
            string sError = "";
            object[] items = blXml.Items;
            BLMessageBLInfo blRec;
            bool bRet = false;
            DataTable Dt_temp;
            try
            {
                Con_Oracle = new DBConnection();
                BLMessageMessageInfo msgInfo = (BLMessageMessageInfo)items[0];
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

                sql = "update edi_header set messagenumber='" + msgInfo.MessageNumber + "', messageprocessed = 'Y'   where headerid='" + HeaderID + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                for (int i = 1; i < items.Count(); i++)
                {

                    blRec = (BLMessageBLInfo)items[i];
                    HBl_PKID = Guid.NewGuid().ToString().ToUpper();
                    HBl_BL_NO = "";
                    if (blRec.House != null && blRec.House.Length > 0)
                        HBl_BL_NO = blRec.House[0].HouseBLNo;

                    sql = "select hbl_pkid from edi_house where hbl_house_no='" + HBl_BL_NO + "' and hbl_sender='" + Sender + "'";
                    Dt_temp = new DataTable();
                    Dt_temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_temp.Rows.Count > 0)
                    {
                        HBl_PKID = Dt_temp.Rows[0]["hbl_pkid"].ToString();

                        sql = "delete from edi_house where hbl_pkid ='" + HBl_PKID + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                        sql = "delete from edi_house_vessel where vsl_hbl_id ='" + HBl_PKID + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                        sql = "delete from edi_house_container where cntr_hbl_id ='" + HBl_PKID + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                        sql = "delete from edi_house_order where ho_hbl_id ='" + HBl_PKID + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                        sql = "delete from edi_house_desc where hd_hbl_id ='" + HBl_PKID + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                    }

                    InsertHouse(blRec.Master, blRec.House, blRec.References, blRec.Remarks, blRec.Parties, blRec.VoyageGroup, msgInfo.MessagePolAgent);
                    foreach (BLMessageBLInfoVoyageGroupVoyageLeg mRec in blRec.VoyageGroup)
                    {
                        InsertVessel(mRec);
                    }
                    foreach (BLMessageBLInfoContainerGroupContainer mRec in blRec.ContainerGroup)
                    {
                        InsertContainer(mRec);
                    }
                    foreach (Line mRec in blRec.MarksAndNumbers)
                    {
                        InsertDescriptions(mRec, "1");
                    }
                    foreach (Line mRec in blRec.Description)
                    {
                        InsertDescriptions(mRec, "2");
                    }
                    foreach (Line mRec in blRec.WeightAndMeasurement)
                    {
                        InsertDescriptions(mRec, "3");
                    }
                    foreach (Line mRec in blRec.AdditionalDetails)
                    {
                        InsertDescriptions(mRec, "4");
                    }
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
        private void InsertHouse(BLMessageBLInfoMaster[] mRecords, BLMessageBLInfoHouse[] hRecords, BLMessageBLInfoReferencesReference[] rRecords,
            BLMessageBLInfoRemarksRemark[] remRecords, BLMessageBLInfoPartiesParty[] pRecords, BLMessageBLInfoVoyageGroupVoyageLeg[] vRecords, string MsgpolAgent)
        {

            DBRecord dbRec = new DBRecord();
            dbRec.CreateRow("edi_house", "ADD", "hbl_pkid", HBl_PKID);
            dbRec.InsertString("hbl_headerid", HeaderID);
            dbRec.InsertString("hbl_sender", Sender);
            dbRec.InsertString("hbl_pol_agent", MsgpolAgent);
            dbRec.InsertString("rec_company_code", company_code);
            dbRec.InsertString("hbl_updated", "N");
            dbRec.InsertString("rec_deleted", "N");

            foreach (BLMessageBLInfoMaster mRec in mRecords)
            {
                dbRec.InsertString("hbl_carrier_name", mRec.Carrier);
                dbRec.InsertString("hbl_agent_name", mRec.MasterAgent);
                dbRec.InsertString("hbl_master_no", mRec.MasterBLNO);
                dbRec.InsertDate("hbl_master_date", mRec.MasterBLDate);
                dbRec.InsertString("hbl_direct_bl", mRec.DirectBL);
                dbRec.InsertString("hbl_etd_confirm", mRec.EtdConfirm);
                break;
            }
            foreach (BLMessageBLInfoHouse hRec in hRecords)
            {
                dbRec.InsertString("hbl_house_no", hRec.HouseBLNo);
                dbRec.InsertDate("hbl_house_date", hRec.HouseBLDate);
                dbRec.InsertString("hbl_ams_no", hRec.Ams);
                dbRec.InsertString("hbl_isf_no", hRec.Isf);
                if (hRec.Por != null)
                {
                    dbRec.InsertString("hbl_por_code", hRec.Por.Length > 0 ? hRec.Por[0].Code : "");
                    dbRec.InsertString("hbl_por_name", hRec.Por.Length > 0 ? hRec.Por[0].Value : "");
                }
                if (hRec.Pol != null)
                {
                    dbRec.InsertString("hbl_pol_code", hRec.Pol.Length > 0 ? hRec.Pol[0].Code : "");
                    dbRec.InsertString("hbl_pol_name", hRec.Pol.Length > 0 ? hRec.Pol[0].Value : "");
                }
                if (hRec.Pod != null)
                {
                    dbRec.InsertString("hbl_pod_code", hRec.Pod.Length > 0 ? hRec.Pod[0].Code : "");
                    dbRec.InsertString("hbl_pod_name", hRec.Pod.Length > 0 ? hRec.Pod[0].Value : "");
                }
                if (hRec.Pofd != null)
                {
                    dbRec.InsertString("hbl_pofd_code", hRec.Pofd.Length > 0 ? hRec.Pofd[0].Code : "");
                    dbRec.InsertString("hbl_pofd_name", hRec.Pofd.Length > 0 ? hRec.Pofd[0].Value : "");
                }
                if (hRec.PlaceDelivery != null)
                {
                    dbRec.InsertString("hbl_delivery_code", hRec.PlaceDelivery.Length > 0 ? hRec.PlaceDelivery[0].Code : "");
                    dbRec.InsertString("hbl_delivery_name", hRec.PlaceDelivery.Length > 0 ? hRec.PlaceDelivery[0].Value : "");
                }
                dbRec.InsertDate("hbl_etd", hRec.Etd);
                dbRec.InsertDate("hbl_eta", hRec.Eta);
                dbRec.InsertDate("hbl_delivery_date", hRec.DeliveryDate);
                dbRec.InsertString("hbl_mode", hRec.Mode);
                dbRec.InsertString("hbl_movement", hRec.Movement);
                dbRec.InsertString("hbl_service_type", hRec.ServiceType);
                dbRec.InsertString("hbl_payment_type", hRec.PaymentTerm);
                dbRec.InsertString("hbl_freight", hRec.FreightTerm);
                if (hRec.Pkgs != null)
                {
                    dbRec.InsertString("hbl_pkg_unit", hRec.Pkgs.Length > 0 ? hRec.Pkgs[0].Unit : "");
                    dbRec.InsertNumeric("hbl_pkg", hRec.Pkgs.Length > 0 ? Lib.NumericFormat(hRec.Pkgs[0].Value, 0) : "0");
                }
                dbRec.InsertNumeric("hbl_grwt", Lib.NumericFormat(hRec.GrWt, 3));
                dbRec.InsertNumeric("hbl_ntwt", Lib.NumericFormat(hRec.NtWt, 3));
                dbRec.InsertNumeric("hbl_pcs", Lib.NumericFormat(hRec.Pcs, 3));
                dbRec.InsertNumeric("hbl_cbm", Lib.NumericFormat(hRec.Cbm, 3));
                break;
            }
            dbRec.InsertString("hbl_book_no", GetReferenceValue(rRecords, "BOOKINGNUMBER"));
            dbRec.InsertString("hbl_contract_no", GetReferenceValue(rRecords, "CONTRACTNUMBER"));
            dbRec.InsertString("hbl_ref_no", GetReferenceValue(rRecords, "REFERENCENUMBER"));
            dbRec.InsertString("hbl_branch", GetReferenceValue(rRecords, "BRANCH"));
            dbRec.InsertString("hbl_issu_place", GetReferenceValue(rRecords, "BLISSUEDPLACE"));
            dbRec.InsertDate("hbl_issu_date", GetReferenceValue(rRecords, "BLISSUEDDATE"));
            dbRec.InsertString("hbl_is_ddu", GetReferenceValue(rRecords, "DDU"));
            dbRec.InsertString("hbl_is_ddp", GetReferenceValue(rRecords, "DDP"));
            dbRec.InsertString("hbl_is_exwork", GetReferenceValue(rRecords, "EXWORK"));
            dbRec.InsertString("hbl_remark", remRecords.Length > 0 ? remRecords[0].Value : "");

            foreach (BLMessageBLInfoPartiesParty rec in pRecords)
            {
                if (rec.Type == "SHIPPER")
                {
                    dbRec.InsertString("hbl_shipper_name", rec.Name);
                    dbRec.InsertString("hbl_shipper_add1", rec.Address1);
                    dbRec.InsertString("hbl_shipper_add2", rec.Address2);
                    dbRec.InsertString("hbl_shipper_add3", rec.Address3);
                    dbRec.InsertString("hbl_shipper_add4", rec.Address4);
                    dbRec.InsertString("hbl_shipper_add5", rec.Address5);
                    dbRec.InsertString("hbl_shipper_contact", rec.Contact);
                    dbRec.InsertString("hbl_shipper_tel", rec.Tel);
                    dbRec.InsertString("hbl_shipper_email", rec.Email);
                }
                else if (rec.Type == "CONSIGNEE")
                {
                    dbRec.InsertString("hbl_consignee_name", rec.Name);
                    dbRec.InsertString("hbl_consignee_add1", rec.Address1);
                    dbRec.InsertString("hbl_consignee_add2", rec.Address2);
                    dbRec.InsertString("hbl_consignee_add3", rec.Address3);
                    dbRec.InsertString("hbl_consignee_add4", rec.Address4);
                    dbRec.InsertString("hbl_consignee_add5", rec.Address5);
                    dbRec.InsertString("hbl_consignee_contact", rec.Contact);
                    dbRec.InsertString("hbl_consignee_tel", rec.Tel);
                    dbRec.InsertString("hbl_consignee_email", rec.Email);
                }
                else if (rec.Type == "NOTIFY")
                {
                    dbRec.InsertString("hbl_notify_name", rec.Name);
                    dbRec.InsertString("hbl_notify_add1", rec.Address1);
                    dbRec.InsertString("hbl_notify_add2", rec.Address2);
                    dbRec.InsertString("hbl_notify_add3", rec.Address3);
                    dbRec.InsertString("hbl_notify_add4", rec.Address4);
                    dbRec.InsertString("hbl_notify_add5", rec.Address5);
                    dbRec.InsertString("hbl_notify_contact", rec.Contact);
                    dbRec.InsertString("hbl_notify_tel", rec.Tel);
                    dbRec.InsertString("hbl_notify_email", rec.Email);
                }
            }

            if (vRecords != null && vRecords.Length > 0)
            {
                dbRec.InsertString("hbl_vessel", vRecords[0].Vessel);
                dbRec.InsertString("hbl_voyage", vRecords[0].Voyage);
                dbRec.InsertString("hbl_mother_vessel", vRecords[vRecords.Length - 1].Vessel);
                dbRec.InsertString("hbl_mother_voyage", vRecords[vRecords.Length - 1].Voyage);
            }

            sql = dbRec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);
        }

        private string GetReferenceValue(BLMessageBLInfoReferencesReference[] rRecords, string _type)
        {
            string sValue = "";
            foreach (BLMessageBLInfoReferencesReference rRec in rRecords)
            {
                if (rRec.Type == _type)
                {
                    sValue = rRec.Value;
                    break;
                }
            }
            return sValue;
        }

        private void InsertVessel(BLMessageBLInfoVoyageGroupVoyageLeg Rec)
        {
    
            string vsl_pkid = Guid.NewGuid().ToString().ToUpper();
            DBRecord dbRec = new DBRecord();
            dbRec.CreateRow("edi_house_vessel", "ADD", "vsl_pkid", vsl_pkid);
            dbRec.InsertString("vsl_hbl_id", HBl_PKID);
            dbRec.InsertNumeric("vsl_seq", Lib.Conv2Integer(Rec.Seq).ToString());
            dbRec.InsertString("vsl_name", Rec.Vessel);
            dbRec.InsertString("vsl_voyage", Rec.Voyage);
            if (Rec.Pol != null)
            {
                dbRec.InsertString("vsl_pol_code", Rec.Pol.Length > 0 ? Rec.Pol[0].Code : "");
                dbRec.InsertString("vsl_pol_name", Rec.Pol.Length > 0 ? Rec.Pol[0].Value : "");
            }
            if (Rec.Pod != null)
            {
                dbRec.InsertString("vsl_pod_code", Rec.Pod.Length > 0 ? Rec.Pod[0].Code : "");
                dbRec.InsertString("vsl_pod_name", Rec.Pod.Length > 0 ? Rec.Pod[0].Value : "");
            }
            dbRec.InsertDate("vsl_etd", Rec.Etd);
            dbRec.InsertString("vsl_etd_confirm", Rec.EtdConfirm);
            dbRec.InsertDate("vsl_eta", Rec.Eta);
            dbRec.InsertString("vsl_eta_confirm", Rec.EtaConfirm);
            sql = dbRec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);
        }

        private void InsertContainer(BLMessageBLInfoContainerGroupContainer Rec)
        {
            string cntr_pkid = Guid.NewGuid().ToString().ToUpper();
            DBRecord dbRec = new DBRecord();
            dbRec.CreateRow("edi_house_container", "ADD", "cntr_pkid", cntr_pkid);
            dbRec.InsertString("cntr_hbl_id", HBl_PKID);
            dbRec.InsertString("cntr_no", Rec.ContainerNumber);
            dbRec.InsertString("cntr_size", Rec.Size);
            dbRec.InsertString("cntr_type", Rec.Type);
            dbRec.InsertString("cntr_aseal", Rec.ASeal);
            dbRec.InsertString("cntr_cseal", Rec.CSeal);
            if (Rec.Pkgs != null)
            {
                dbRec.InsertString("cntr_pkgs_unit", Rec.Pkgs.Length > 0 ? Rec.Pkgs[0].Unit : "");
                dbRec.InsertNumeric("cntr_pkgs", Rec.Pkgs.Length > 0 ? Lib.NumericFormat(Rec.Pkgs[0].Value, 0) : "0");
            }
            dbRec.InsertNumeric("cntr_grwt", Lib.NumericFormat(Rec.GrWt, 3));
            dbRec.InsertNumeric("cntr_ntwt", Lib.NumericFormat(Rec.NtWt, 3));
            dbRec.InsertNumeric("cntr_pcs", Lib.NumericFormat(Rec.Pcs, 3));
            dbRec.InsertNumeric("cntr_cbm", Lib.NumericFormat(Rec.Cbm, 3));

            sql = dbRec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);

            foreach (BLMessageBLInfoContainerGroupContainerOrdersOrder rec in Rec.Orders)
            {
                InsertOrders(rec, cntr_pkid);
            }
        }

        private void InsertOrders(BLMessageBLInfoContainerGroupContainerOrdersOrder Rec, string CNTR_PKID)
        {
            string ho_pkid = Guid.NewGuid().ToString().ToUpper();
            DBRecord dbRec = new DBRecord();
            dbRec.CreateRow("edi_house_order", "ADD", "ho_pkid", ho_pkid);
            dbRec.InsertString("ho_hbl_id", HBl_PKID);
            dbRec.InsertString("ho_cntr_id", CNTR_PKID);
            dbRec.InsertString("ho_ordno", Rec.OrderNumber);
            dbRec.InsertString("ho_style", Rec.Style);
            dbRec.InsertString("ho_color", Rec.Color);
            dbRec.InsertString("ho_invno", Rec.InvoiceNumber);
            if (Rec.Pkgs != null)
            {
                dbRec.InsertString("ho_pkgs_unit", Rec.Pkgs.Length > 0 ? Rec.Pkgs[0].Unit : "");
                dbRec.InsertNumeric("ho_pkgs", Rec.Pkgs.Length > 0 ? Lib.NumericFormat(Rec.Pkgs[0].Value, 0) : "0");
            }
            dbRec.InsertNumeric("ho_grwt", Lib.NumericFormat(Rec.GrWt, 3));
            dbRec.InsertNumeric("ho_ntwt", Lib.NumericFormat(Rec.NtWt, 3));
            dbRec.InsertNumeric("ho_pcs", Lib.NumericFormat(Rec.Pcs, 3));
            dbRec.InsertNumeric("ho_cbm", Lib.NumericFormat(Rec.Cbm, 3));

            sql = dbRec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);
        }

        private void InsertDescriptions(Line Rec, string _type)
        {
            string hd_pkid = Guid.NewGuid().ToString().ToUpper();
            DBRecord dbRec = new DBRecord();
            dbRec.CreateRow("edi_house_desc", "ADD", "hd_pkid", hd_pkid);
            dbRec.InsertString("hd_hbl_id", HBl_PKID);
            dbRec.InsertString("hd_type", _type);
            dbRec.InsertString("hd_desc", Rec.Value);
            dbRec.InsertNumeric("hd_seq", Lib.Conv2Integer(Rec.Seq).ToString());
            sql = dbRec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);
        }
    }
}
