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
using BLXml.models.MexicoOrder;
using System.Text;

namespace BLXml 
{
   public class MexicoOrdersRpt : BaseReport
    {
        private DataTable DT_ORDER = new DataTable();
        public Boolean IsError = false;
        public string ErrorMessage = "";
        private MessageCargoProcess CargoMessage = null;
        private string sql = "";
        public string PKID = "";
        public string InvokeType = "";
        public string File_Name = "";
        public int ProcessOrdCount = 0;
        public string File_Subject = "";
        public string Ftp_updtsql = "";
        DBConnection Con_Oracle = null;
        public string RowType = "";
        StringBuilder sb = new StringBuilder();
        public void Generate()
        {
            try
            {
                ErrorMessage = "";
                IsError = false;
                PKID = PKID.Replace(",", "','");
                ReadData();
                if (DT_ORDER.Rows.Count <= 0)
                {
                    IsError = true;
                    ErrorMessage = "Details not Found";
                    return;
                }

                IsError = false;
                if (RowType == "CHECK-LIST")
                {
                    GenerateCsvFiles();
                    WriteCsvFiles();
                }
                else
                {
                    GenerateXmlFiles();
                    WriteXmlFiles();
                }
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

            sql = "";
            sql  = " Select  ord_pkid,nvl(ord_status,'REPORTED') as ord_status";
            sql += "   ,ord_uid as poid";
            sql += "   ,ord_po as po";
            sql += "   ,ord_po as posupplier";
            sql += "   ,'' as oldpo";
            sql += "   ,ord_style as styleno";
            sql += "   ,ord_grwt as gw";
            sql += "   ,ord_pkg as packageno ";
            sql += "   ,ord_cbm as cbm";
            sql += "   ,ord_pcs as piecenumber";
            sql += "   ,ord_desc as Description";
            sql += "   ,ord_booking_date as BookingRequestDate";
            sql += "   ,ord_rnd_insp_date as RandomInspectionDate";
            sql += "   ,ord_po_rel_date as poreleasedate";
            sql += "   ,ord_cargo_ready_date as cargoreadydate";
            sql += "   ,ord_fcr_date as fcrdate";
            sql += "   ,ord_insp_date as inspectiondate";
            sql += "   ,ord_stuf_date as stuffingdate";
            sql += "   ,ord_whd_date as warehousedeparturedate";

            if (InvokeType == "CONTAINER")
            {
                sql += "   from packingm pkg";
                sql += "   inner join joborderm ord on pkg.pack_job_id = ord.ord_parent_id";
                sql += "   where pkg.pack_cntr_id='" + PKID + "'";
            }else
            {
                sql += " from joborderm ord";
                sql += " where  ord_pkid in ('" + PKID + "')";
            }

            sql += "   order by ord_po";

            DT_ORDER = new DataTable();
            DT_ORDER = Con_Oracle.ExecuteQuery(sql);
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

                if (dr["styleno"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Style No Cannot be Blank for PO " + dr["PO"].ToString();
                }

                if (Lib.Conv2Decimal(dr["gw"].ToString()) <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | GrWt Cannot be Blank for PO " + dr["PO"].ToString();
                }
                if (Lib.Conv2Decimal(dr["packageno"].ToString()) <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Package No Cannot be Blank for PO " + dr["PO"].ToString();
                }
                if (Lib.Conv2Decimal(dr["cbm"].ToString()) <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | CBM Cannot be Blank for PO " + dr["PO"].ToString();
                }
                if (Lib.Conv2Decimal(dr["piecenumber"].ToString()) <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Piece No Cannot be Blank for PO " + dr["PO"].ToString();
                }

            }
            File_Subject += str;

            Ftp_updtsql = "";
            //if (ord_pkid != "")
            //{
            //    if (ord_pkid.Contains(","))
            //        ord_pkid = ord_pkid.Replace(",", "','");
            //    Ftp_updtsql = "update joborderm set ord_ftp_status ='{STATUS}' where ord_pkid in ('" + ord_pkid + "')";
            //}

            if (InvokeType == "CONTAINER")
            {
                File_Subject = "";
                sql = "select cntr_no||'('||b.param_code||')' as cntr_no from containerm a ";
                sql += " left join param b on a.cntr_type_id= b.param_pkid";
                sql += " where cntr_pkid = '" + PKID + "'";
                Object sVal = Con_Oracle.ExecuteScalar(sql);
                File_Subject = sVal.ToString();
            }

            Con_Oracle.CloseConnection();
        }

        private void GenerateXmlFiles()
        {
            string yymmdd = DateTime.Now.ToString("yyyyMMdd");
            CargoMessage = new MessageCargoProcess();
           // CargoMessage.ProcessID = String.Concat(yymmdd, XmlLib.PROCESSID);
            CargoMessage.ProcessID = XmlLib.PROCESSID;
            CargoMessage.Orders= Generate_CargoOrders();
        }

        private MessageCargoProcessOrdersOrder[] Generate_CargoOrders()
        {
            MessageCargoProcessOrdersOrder Rec = null;
            MessageCargoProcessOrdersOrder[] mOrdList = null;
            int ArrIndex = 0;
            try
            {
                mOrdList = new MessageCargoProcessOrdersOrder[DT_ORDER.Rows.Count];
                foreach (DataRow Dr in DT_ORDER.Rows)
                {
                    Rec = new MessageCargoProcessOrdersOrder();
                    Rec.POID = Dr["poid"].ToString();
                    Rec.PO = Lib.GetTruncated(Dr["po"].ToString(),20);
                    Rec.POSupplier = Lib.GetTruncated(Dr["posupplier"].ToString(), 20);
                    Rec.OLDPO = Lib.GetTruncated(Dr["oldpo"].ToString(), 20);
                    Rec.StyleNo = Lib.GetTruncated(Dr["styleno"].ToString(), 20);
                    Rec.GW = Lib.NumericFormat(Dr["gw"].ToString(), 2);
                    Rec.PackageNo = Lib.NumericFormat(Dr["packageno"].ToString(),0);
                    Rec.CBM = Lib.NumericFormat(Dr["cbm"].ToString(),3);
                    Rec.PieceNumber = Lib.NumericFormat(Dr["piecenumber"].ToString(),0);
                    Rec.Description = Lib.GetTruncated(Dr["Description"].ToString(),100);
                    Rec.BookingRequestDate = Lib.DatetoStringDisplayformat(Dr["BookingRequestDate"]);
                    Rec.RandomInspectionDate = Lib.DatetoStringDisplayformat(Dr["RandomInspectionDate"]);
                    Rec.POReleaseDate = Lib.DatetoStringDisplayformat(Dr["poreleasedate"]);
                    Rec.CargoReadyDate = Lib.DatetoStringDisplayformat(Dr["cargoreadydate"]);
                    Rec.FCRDate = Lib.DatetoStringDisplayformat(Dr["fcrdate"]);
                    Rec.InspectionDate = Lib.DatetoStringDisplayformat(Dr["inspectiondate"]);
                    Rec.StuffingDate = Lib.DatetoStringDisplayformat(Dr["stuffingdate"]);
                    Rec.WarehouseDepartureDate = Lib.DatetoStringDisplayformat(Dr["warehousedeparturedate"]);
                    mOrdList[ArrIndex++] = Rec;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mOrdList;
        }
       
        private void WriteXmlFiles()
        {
            try
            {
                if (CargoMessage == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | Cargo Orders Not Generated.";
                    return;
                }
                
                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");
                XmlSerializer mySerializer = new XmlSerializer(typeof(MessageCargoProcess));
                StreamWriter writer = new StreamWriter(File_Name);
                mySerializer.Serialize(writer, CargoMessage, ns);
                writer.Close();
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }


        private void GenerateCsvFiles()
        {
         
            if (DT_ORDER.Rows.Count <= 0)
                return;

            sb = new StringBuilder();
            sb.Append("POID"); sb.Append(",");
            sb.Append("PO"); sb.Append(",");
            sb.Append("POSupplier"); sb.Append(",");
            sb.Append("OLDPO"); sb.Append(",");
            sb.Append("StyleNo"); sb.Append(",");
            sb.Append("GW"); sb.Append(",");
            sb.Append("PackageNo"); sb.Append(","); 
            sb.Append("CBM"); sb.Append(",");
            sb.Append("PieceNumber"); sb.Append(",");
            sb.Append("Description"); sb.Append(",");
            sb.Append("BookingRequestDate"); sb.Append(",");
            sb.Append("RandomInspectionDate"); sb.Append(",");
            sb.Append("POReleaseDate"); sb.Append(",");
            sb.Append("CargoReadyDate"); sb.Append(",");
            sb.Append("FCRDate"); sb.Append(",");
            sb.Append("InspectionDate"); sb.Append(",");
            sb.Append("StuffingDate"); sb.Append(",");
            sb.Append("WarehouseDepartureDate");
            foreach (DataRow Dr in DT_ORDER.Rows)
            {
                sb.AppendLine();
                sb.Append(Dr["poid"].ToString()); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["po"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["posupplier"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["oldpo"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["styleno"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.NumericFormat(Dr["gw"].ToString(), 2)); sb.Append(",");
                sb.Append(Lib.NumericFormat(Dr["packageno"].ToString(), 0)); sb.Append(",");
                sb.Append(Lib.NumericFormat(Dr["cbm"].ToString(), 3)); sb.Append(",");
                sb.Append(Lib.NumericFormat(Dr["piecenumber"].ToString(), 0)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["Description"].ToString(), 100)); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(Dr["BookingRequestDate"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(Dr["RandomInspectionDate"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(Dr["poreleasedate"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(Dr["cargoreadydate"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(Dr["fcrdate"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(Dr["inspectiondate"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(Dr["stuffingdate"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(Dr["warehousedeparturedate"]));
            }
        }

        private void WriteCsvFiles()
        {
            try
            {
                if (sb == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | Cargo Tracking Not Generated.";
                    return;
                }

                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                System.IO.File.AppendAllText(File_Name, sb.ToString());
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }

    }
}
