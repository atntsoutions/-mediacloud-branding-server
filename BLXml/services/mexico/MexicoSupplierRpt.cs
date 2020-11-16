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
using BLXml.models.MexicoSupplier;

namespace BLXml 
{
   public class MexicoSupplierRpt : BaseReport
    {
        private DataTable DT_Supplier = new DataTable();
        private DataTable DT_Comp = new DataTable();
        public Boolean IsError = false;
        public string ErrorMessage = "";
        private MessageSupplier SupMessage = null;
        private string sql = "";
        public string PKID = "";
        public string File_Name = "";
        public string branch_code = "";
        DBConnection Con_Oracle = null;
        public void Generate()
        {
            try
            {
                ErrorMessage = "";
                IsError = false;
                ReadData();
                if (DT_Supplier.Rows.Count <= 0)
                {
                    IsError = true;
                    ErrorMessage = "Supplier Details not Found";
                    return;
                }
                if (DT_Comp.Rows.Count <= 0)
                {
                    IsError = true;
                    ErrorMessage = "Company city code not Found";
                    return;
                }
                IsError = false;
                GenerateXmlFiles();
                WriteXmlFiles();
                DT_Supplier.Rows.Clear();
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

            sql = " select cust_code,cust_name,b.add_country as cust_country  from customerm a";
            sql += "  left join (";
            sql += "  select add_parent_id, max(cntry.param_code) as add_country ";
            sql += "   from addressm a ";
            sql += "   left join param cntry on a.add_country_id=cntry.param_pkid";
            sql += "   where add_country_id is not null ";
            sql += "   and add_parent_id in ('" + PKID + "') ";
            sql += "   group by add_parent_id";
            sql += "  ) b on a.cust_pkid = b.add_parent_id";
            sql += "  where cust_pkid  in ('" + PKID + "') ";

            DT_Supplier = new DataTable();
            DT_Supplier = Con_Oracle.ExecuteQuery(sql);

            sql = "select comp_pol_code,comp_country_code,comp_tel from companym where rec_branch_code = '" + branch_code + "'";
            DT_Comp = new DataTable();
            DT_Comp = Con_Oracle.ExecuteQuery(sql);

            Con_Oracle.CloseConnection();

        }

        private void GenerateXmlFiles()
        {
            SupMessage = new MessageSupplier();
            SupMessage.ProcessID = XmlLib.PROCESSID;
            SupMessage.Suppliers = Generate_Suppliers();
        }

        private MessageSupplierSuppliersSupplier [] Generate_Suppliers()
        {
            MessageSupplierSuppliersSupplier Rec = null;
            MessageSupplierSuppliersSupplier[] mSupList = null;
            int ArrIndex = 0;
            try
            {
                mSupList = new MessageSupplierSuppliersSupplier[DT_Supplier.Rows.Count];
                foreach (DataRow Dr in DT_Supplier.Rows)
                {
                    Rec = new MessageSupplierSuppliersSupplier();
                    Rec.Customer = "";
                    Rec.SupplierName = Dr["cust_name"].ToString();
                    Rec.SupplierCode = Dr["cust_code"].ToString();
                    Rec.Country = Dr["cust_country"].ToString();
                    Rec.TCARGOMARountryCode = DT_Comp.Rows[0]["comp_country_code"].ToString();
                    Rec.TCARGOMARityCode = Lib.GetPortCode(DT_Comp.Rows[0]["comp_pol_code"].ToString());
                    Rec.TelNo = DT_Comp.Rows[0]["comp_tel"].ToString();
                    mSupList[ArrIndex++] = Rec;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mSupList;
        }
       
        private void WriteXmlFiles()
        {
            try
            {
                if (SupMessage == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | Suppliers Not Generated.";
                    return;
                }
                
                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");
                XmlSerializer mySerializer = new XmlSerializer(typeof(MessageSupplier));
                StreamWriter writer = new StreamWriter(File_Name);
                mySerializer.Serialize(writer, SupMessage, ns);
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
