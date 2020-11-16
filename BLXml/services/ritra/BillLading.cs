using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.BL;

namespace BLXml
{
    public partial class BillLading : BLXml.models.XmlRoot
    {
        DBConnection Con_Oracle = null;
        string Container_Nature = "";
        // private string SORDERNOS = "";
        public override void Generate()
        {
            this.MODULE_ID = "BL";
            this.FileName = "BL";
            if (XmlLib.Agent_Name == "RITRA")
            {
                this.MODULE_ID = "BLCE";
                this.FileName = "BLCE";
            }

            BillLadingMessage MyList = new BillLadingMessage();

            MyList.BillLading = GetRecords();
            if (Total_Records > 0)
            {
                MyList.MessageInfo = GetMessageInfo();
                if (XmlLib.SaveInTempFolder)
                {
                    XmlLib.FolderId = Guid.NewGuid().ToString().ToUpper();
                    XmlLib.File_Type = "XML";
                    XmlLib.File_Category = "BLCE";
                    XmlLib.File_Processid = "BLCE" + this.MessageNumber;
                    XmlLib.File_Display_Name = FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlLib.File_Name = Lib.GetFileName(XmlLib.report_folder, XmlLib.FolderId, XmlLib.File_Display_Name);
                    XmlSerializer serializer =
                        new XmlSerializer(typeof(BillLadingMessage));
                    TextWriter writer = new StreamWriter(XmlLib.File_Name);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
                else
                {

                    XmlLib.CreateSentFolder();
                    FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlSerializer serializer =
                          new XmlSerializer(typeof(BillLadingMessage));
                    TextWriter writer = new StreamWriter(FileName);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
            }
        }
        public BillLadingMessageMessageInfo GetMessageInfo()
        {
            BillLadingMessageMessageInfo VMInfo = new BillLadingMessageMessageInfo();
            this.MessageNumber = XmlLib.GetNewMessageNumber();
            VMInfo.MessageNumber = this.MessageNumber;
            VMInfo.MessageSender = XmlLib.messageSenderField;
            VMInfo.MessageRecipient = XmlLib.messageRecipientField;
            VMInfo.CreatedDateTime = XmlLib.GetCreatedDate();
            VMInfo.CreatedDateTimeSpecified = true;

            return VMInfo;
        }
        public BillLadingMessageBillLading[] GetRecords()
        {
            string mID = "";
            int nCtr = 0;
            string sql = "";
            System.Collections.ArrayList aList = new ArrayList();
            BillLadingMessageBillLading Record;
            /*
            sql += " select   ";
            sql += " HBLS_HBL_ID,JOB_ID,HBL_NO,HBLS_BL_NO,HBL_YEAR,HBLS_BL_DATE,";
            sql += " OPR_SBILL_NO, OPR_SBILL_DATE, MBL_NO,";
            sql += " OPR_GR_NUMBER,OPR_GR_DATE, ";
            sql += " JOB_BOOKING_DATE,OPR_CARGO_RECEIVED_ON,OPR_CLEARED_DATE, ";
            sql += " VESSEL1_CODE, VESSEL2_CODE, VESSEL3_CODE,VESSEL4_CODE, ";
            sql += " VESSEL1_NAME,VESSEL2_NAME,VESSEL3_NAME,VESSEL4_NAME, ";
            sql += " Vessel1_voyage,Vessel2_voyage,Vessel3_voyage,Vessel4_voyage, ";
            sql += " TRANSIT1_CODE,TRANSIT2_CODE,TRANSIT3_CODE,TRANSIT4_CODE ";
            sql += " TRANSIT1_NAME,TRANSIT2_NAME,TRANSIT3_NAME,TRANSIT4_NAME, ";
            sql += " Vessel1_Etd,Vessel1_Eta, ";
            sql += " Vessel2_Etd,Vessel2_Eta, ";
            sql += " Vessel3_Etd,Vessel3_Eta, ";
            sql += " Vessel4_Etd,Vessel4_Eta, ";
            sql += " POFD_CODE,POFD_NAME,POD_CODE,POD_NAME,POL_CODE,POL_NAME, ";
            sql += " HBLS_POFD_ETA_CONF, HBLS_POFD_ETA,";
            sql += " Liner_Code,Liner_Name, ";
            sql += " Agent_Code,Agent_Name, ";
            sql += " Consignee_Code,Consignee_Name, ";
            sql += " PLACE_CODE,PLACE_NAME,NATURE_CODE, ";
            sql += " SHIPPER_CODE,SHIPPER_NAME,SHIPPER_ADD1,SHIPPER_ADD2,SHIPPER_ADD3,SHIPPER_ADD4,";
            sql += " hbl_consignee_line1 ,hbl_consignee_line2 ,hbl_consignee_line3 ,hbl_consignee_line4, ";
            sql += " AGENT_ADD1,AGENT_ADD2,AGENT_ADD3, AGENT_ADD4, ";
            sql += " HBL_NOTIFY_LINE1,HBL_NOTIFY_LINE2,HBL_NOTIFY_LINE3,HBL_NOTIFY_LINE4,HBL_NOTIFY_LINE5, ";
            sql += " COMMODITY_CODE,COMMODITY_NAME,PAY_CODE,HBLSTATUS_CODE,JOB_NOMINATION, ";
            sql += " MBL_FREIGHT_STATUS_ID,hblstatus_name";
            sql += " from  TABLE_XMLEDI  ";
            if (HBLNO.Length > 0)
            {
                sql += " where hbls_bl_no = '" + HBLNO + "'";
            }
            sql += " order by hbl_no";
            */


            sql = " select hbl.hbl_pkid,job.job_docno,hbl.hbl_no,hbl.hbl_bl_no,hbl.hbl_year,hbl.hbl_date ";
            sql += " ,rcpt.param_code as place_code ,rcpt.param_name as place_name";
            sql += " ,pol.param_code as pol_code,pol.param_name as pol_name";
            sql += " ,pod.param_code as pod_code,pod.param_name as pod_name";
            sql += " ,pofd.param_code as pofd_code,pofd.param_name as pofd_name";
            sql += " ,mbl.hbl_pofd_eta as pofd_eta";
            sql += " ,jexp.jexp_contract_nature as pay_code";
            sql += " ,hbl.hbl_nature as nature_code";
            sql += " ,mbl.hbl_terms as mbl_freight_status";
            sql += " ,hbl.hbl_terms as hbl_freight_status";
            sql += " ,nvl(job.job_nomination,cnge.cust_nomination) as job_nomination ";
            sql += " ,shpr.cust_code as hbl_exp_code,shpr.cust_name as hbl_exp_name";
            sql += " ,shpraddr.add_line1 as hbl_exp_add1,shpraddr.add_line2 as hbl_exp_add2,shpraddr.add_line3 as hbl_exp_add3,shpraddr.add_line4 as hbl_exp_add4";
            sql += " ,cnge.cust_code as hbl_imp_code,cnge.cust_name as hbl_imp_name";
            sql += " ,cngeaddr.add_line1 as hbl_imp_add1,cngeaddr.add_line2 as hbl_imp_add2,cngeaddr.add_line3 as hbl_imp_add3,cngeaddr.add_line4 as hbl_imp_add4";
            sql += " ,agnt.cust_code as mbl_agent_code,agnt.cust_name as mbl_agent_name";
            sql += " ,agntaddr.add_line1 as mbl_agent_add1,agntaddr.add_line2 as mbl_agent_add2,agntaddr.add_line3 as mbl_agent_add3,agntaddr.add_line4 as mbl_agent_add4";
            sql += " ,nfy.cust_code as  bl_notify_code,nfy.cust_name as  bl_notify_name";
            sql += " ,nfyaddr.add_line1 as bl_notify_add1,nfyaddr.add_line2 as bl_notify_add2,nfyaddr.add_line3 as bl_notify_add3,nfyaddr.add_line4 as bl_notify_add4";
            sql += " ,opr_cargo_received_on";
            sql += " ,com.param_code as commodity_code";
            sql += " ,hbl.hbl_pkg as tot_cartons,hbl.hbl_cbm as cbm,hbl.hbl_ntwt as ntwt,hbl.hbl_grwt as grwt";
            sql += " ,mbl.hbl_pkid as mbl_pkid,mbl.hbl_bl_no as mbl_no";
            sql += " ,lnr.param_code  as liner_code,lnr.param_name as liner_name ";
            sql += " ,mbl.rec_created_date,nvl(mbl.hbl_pol_etd,mbl.rec_created_date) as hbl_pol_etd  ";
            sql += "  from hblm mbl";
            sql += "  inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
            sql += "  inner join jobm job on hbl.hbl_pkid = job.jobs_hbl_id";
            sql += "  left join param rcpt on job.job_place_receipt_id = rcpt.param_pkid";
            sql += "  left join param pol on job.job_pol_id = pol.param_pkid";
            sql += "  left join param pod on job.job_pod_id = pod.param_pkid";
            sql += "  left join param pofd on job.job_pofd_id = pofd.param_pkid ";
            sql += "  left join jobexpm jexp on  job.job_pkid = jexp.jexp_job_id";
            sql += "  left join customerm shpr on hbl.hbl_exp_id = shpr.cust_pkid";
            sql += "  left join addressm shpraddr on hbl.hbl_exp_br_id = shpraddr.add_pkid";
            sql += "  left join customerm cnge on hbl.hbl_imp_id = cnge.cust_pkid";
            sql += "  left join addressm cngeaddr on hbl.hbl_imp_br_id = cngeaddr.add_pkid";
            sql += "  left join customerm agnt on mbl.hbl_agent_id = agnt.cust_pkid";
            sql += "  left join addressm agntaddr on mbl.hbl_agent_br_id = agntaddr.add_pkid";
            sql += "  left join bl on hbl.hbl_pkid = bl.bl_pkid";
            sql += "  left join customerm nfy on bl.bl_notify_id = nfy.cust_pkid";
            sql += "  left join addressm nfyaddr on bl.bl_notify_br_id = nfyaddr.add_pkid";
            sql += "  left join joboperationsm opr on job.job_pkid = opr.opr_job_id";
            sql += "  left join param com on job.job_commodity_id = com.param_pkid";
            sql += "  left join param lnr on mbl.hbl_carrier_id = lnr.param_pkid";
            sql += " where mbl.rec_company_code = '" + XmlLib.Company_Code + "' ";
            sql += " and mbl.rec_branch_code  = '" + XmlLib.Branch_Code + "' ";
            sql += " and mbl.rec_category  = 'SEA EXPORT' ";
            sql += " and mbl.hbl_agent_id = '" + XmlLib.Agent_Id + "' ";
            if (XmlLib.HBL_BL_NOS.Length > 0)
            {
                sql += " and hbl.hbl_bl_no in (" + XmlLib.HBL_BL_NOS + ")";
            }
            else if (XmlLib.MBL_IDS.Length > 0)
            {
                sql += " and mbl.hbl_pkid in (" + XmlLib.MBL_IDS + ")";
            }
            else
            {
               // sql += " and (sysdate - mbl.hbl_pol_etd) between 2 and  60 ";
                sql += " and (((sysdate - mbl.hbl_pol_etd) between 2 and  60) or (sysdate - mbl.hbl_prealert_date) <= 5 ) ";
            }
            sql += " order by hbl.hbl_no";

            try
            {
                DataTable Dt = new DataTable();
                Con_Oracle = new DBConnection();
                Dt = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt.Rows)
                {
                    if (mID != Dr["HBL_PKID"].ToString())
                    {
                        mID = Dr["HBL_PKID"].ToString();
                        nCtr++;
                        Record = new BillLadingMessageBillLading();
                        // Elements
                        Record.MemberCode = XmlLib.memberCode; // Sender Code 

                        //unlocode "JOB_RECD_PLACE,JOB_POL,JOB_POD,JOB_POFD"
                        Record.PlaceOfReceipt = Dr["PLACE_CODE"].ToString(); //Code for place of receipt
                        Record.PlaceOfReceipt_Name = Dr["PLACE_NAME"].ToString();
                        Record.PortOfLoading = XmlLib.GetPortCode(Dr["POL_CODE"].ToString()); //Code for port of loading
                        Record.PortOfLoading_Name = Dr["POL_NAME"].ToString();
                        Record.PortOfDischarge = XmlLib.GetPortCode(Dr["POD_CODE"].ToString()); // code for port of discharge
                        Record.PortOfDischarge_Name = Dr["POD_NAME"].ToString();
                        Record.PlaceOfDelivery = XmlLib.GetPortCode( Dr["POFD_CODE"].ToString()); // code for place of delivery

                        Record.PlaceOfDelivery_Name = Dr["POFD_NAME"].ToString();

                        if (Dr["POFD_ETA"].ToString() != "")
                        {
                            Record.ETADestination = (DateTime)Dr["POFD_ETA"]; // ETA At Destination
                            Record.ETADestinationSpecified = true; 
                        }
                        Record.PaymentTermSpecified = false;
                        if (Dr["PAY_CODE"].ToString() == "CIF")
                        {
                            Record.PaymentTerm = paymentTerm.CIF; // min1max1 Payment Term Object
                            Record.PaymentTermSpecified = true;
                        }
                        else if (Dr["PAY_CODE"].ToString() == "CF")
                        {
                            Record.PaymentTerm = paymentTerm.CFR; // min1max1 Payment Term Object
                            Record.PaymentTermSpecified = true;
                        }
                        else if (Dr["PAY_CODE"].ToString() == "CI")
                        {
                            Record.PaymentTerm = paymentTerm.CIF; // min1max1 Payment Term Object
                            Record.PaymentTermSpecified = true;
                        }
                        else if (Dr["PAY_CODE"].ToString() == "FOB")
                        {
                            Record.PaymentTerm = paymentTerm.FOB; // min1max1 Payment Term Object
                            Record.PaymentTermSpecified = true;
                        }

                        Record.MemberReferences = GetReferences(Dr);
                        Container_Nature = Dr["NATURE_CODE"].ToString();

                        Record.OceanBL = GetOceanBL(Dr);
                        Record.Remarks = GetRemarks(Dr);  // Remarks Object
                        Record.VoyageLegs = GetVoyageLegs(Dr); // Voyage Legs Object
                        Record.Parties = GetParties(Dr); // Parties Object
                        Record.Commodities = GetCommodities(Dr);

                        Record.Action = BillLadingMessageBillLadingAction.Replace;
                        Record.BLSeq = nCtr.ToString();
                        Record.BookingNo = Dr["HBL_YEAR"].ToString() + Dr["JOB_DOCNO"].ToString();
         
                        Record.HouseBLNo = Dr["HBL_BL_NO"].ToString(); // optional
                        //if (ActualShipper == "N")//Ajith 22/02/2018 Billing Party
                        //{
                        //    if (Dr["HBLS_BL_NO"].ToString().Contains(","))
                        //    {
                        //        string[] sdata = Dr["HBLS_BL_NO"].ToString().Split(',');
                        //        Record.HouseBLNo = sdata[0]; // optional
                        //    }
                        //}

                        if ( Dr["NATURE_CODE"].ToString() == "LCL/LCL" || Dr["NATURE_CODE"].ToString() == "CFS/CFS")
                        {
                            Record.Movement = movement.CFSCFS; // optional
                            Record.MovementSpecified = true;
                        }
                        else if (Dr["NATURE_CODE"].ToString() == "LCL/FCL" || Dr["NATURE_CODE"].ToString() == "CFS/CY")
                        {
                            Record.Movement = movement.CFSCY; // optional
                            Record.MovementSpecified = true;
                        }
                        else if (Dr["NATURE_CODE"].ToString() == "FCL/LCL" || Dr["NATURE_CODE"].ToString() == "CY/CFS")
                        {
                            Record.Movement = movement.CYCFS; // optional
                            Record.MovementSpecified = true;
                        }
                        else if (Dr["NATURE_CODE"].ToString() == "FCL/FCL" || Dr["NATURE_CODE"].ToString() == "CY/CY")
                        {
                            Record.Movement = movement.CYCY; // optional
                            Record.MovementSpecified = true;
                        }
                        else if (Dr["NATURE_CODE"].ToString() == "LCL" || Dr["NATURE_CODE"].ToString() == "LCLS")
                        {
                            Record.Movement = movement.CFSCFS;
                            Record.MovementSpecified = true;
                        }
                        

                        //Record.SecondMovement = ""; // optional
                        if (Dr["hbl_freight_status"].ToString().ToUpper().IndexOf("COL") > 0)
                        {
                            Record.Payment = "Collect";    // optional
                            Record.PayAt = "DESTINATION";      // optional
                        }
                        if (Dr["hbl_freight_status"].ToString().ToUpper().IndexOf("PRE") > 0)
                        {
                            Record.Payment = "Prepaid";    // optional
                            Record.PayAt = "ORIGIN";      // optional
                        }

                        if (Dr["JOB_NOMINATION"].ToString() == "NOMINATION")
                            Record.BLType = BillLadingMessageBillLadingBLType.Nomination;     // optional
                        if (Dr["JOB_NOMINATION"].ToString() == "FREEHAND")
                            Record.BLType = BillLadingMessageBillLadingBLType.Freehand;
                        if (Dr["JOB_NOMINATION"].ToString() == "MUTUAL")
                            Record.BLType = BillLadingMessageBillLadingBLType.JointEffort; 

                        Record.BLTypeSpecified = true;


                        //Record.IsExpress = false; // optional
                        if (Dr["HBL_DATE"].ToString().Trim().Length > 0)
                        {
                            Record.issueDate = (DateTime) Dr["HBL_DATE"]; // optional
                            Record.issuePlace = Dr["PLACE_NAME"].ToString();
                        }
                        Record.IsBooking = true; //optional                    
                        if (Dr["job_docno"].ToString().Trim().Length <= 0)
                            Record.IsBooking = false; //optional
                        Record.IsBookingSpecified = true;

                        //Record. = (DateTime) Dr[JOB.COL_JOB_BOOKING_DATE];  // optional - not in xml schema

                        aList.Add(Record);
                    }
                }
                
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
            Total_Records = nCtr;
            return (BillLadingMessageBillLading[])aList.ToArray(typeof(BillLadingMessageBillLading));
        }

        public BillLadingMessageBillLadingReference[] GetReferences(DataRow dRow)
        {
            System.Collections.ArrayList aList = new ArrayList();
            BillLadingMessageBillLadingReference Record = null;

            Record = new BillLadingMessageBillLadingReference();
            Record.refType = referenceType.JobNumber;
            Record.Value = dRow["JOB_DOCNO"].ToString();  
            aList.Add(Record);
            return (BillLadingMessageBillLadingReference[])aList.ToArray(typeof(BillLadingMessageBillLadingReference));
        }

        public BillLadingMessageBillLadingOceanBL GetOceanBL(DataRow dRow)
        {
            System.Collections.ArrayList aList = new ArrayList();
            BillLadingMessageBillLadingOceanBL Record = null;
            DataRow Dr = dRow;
            Record = new BillLadingMessageBillLadingOceanBL();
            BillLadingMessageBillLadingOceanBLCarrier Carrier = new BillLadingMessageBillLadingOceanBLCarrier();
            Carrier.Value = Dr["LINER_NAME"].ToString();
            Record.Carrier = Carrier;
            Record.OceanBLNo = Dr["MBL_NO"].ToString();
            Record.OceanBLPOR = Dr["PLACE_CODE"].ToString();
            Record.OceanBLPOR_Name = Dr["PLACE_NAME"].ToString();
            Record.OceanBLPOL = XmlLib.GetPortCode(Dr["POL_CODE"].ToString());
            Record.OceanBLPOL_Name = Dr["POL_NAME"].ToString();
            Record.OceanBLPOD = XmlLib.GetPortCode(Dr["POD_CODE"].ToString());
            Record.OceanBLPOD_Name = Dr["POD_NAME"].ToString();
            Record.OceanBLDEL = XmlLib.GetPortCode(Dr["POFD_CODE"].ToString());
            Record.OceanBLDEL_Name = Dr["POFD_NAME"].ToString();


            if (Container_Nature == "LCL/LCL" || Container_Nature == "CFS/CFS")
            {
                Record.OceanBLMovement = movement.CFSCFS; // optional
                Record.OceanBLMovementSpecified = true;
            }
            else if (Container_Nature == "LCL/FCL" || Container_Nature == "CFS/CY")
            {
                Record.OceanBLMovement = movement.CFSCY; // optional
                Record.OceanBLMovementSpecified = true;
            }
            else if (Container_Nature == "FCL/LCL" || Container_Nature == "CY/CFS")
            {
                Record.OceanBLMovement = movement.CYCFS; // optional
                Record.OceanBLMovementSpecified = true;
            }
            else if (Container_Nature == "FCL/FCL" || Container_Nature == "CY/CY")
            {
                Record.OceanBLMovement = movement.CYCY; // optional
                Record.OceanBLMovementSpecified = true;
            }
            else if (Container_Nature == "LCL" || Container_Nature == "LCLS")
            {
                Record.OceanBLMovement = movement.CFSCFS;
                Record.OceanBLMovementSpecified = true;
            }

            //if (Dr["MBL_FREIGHT_STATUS_ID"].ToString() == "4533FE67-9785-4779-9394-EB6400E68E8D")
            //    Record.OceanPayment = "Collect";
            //else
            //    Record.OceanPayment = "Prepaid";

            if (Dr["hbl_freight_status"].ToString().ToUpper().IndexOf("COL") > 0)
                Record.OceanPayment = "Collect";
            else
                Record.OceanPayment = "Prepaid";

            aList.Add(Record);
            return Record;
        }
        
        public String [] GetRemarks(DataRow dRow)
        {

            System.Collections.ArrayList aList = new ArrayList();
            aList.Add(" ");
            return (String[])aList.ToArray(typeof(String));
        }

        public BillLadingMessageBillLadingVoyageLeg[] GetVoyageLegs(DataRow Dr)
        {
            int nCtr = 0;
            DataTable Dt = null;
            string sql = "";
            ArrayList aList = new ArrayList();
            BillLadingMessageBillLadingVoyageLeg Record;

            /*
            sql += " select   ";
            sql += " HBLS_HBL_ID,JOB_ID,HBL_NO,HBLS_BL_NO,HBL_YEAR,";
            sql += " OPR_SBILL_NO, OPR_SBILL_DATE, ";
            sql += " OPR_GR_NUMBER,OPR_GR_DATE, ";
            sql += " JOB_BOOKING_DATE,OPR_CARGO_RECEIVED_ON,OPR_CLEARED_DATE, ";
            sql += " VESSEL1_CODE, VESSEL2_CODE, VESSEL3_CODE,VESSEL4_CODE, ";
            sql += " VESSEL1_NAME,VESSEL2_NAME,VESSEL3_NAME,VESSEL4_NAME, ";
            sql += " Vessel1_voyage,Vessel2_voyage,Vessel3_voyage,Vessel4_voyage, ";
            sql += " TRANSIT1_CODE,TRANSIT2_CODE,TRANSIT3_CODE,TRANSIT4_CODE ";
            sql += " TRANSIT1_NAME,TRANSIT2_NAME,TRANSIT3_NAME,TRANSIT4_NAME, ";
            sql += " Vessel1_Etd,Vessel1_Eta, ";
            sql += " Vessel2_Etd,Vessel2_Eta, ";
            sql += " Vessel3_Etd,Vessel3_Eta, ";
            sql += " Vessel4_Etd,Vessel4_Eta, ";
            sql += " POFD_CODE,POFD_NAME,POD_CODE,POD_NAME,POL_CODE,POL_NAME, ";
            sql += " HBLS_POFD_ETA_CONF, HBLS_POFD_ETA,";
            sql += " Liner_Code, ";
            sql += " Liner_Name, ";
            sql += " Agent_Code, ";
            sql += " Agent_Name ";
            sql += " PLACE_CODE,PLACE_NAME ";
            sql += " from VIEW_XMLEDI  "; ;
            sql += " where (AGENT_ID = '" + Lib.Agent_Id + "') and ";
            sql += " HBL_YEAR =" + dRow["HBL_YEAR"].ToString() + " and ";
            sql += " JOB_PKID =" + dRow["JOB_PKID"].ToString();

            Dt = new DataTable();
            StoredProcedure.CreateCommand(sql ) ;
            StoredProcedure.Run(Dt);
            */
            //foreach (DataRow Dr in Dt.Rows)
            /*
            if (Dr != null)
            {
                if (Dr["Vessel1_Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new BillLadingMessageBillLadingVoyageLeg();
                    //Liner.LinerName = Dr["LINER_NAME"].ToString();
                    Record.LinerCode = Dr["LINER_CODE"].ToString();
                    Record.LinerCode_Name = Dr["LINER_NAME"].ToString();
                     
                    Record.VesselCode = Dr["VESSEL1_CODE"].ToString();
                    Record.VoyageNumber = Dr["VESSEL1_VOYAGE"].ToString();
                    Record.LoadPort  = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                    Record.LoadPort_Name = Dr["POL_NAME"].ToString();


                    if (Dr["TRANSIT1_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = Dr["TRANSIT1_CODE"].ToString();
                        Record.DischargePort_Name = Dr["TRANSIT1_NAME"].ToString();

                    }
                    else if (Dr["TRANSIT2_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = "";
                        Record.DischargePort_Name = "";
                    }
                    else if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = "";
                        Record.DischargePort_Name = "";
                    }
                    else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = "";
                        Record.DischargePort_Name = "";
                    }
                    else
                    {
                        Record.DischargePort = Lib.GetPortCode( Dr["POD_CODE"].ToString());
                        Record.DischargePort_Name = Dr["POD_NAME"].ToString();
                    }


                    Record.VoyLegSeq = nCtr.ToString();
                    Record.PrintOnBL = true ;
                    aList.Add(Record);
                }
                if (Dr["Vessel2_Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new BillLadingMessageBillLadingVoyageLeg();
                    //Liner.LinerName = Dr["LINER_NAME"].ToString();
                    Record.LinerCode = Dr["LINER_CODE"].ToString();
                    Record.LinerCode_Name = Dr["LINER_NAME"].ToString();
                    Record.VesselCode = Dr["VESSEL2_CODE"].ToString();
                    Record.VoyageNumber  = Dr["VESSEL2_VOYAGE"].ToString();
                    Record.LoadPort = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                    Record.LoadPort_Name = Dr["POL_NAME"].ToString();


                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                    {
                        Record.LoadPort = Dr["TRANSIT1_CODE"].ToString();
                        Record.LoadPort_Name = Dr["TRANSIT1_NAME"].ToString();
                    }

                    if (Dr["TRANSIT2_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = Dr["TRANSIT2_CODE"].ToString();
                        Record.DischargePort_Name = Dr["TRANSIT2_NAME"].ToString();

                    }
                    else if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = "";
                        Record.DischargePort_Name = "";
                    }
                    else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = "";
                        Record.DischargePort_Name = "";
                    }
                    else
                    {
                        Record.DischargePort = Lib.GetPortCode( Dr["POD_CODE"].ToString());
                        Record.DischargePort_Name = Dr["POD_NAME"].ToString();
                    }




                    Record.VoyLegSeq = nCtr.ToString();
                    if (Dr["Vessel1_Code"].ToString().Trim().Length <= 0)
                        Record.PrintOnBL = true;
                    aList.Add(Record);
                }

                if (Dr["Vessel3_Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new BillLadingMessageBillLadingVoyageLeg();
                    //Liner.LinerName = Dr["LINER_NAME"].ToString();
                    Record.LinerCode = Dr["LINER_CODE"].ToString();
                    Record.LinerCode_Name = Dr["LINER_NAME"].ToString();
                    Record.VesselCode = Dr["VESSEL3_CODE"].ToString();
                    Record.VoyageNumber  = Dr["VESSEL3_VOYAGE"].ToString();
                    Record.LoadPort = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                    Record.LoadPort_Name = Dr["POL_NAME"].ToString();
                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                    {
                        Record.LoadPort = Dr["TRANSIT1_CODE"].ToString();
                        Record.LoadPort_Name  = Dr["TRANSIT1_NAME"].ToString();
                    }
                    if (Dr["TRANSIT2_CODE"].ToString().Trim().Length > 0)
                    {
                        Record.LoadPort = Dr["TRANSIT2_CODE"].ToString();
                        Record.LoadPort_Name = Dr["TRANSIT2_NAME"].ToString();
                    }

                    if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = Dr["TRANSIT3_CODE"].ToString();
                        Record.DischargePort_Name = Dr["TRANSIT3_NAME"].ToString();

                    }
                    else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                    {
                        Record.DischargePort = "";
                        Record.DischargePort_Name = "";
                    }
                    else
                    {
                        Record.DischargePort =Lib.GetPortCode(  Dr["POD_CODE"].ToString());
                        Record.DischargePort_Name = Dr["POD_NAME"].ToString();
                    }




                    Record.VoyLegSeq = nCtr.ToString();
                    if (Dr["Vessel1_Code"].ToString().Trim().Length <= 0)
                        if (Dr["Vessel2_Code"].ToString().Trim().Length <= 0)
                            Record.PrintOnBL = true;
                    aList.Add(Record);
                }
                if (Dr["Vessel4_Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new BillLadingMessageBillLadingVoyageLeg();
                    //Liner.LinerName = Dr["LINER_NAME"].ToString();
                    Record.LinerCode = Dr["LINER_CODE"].ToString();
                    Record.LinerCode_Name = Dr["LINER_NAME"].ToString();
                    Record.VesselCode = Dr["VESSEL4_CODE"].ToString();
                    Record.VoyageNumber  = Dr["VESSEL4_VOYAGE"].ToString();
                    Record.LoadPort =Lib.GetPortCode(  Dr["POL_CODE"].ToString());
                    Record.LoadPort_Name = Dr["POL_NAME"].ToString();
                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                    {
                        Record.LoadPort = Dr["TRANSIT1_CODE"].ToString();
                        Record.LoadPort_Name = Dr["TRANSIT1_CODE"].ToString();
                    }
                    if (Dr["TRANSIT2_CODE"].ToString().Trim().Length > 0)
                    {
                        Record.LoadPort = Dr["TRANSIT2_CODE"].ToString();
                        Record.LoadPort_Name = Dr["TRANSIT2_NAME"].ToString();
                    }
                    if (Dr["TRANSIT3_CODE"].ToString().Trim().Length > 0)
                    {
                        Record.LoadPort = Dr["TRANSIT3_CODE"].ToString();
                        Record.LoadPort_Name = Dr["TRANSIT3_NAME"].ToString();
                    }
                    Record.DischargePort = Lib.GetPortCode( Dr["POD_CODE"].ToString());
                    Record.DischargePort_Name = Dr["POD_NAME"].ToString();
                    
                    Record.VoyLegSeq = nCtr.ToString();
                    if (Dr["Vessel1_Code"].ToString().Trim().Length <= 0)
                        if (Dr["Vessel2_Code"].ToString().Trim().Length <= 0)
                            if (Dr["Vessel3_Code"].ToString().Trim().Length <= 0)
                                Record.PrintOnBL = true;
                    aList.Add(Record);
                }
            }*/


            sql = "  select vsl.param_code as vessel_code,vsl.param_name as vessel_name,trk.trk_voyage as vessel_voyage";
            sql += "  ,pol.param_code as pol_code,pol.param_name as pol_name";
            sql += "  ,pod.param_code as pod_code,pod.param_name as pod_name";
            sql += "  ,lnr.param_code  as liner_code,lnr.param_name as liner_name ";
            sql += "  from hblm mbl";
            sql += "  inner join trackingm trk on mbl.hbl_pkid = trk.trk_parent_id";
            sql += "  left join param vsl on trk_vsl_id= vsl.param_pkid ";
            sql += "  left join param pol on trk.trk_pol_id = pol.param_pkid";
            sql += "  left join param pod on trk.trk_pod_id = pod.param_pkid";
            sql += "  left join param lnr on mbl.hbl_carrier_id = lnr.param_pkid";
            sql += "  where  mbl.hbl_pkid = '"+ Dr["mbl_pkid"].ToString() + "'";
            sql += "  order by  trk_order";

            Dt = new DataTable();
            Con_Oracle = new DBConnection();
            Dt = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow dr in Dt.Rows)
            {
                if (dr["Vessel_Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new BillLadingMessageBillLadingVoyageLeg();
                    //Liner.LinerName = Dr["LINER_NAME"].ToString();
                    Record.LinerCode = dr["LINER_CODE"].ToString();
                    Record.LinerCode_Name = dr["LINER_NAME"].ToString();
                  //  Record.VesselCode = dr["VESSEL_CODE"].ToString();
                    Record.VesselCode = dr["VESSEL_NAME"].ToString();
                    Record.VoyageNumber = dr["VESSEL_VOYAGE"].ToString();
                    Record.LoadPort = XmlLib.GetPortCode(dr["POL_CODE"].ToString());
                    Record.LoadPort_Name = dr["POL_NAME"].ToString();
                    Record.DischargePort = XmlLib.GetPortCode(dr["POD_CODE"].ToString());
                    Record.DischargePort_Name = dr["POD_NAME"].ToString();
                    Record.VoyLegSeq = nCtr.ToString();
                    Record.PrintOnBL = true;
                    aList.Add(Record);
                }
            }
            return (BillLadingMessageBillLadingVoyageLeg[])aList.ToArray(typeof(BillLadingMessageBillLadingVoyageLeg));
        }
        public BillLadingMessageBillLadingParty[] GetParties(DataRow dRow)
        {
            int nCtr = 0;
            DataRow Dr = dRow;

            System.Collections.ArrayList aList = new ArrayList();
            BillLadingMessageBillLadingParty Record = null;
            //Shipper Record 
            Record = new BillLadingMessageBillLadingParty();
            nCtr++;
            //Attributes 
            Record.CompanyType = companyType.Shipper;
            Record.CompanyCode = Dr["hbl_exp_code"].ToString();
            //Record.AddressSeqNo = 0; // Branch Address
            Record.Name = Dr["hbl_exp_name"].ToString();
            Record.AddressLine1 = Dr["hbl_exp_add1"].ToString();
            Record.AddressLine2 = Dr["hbl_exp_add2"].ToString();
            Record.AddressLine3 = Dr["hbl_exp_add3"].ToString();
            Record.AddressLine4 = Dr["hbl_exp_add4"].ToString();

            //Record.City = "";
            //Record.ZipCode  = "";

            //if ( Dr[JOB.COL_EXPORTER_COUNTRY_EDI_CODE].ToString().Length  >0)
            //    Record.Country = GetCountry(Dr[JOB.COL_EXPORTER_COUNTRY_EDI_CODE].ToString());

            //Record.ContactPerson = "";
            //Record.Department = "";
            //Record.TelephoneNumber = Dr["SHIPPER_TEL"].ToString();
            //Record.FaxNumber = "";
            aList.Add(Record);

            //Consignee Record 
            Record = new BillLadingMessageBillLadingParty();
            nCtr++;
            //Attributes 
            Record.CompanyType = companyType.Consignee;
            Record.CompanyCode = Dr["hbl_imp_code"].ToString();
            //Record.AddressSeqNo = 0; // Branch Address

            Record.Name = Dr["hbl_imp_name"].ToString();
            Record.AddressLine1 = Dr["hbl_imp_add1"].ToString();
            Record.AddressLine2 = Dr["hbl_imp_add2"].ToString();
            Record.AddressLine3 = Dr["hbl_imp_add3"].ToString();
            Record.AddressLine4 = Dr["hbl_imp_add4"].ToString();
            //Record.City = "";
            //Record.ZipCode = "";
            //if (Dr[JOB.COL_CONS_COUNTRY_EDI_CODE].ToString().Length > 0)
            //    Record.Country = GetCountry(Dr[JOB.COL_CONS_COUNTRY_EDI_CODE].ToString());

            //Record.ContactPerson = "";
            //Record.Department = "";
            //Record.TelephoneNumber = Dr["CONSIGNEE_TEL"].ToString();
            //Record.FaxNumber = "";
            aList.Add(Record);


            //Agent Record 
            Record = new BillLadingMessageBillLadingParty();
            nCtr++;
            //Attributes 
            Record.CompanyType = companyType.Agent;
            Record.CompanyCode = Dr["mbl_agent_code"].ToString();
            //Record.AddressSeqNo = 0; // Branch Address

            Record.Name = Dr["mbl_agent_name"].ToString();
            Record.AddressLine1 = Dr["mbl_agent_add1"].ToString();
            Record.AddressLine2 = Dr["mbl_agent_add2"].ToString();
            Record.AddressLine3 = Dr["mbl_agent_add3"].ToString();
            Record.AddressLine4 = Dr["mbl_agent_add4"].ToString();

            //Record.City = "";
            //Record.ZipCode = "";
            //if (Dr[JOB.COL_CONS_COUNTRY_EDI_CODE].ToString().Length > 0)
            //    Record.Country = GetCountry(Dr[JOB.COL_CONS_COUNTRY_EDI_CODE].ToString());

            //Record.ContactPerson = "";
            //Record.Department = "";
            //Record.TelephoneNumber = Dr["AGENT_TEL"].ToString();
            //Record.FaxNumber = "";
            aList.Add(Record);


            //Notify Record 
            /*
            sql = "select noty_notify_Line1,noty_notify_Line2,noty_notify_Line3,noty_notify_Line4,";
            sql += " from notifym where rec_Deleted = 'N' and ";
            sql += " noty_hbl_id ='" + Dr["HBLS_HBL_ID"].ToString() + "' order by noty_slno" ;
            Dt_Temp = new DataTable(); 
            StoredProcedure.CreateCommand(sql);
            StoredProcedure.Run(Dt_Temp);
            */

            Record = new BillLadingMessageBillLadingParty();
            nCtr++;
            //Attributes 
            Record.CompanyType = companyType.Notify1;
            Record.CompanyCode = Dr["bl_notify_code"].ToString();
            //Record.AddressSeqNo = 0; // Branch Address


            Record.Name = Dr["bl_notify_name"].ToString();
            Record.AddressLine1 = Dr["bl_notify_add1"].ToString();
            Record.AddressLine2 = Dr["bl_notify_add2"].ToString();
            Record.AddressLine3 = Dr["bl_notify_add3"].ToString();
            Record.AddressLine4 = Dr["bl_notify_add4"].ToString();
            aList.Add(Record);

            return (BillLadingMessageBillLadingParty[])aList.ToArray(typeof(BillLadingMessageBillLadingParty));
        }
        public BillLadingMessageBillLadingCommodity[] GetCommodities(DataRow dRow)
        {
            try
            {
                System.Collections.ArrayList aList = new ArrayList();
                BillLadingMessageBillLadingCommodity Record = null;
                Record = new BillLadingMessageBillLadingCommodity();
                Record.ItemSeq = "1";
                Record.CommodityCode = dRow["COMMODITY_CODE"].ToString();
                Record.PurchaseOrder = "";
                Record.ItemNumber = "";
                if (dRow["OPR_CARGO_RECEIVED_ON"].ToString().Length > 0)
                {
                    Record.CargoReceivedDate = (DateTime)dRow["OPR_CARGO_RECEIVED_ON"];
                }
                Record.CargoDescriptions = getCargoDescriptions(dRow);
                Record.MarksNumbers = getMarksAndNumbers(dRow);

               // DataRow DrDet = getPackage(dRow);
                //Record.CommPackaging = getPacking(DrDet);
                //Record.CommMeasurement = getMeasurement(DrDet);
                //Record.CommWeight = getWeight(DrDet);

                Record.CommPackaging = getPacking(dRow);
                Record.CommMeasurement = getMeasurement(dRow);
                Record.CommWeight = getWeight(dRow);
                Record.CommContainer = getContainer(dRow);
                aList.Add(Record);
                return (BillLadingMessageBillLadingCommodity[])aList.ToArray(typeof(BillLadingMessageBillLadingCommodity));
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }

        private BillLadingMessageBillLadingCommodityMarksNumber[] getMarksAndNumbers(DataRow dRow)
        {
            DataTable Dt_Marks = new DataTable();
            string sql = "";
            //string[] str ;
            //string sData = "";
            string sPO = "";
            string sMarks = "";
            string[] BLMARKS;
            //string BLNO = "";
            BillLadingMessageBillLadingCommodityMarksNumber Record;
            System.Collections.ArrayList aList = new ArrayList();

            /*
            Boolean bAdded = false;
            if (XmlLib.Agent_Name == "RITRA")
            {
                if (bAdded == false)
                {
                    sql = " select distinct itm_orderno from itemm where rec_deleted = 'N' and itm_job_id ";
                    sql += " in (select jobs_job_id from job_summary where jobs_hbl_id = '" + dRow["HBLS_HBL_ID"].ToString() + "')";
                    StoredProcedure.CreateCommand(sql);
                    StoredProcedure.Run(Dt_Marks);
                    if (Dt_Marks.Rows.Count > 0)
                    {
                        foreach (DataRow Dr in Dt_Marks.Rows)
                        {
                            if (Dr["itm_orderno"].ToString().Trim().Trim().Length > 0)
                            {
                                Record = new BillLadingMessageBillLadingCommodityMarksNumber();
                                Record.Value = Dr["itm_orderno"].ToString();
                                aList.Add(Record);
                                bAdded = true;
                            }
                        }
                    }
                }
                
                if (bAdded == false)
                {
                    if (SORDERNOS.Trim().Length > 0)
                    {
                        string[] sRec = SORDERNOS.Split(',');
                        if (sRec.Length > 0)
                        {
                            for (int k = 0; k < sRec.Length; k++)
                            {
                                PO = sRec[k].ToString().Trim();
                                if (PO.Length > 0)
                                {
                                    Record = new BillLadingMessageBillLadingCommodityMarksNumber();
                                    Record.Value = PO;
                                    aList.Add(Record);
                                    bAdded = true;
                                }
                            }
                        }
                    }
                }

            }
            else 
            {
                sql = "select * from marks  where marks_hbl_id = '" + dRow["HBLS_HBL_ID"].ToString() + "'";
                StoredProcedure.CreateCommand(sql);
                StoredProcedure.Run(Dt_Marks);
                if (Dt_Marks.Rows.Count > 0)
                {
                    str = Dt_Marks.Rows[0]["marks_name"].ToString().Split('\n');
                    for (int i = 0; i < str.GetUpperBound(0); i++)
                    {
                        sData = str[i].Replace("\n", "").ToString();
                        if (sData.ToString().Length > 0)
                        {
                            Record = new BillLadingMessageBillLadingCommodityMarksNumber();
                            Record.Value = sData;
                            aList.Add(Record);
                        }
                    }
                }

               
            }*/
             
            sql = " select bl_itm_po from bl a  ";
            sql += " where  a.bl_pkid ='" + dRow["hbl_pkid"].ToString() + "'";
            Dt_Marks = new DataTable();
            Con_Oracle = new DBConnection();
            Dt_Marks = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            if (Dt_Marks.Rows.Count > 0)
            {
                sMarks = Dt_Marks.Rows[0]["bl_itm_po"].ToString();
                sMarks = sMarks.Replace("\r", "");
                sMarks = sMarks.Replace("\n", ",");
                if (sMarks.IndexOf(',') <= 0)
                    sMarks += ",";

                BLMARKS = sMarks.Split(',');
                for (int i = 0; i <= BLMARKS.GetUpperBound(0); i++)
                {
                    sPO = BLMARKS[i].ToString().Trim();
                    if (sPO.Length > 0)
                    {
                        Record = new BillLadingMessageBillLadingCommodityMarksNumber();
                        Record.Value = sPO;
                        aList.Add(Record);
                    }
                }
            }
            
            /*
             else
            {
                DateTime Dt_etd = (DateTime)dRow["hbl_pol_etd"];
                DateTime Dt_date = new DateTime(2018, 05, 05);
                //upto  bl_desc_ctr 18 is marks and number others bottom general statement
                sql = "select bl_marks from bldesc where bl_desc_ctr <=18 and bl_parent_id ='" + dRow["hbl_pkid"].ToString() + "'";
                if (Dt_etd > Dt_date)
                    sql += " and nvl(bl_is_mark,'N')='Y' ";
                sql += " order by bl_desc_ctr";


                Dt_Marks = new DataTable();
                Con_Oracle = new DBConnection();
                Dt_Marks = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow dr in Dt_Marks.Rows)
                {
                    if (dr["bl_marks"].ToString().Trim().Length > 0)
                    {
                        Record = new BillLadingMessageBillLadingCommodityMarksNumber();
                        Record.Value = dr["bl_marks"].ToString();
                        aList.Add(Record);
                    }
                }
            }
            */

            return (BillLadingMessageBillLadingCommodityMarksNumber[])aList.ToArray(typeof(BillLadingMessageBillLadingCommodityMarksNumber));
        }
        private BillLadingMessageBillLadingCommodityCargoDescriptionsCargoDescription[] getCargoDescriptions(DataRow dRow)
        {
            DataTable Dt = new DataTable();
            string sql = "";
            string sDesc = "";
            string[] BL_DESC;
            BillLadingMessageBillLadingCommodityCargoDescriptionsCargoDescription Record;
            System.Collections.ArrayList aList = new ArrayList();
            /*
            sql = " select hbls_desc, hbls_orderno from hbl_summary  where hbls_hbl_id = '" + dRow["HBLS_HBL_ID"].ToString() + "'";
            StoredProcedure.CreateCommand(sql);
            StoredProcedure.Run(Dt);

            sDesc = Dt.Rows[0]["HBLS_DESC"].ToString();
            if (Lib.Agent_Name == "RITRA")
                SORDERNOS = Dt.Rows[0]["HBLS_ORDERNO"].ToString();
            else
                SORDERNOS = "";

            sDesc = Dt.Rows[0]["HBLS_DESC"].ToString();
            sDesc = sDesc.Replace("\r", "");
            if (sDesc.Length > 0)
            {
                if (sDesc.IndexOf('\n') <= 0)
                    sDesc += "\n";
            }
            BL_DESC = sDesc.Split('\n');
            for (int i = 0; i <= BL_DESC.GetUpperBound(0); i++)
            {
                Record = new BillLadingMessageBillLadingCommodityCargoDescriptionsCargoDescription();
                //Record.Value = Dr["DESC"].ToString() + Dr["JITEM_DESC2"].ToString() + Dr["JITEM_DESC3"].ToString() + Dr["JITEM_DESC4"].ToString();
                Record.Value = BL_DESC[i].ToString();
                aList.Add(Record);
            }*/




            //sql = " select distinct ord_desc from joborderm a  ";
            //sql += " inner join jobm b on a.ord_parent_id = b.job_pkid";
            //sql += " where  b.jobs_hbl_id ='" + dRow["hbl_pkid"].ToString() + "'";

            sql = " select bl_itm_desc from bl a  ";
            sql += " where nvl(length(bl_itm_desc),0) > 0 and  a.bl_pkid ='" + dRow["hbl_pkid"].ToString() + "'";
            Dt = new DataTable();
            Con_Oracle = new DBConnection();
            Dt = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            if (Dt.Rows.Count > 0)
            {
                sDesc = Dt.Rows[0]["bl_itm_desc"].ToString();
                sDesc = sDesc.Replace("\r", "");
                if (sDesc.Length > 0)
                {
                    if (sDesc.IndexOf('\n') <= 0)
                        sDesc += "\n";
                }
                BL_DESC = sDesc.Split('\n');
                for (int i = 0; i <= BL_DESC.GetUpperBound(0); i++)
                {
                    Record = new BillLadingMessageBillLadingCommodityCargoDescriptionsCargoDescription();
                    Record.Value = BL_DESC[i].ToString();
                    aList.Add(Record);
                }
            }
            else
            {
                sql = "select bl_desc from bldesc where bl_parent_id ='" + dRow["hbl_pkid"].ToString() + "' and bl_desc_ctr > 2 order by bl_desc_ctr";
                Dt = new DataTable();
                Con_Oracle = new DBConnection();
                Dt = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow dr in Dt.Rows)
                {
                    if (dr["bl_desc"].ToString().Trim().Length > 0)
                    {
                        Record = new BillLadingMessageBillLadingCommodityCargoDescriptionsCargoDescription();
                        Record.Value = dr["bl_desc"].ToString();
                        aList.Add(Record);

                    }
                }
            }
            return (BillLadingMessageBillLadingCommodityCargoDescriptionsCargoDescription[])aList.ToArray(typeof(BillLadingMessageBillLadingCommodityCargoDescriptionsCargoDescription));
        }

        private DataRow getPackage(DataRow dRow)
        {
            DataTable Dt = new DataTable();
            //string sql = "";
            //sql  = " select sum(HBLS_TOTAL_CARTONS) as TOT_CARTONS, ";
            //sql += " sum(HBLS_TOTAL_CBM) as CBM,sum(HBLS_TOTAL_NET_WEIGHT) as NTWT,sum(HBLS_TOTAL_GROSS_WEIGHT) as GRWT ";
            //sql += " from HBL_SUMMARY where hbls_hbl_id = '" +dRow["HBLS_HBL_ID"].ToString() + "'" ;
            //StoredProcedure.CreateCommand(sql);
            //StoredProcedure.Run(Dt);
            return (Dt.Rows[0]);
        }

        private BillLadingMessageBillLadingCommodityCommPackaging[] getPacking(DataRow dRow)
        {
            System.Collections.ArrayList aList = new ArrayList();
            BillLadingMessageBillLadingCommodityCommPackaging Record;
            Record = new BillLadingMessageBillLadingCommodityCommPackaging();
            Record.PackageType = packageType.CT;
            Record.PackageTypeSpecified = true; 
            Record.Value = dRow["TOT_CARTONS"].ToString();
            aList.Add(Record);
            return (BillLadingMessageBillLadingCommodityCommPackaging[])aList.ToArray(typeof(BillLadingMessageBillLadingCommodityCommPackaging));
        }

        

        private BillLadingMessageBillLadingCommodityCommMeasurement[] getMeasurement(DataRow dRow)
        {
            try
            {
                BillLadingMessageBillLadingCommodityCommMeasurement Record;
                System.Collections.ArrayList aList = new ArrayList();
                Record = new BillLadingMessageBillLadingCommodityCommMeasurement();
                Record.Value = Lib.Convert2Decimal(  dRow["CBM"].ToString()) ;
                aList.Add(Record);
                return (BillLadingMessageBillLadingCommodityCommMeasurement[])aList.ToArray(typeof(BillLadingMessageBillLadingCommodityCommMeasurement));
            }
            catch (Exception Ex)
            {
                throw Ex;
            }

        }

        private BillLadingMessageBillLadingCommodityCommWeight[] getWeight(DataRow dRow)
        {
            BillLadingMessageBillLadingCommodityCommWeight Record;
            System.Collections.ArrayList aList = new ArrayList();
            Record = new BillLadingMessageBillLadingCommodityCommWeight();

            Record.Value = Lib.Convert2Decimal(dRow["GRWT"].ToString());
            aList.Add(Record);
            return (BillLadingMessageBillLadingCommodityCommWeight[])aList.ToArray(typeof(BillLadingMessageBillLadingCommodityCommWeight));
        }




        private BillLadingMessageBillLadingCommodityCommContainer[] getContainer(DataRow dRow)
        {
            DataTable Dt = new DataTable();
            string sql = "";
            string sType = "";
            BillLadingMessageBillLadingCommodityCommContainer Record;
            BillLadingMessageBillLadingCommodityCommContainerContainerPackaging mPack;
            BillLadingMessageBillLadingCommodityCommContainerContainerMeasurement mCbm;
            BillLadingMessageBillLadingCommodityCommContainerContainerWeight mWeight;
            System.Collections.ArrayList aList = new ArrayList();

            /*
            sql = " select substr(pack_container_no,1,13) as container_no,cntr_type as container_type,";
            sql += " pack_container_csealno as CSEAL_NO,  ";
            sql += " sum(pack_no_of_cartons) as TOT_CARTONS, ";
            sql += " sum(pack_qty) as QTY, ";
            sql += " sum(pack_cbm) as CBM, ";
            sql += " sum(pack_net_weight) as NTWT, ";
            sql += " sum(pack_gross_weight) as GRWT ";
            sql += " from view_blpackingm where hbls_hbl_id ='" + dRow["HBLS_HBL_ID"].ToString() + "' ";
            sql += " group by pack_container_no,cntr_type, pack_container_csealno ";

            sql = " select substr(pack_container_no,1,13) as container_no, ";
            sql += "  cntr_type as container_type, pack_container_csealno as CSEAL_NO,pack_container_asealno as ASEAL_NO,    ";
            sql += "  sum(pack_no_of_cartons) as TOT_CARTONS,  sum(pack_qty) as QTY,  sum(pack_cbm) as CBM,  ";
            sql += "  sum(pack_net_weight) as NTWT,  sum(pack_gross_weight) as GRWT  ";
            sql += "  from ";
            sql += " ( ";
            sql += " select cntr_no  as pack_container_no, param_code as cntr_type, ";
            sql += " cntr_csealno as pack_container_csealno,cntr_asealno as pack_container_asealno, ";
            sql += " pack_no_of_cartons, pack_qty,pack_cbm, ";
            sql += " pack_net_weight,pack_gross_weight ";
            sql += " from BLPackingm a ";
            sql += " inner join jobm 		  b on (a.pack_job_id = b.job_pkid) ";
            sql += " inner join job_summary c on (a.pack_job_id = c.jobs_job_id) ";
            sql += " inner join containerm  d on (a.pack_container_id = d.cntr_pkid) ";
            sql += " inner join param		  p on (d.cntr_type_id = p.param_pkid) ";
            sql += " inner join hbl_summary q on (c.jobs_hbl_id = q.hbls_hbl_id) ";
            sql += " where a.rec_deleted = 'N' and  hbls_hbl_id ='" + dRow["HBLS_HBL_ID"].ToString() + "' ";
            sql += " ) ";
            sql += " group by pack_container_no,cntr_type, pack_container_csealno, pack_container_asealno ";
            */

            sql = " select pack_container_no as container_no, "; //substr(replace(,' '),1,12)
            sql += "   cntr_type as container_type, pack_container_csealno as CSEAL_NO,pack_container_asealno as ASEAL_NO,    ";
            sql += "   sum(pack_pkg) as TOT_CARTONS,  sum(pack_pcs) as QTY,  sum(pack_cbm) as CBM,  ";
            sql += "   sum(pack_ntwt) as NTWT,  sum(pack_grwt) as GRWT  ";
            sql += "   from ";
            sql += "  ( ";
            sql += "  select cntr_no  as pack_container_no, param_code as cntr_type, ";
            sql += "  cntr_csealno as pack_container_csealno,cntr_asealno as pack_container_asealno, ";
            sql += "  PACK_PKG, PACK_PCS ,pack_cbm, ";
            sql += "  pack_ntwt,pack_grwt ";
            sql += "  from packingm a ";
            sql += "  inner join jobm 		  b on (a.pack_job_id = b.job_pkid) ";
            sql += "  inner join containerm  d on (a.pack_cntr_id = d.cntr_pkid) ";
            sql += "  inner join param		  p on (d.cntr_type_id = p.param_pkid) ";
            sql += "  where b.jobs_hbl_id ='" + dRow["hbl_pkid"].ToString() + "' ";
            sql += "  ";
            sql += "  ) ";
            sql += "  group by pack_container_no,cntr_type, pack_container_csealno, pack_container_asealno ";

            Con_Oracle = new DBConnection();
            Dt = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();

            string SEALNO = "";
            foreach (DataRow Dr in Dt.Rows)
            {
                Record = new BillLadingMessageBillLadingCommodityCommContainer();
                Record.ContainerNumber = Lib.GetCntrno(Dr["CONTAINER_NO"].ToString());
                sType = Dr["CONTAINER_TYPE"].ToString().Replace(" ", "");
                SEALNO = Dr["CSEAL_NO"].ToString();
                if (SEALNO.ToString().Trim() == "")
                    SEALNO = Dr["ASEAL_NO"].ToString();

                if (sType == "20FR")
                {
                    Record.ContainerType = containerType.Item20FR;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "20GH")
                {
                    Record.ContainerType = containerType.Item20GH;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }

                if (sType == "20HC")
                {
                    Record.ContainerType = containerType.Item20HC;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "20OT")
                {
                    Record.ContainerType = containerType.Item20OT;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "20RD")
                {
                    Record.ContainerType = containerType.Item20RD;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "20RF")
                {
                    Record.ContainerType = containerType.Item20RF;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "20ST" || sType == "20D" || sType == "20SD")
                {
                    Record.ContainerType = containerType.Item20ST;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "40FR")
                {
                    Record.ContainerType = containerType.Item40FR;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "40GH")
                {
                    Record.ContainerType = containerType.Item40GH;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "40HC")
                {
                    Record.ContainerType = containerType.Item40HC;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }

                if (sType == "40HD")
                {
                    Record.ContainerType = containerType.Item40HD;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }

                if (sType == "40HG")
                {
                    Record.ContainerType = containerType.Item40HG;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "40HR")
                {
                    Record.ContainerType = containerType.Item40HR;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }

                if (sType == "40OT")
                {
                    Record.ContainerType = containerType.Item40OT;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }

                if (sType == "40RD")
                {
                    Record.ContainerType = containerType.Item40RD;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "40RF")
                {
                    Record.ContainerType = containerType.Item40RF;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "40ST")
                {
                    Record.ContainerType = containerType.Item40ST;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }
                if (sType == "45ST")
                {
                    Record.ContainerType = containerType.Item45ST;
                    Record.SealNumber = SEALNO;
                    Record.ContainerTypeSpecified = true;
                }


                if (Container_Nature == "LCL/LCL" || Container_Nature == "CFS/CFS")
                {
                    Record.ContMovement = movement.CFSCFS;
                    Record.ContMovementSpecified = true;
                }
                else if (Container_Nature == "LCL/FCL" || Container_Nature == "CFS/CY")
                {
                    Record.ContMovement = movement.CFSCY;
                    Record.ContMovementSpecified = true;
                }
                else if (Container_Nature == "FCL/LCL" || Container_Nature == "CY/CFS")
                {
                    Record.ContMovement = movement.CYCFS;
                    Record.ContMovementSpecified = true;
                }
                else if (Container_Nature == "FCL/FCL" || Container_Nature == "CY/CY")
                {
                    Record.ContMovement = movement.CYCY;
                    Record.ContMovementSpecified = true;
                }
                else if (Container_Nature == "LCL" || Container_Nature == "LCLS")
                {
                    Record.ContMovement = movement.CFSCFS;
                    Record.ContMovementSpecified = true;
                }

                mPack = new BillLadingMessageBillLadingCommodityCommContainerContainerPackaging();
                mCbm = new BillLadingMessageBillLadingCommodityCommContainerContainerMeasurement();
                mWeight = new BillLadingMessageBillLadingCommodityCommContainerContainerWeight();

                mPack.Value = Dr["TOT_CARTONS"].ToString();

                mPack.PackageType = packageType.CT;
                mPack.PackageTypeSpecified = true;

                mCbm.Value = Lib.Convert2Decimal(Dr["CBM"].ToString());
                mWeight.Value = Lib.Convert2Decimal(Dr["GRWT"].ToString());

                Record.ContainerPackaging = mPack;
                Record.ContainerMeasurement = mCbm;
                Record.ContainerWeight = mWeight;

                aList.Add(Record);
            }
            return (BillLadingMessageBillLadingCommodityCommContainer[])aList.ToArray(typeof(BillLadingMessageBillLadingCommodityCommContainer));
        }


        private country GetCountry(string sCountry)
        {
            if (sCountry == "IN")
                return country.IN; 
            else
                return 0; 
        }
    }
}