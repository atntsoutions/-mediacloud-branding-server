using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text.RegularExpressions;
using DataBase.Connections;


namespace DataBase
{
    public class SmtpMail
    {
        public Boolean SendEmail(Dictionary<string, object> SearchData, out string ErrorMessage)
        {
            Boolean bRet = false;
            string Report_File_Name = "";
            string Report_File_DispName = "";
            //string Auto_Bcc_Email_ID_Branch = "";
            string sql = "";
            string ReadReceipt = "NO";
            string DeliveryReceipt = "NO";
            string TO_IDS = "";
            string CC_IDS = "";
            string BCC_IDS = "";
            string Subject = "";
            string Message = "";
            string User_Pkid = "";
            Boolean isCommonId = true;
            bool canftp = false;
            try
            {
                ErrorMessage = "";
                // physicalPath = AppDomain.CurrentDomain.BaseDirectory;
                //  SPATH = Filter["PATH"].ToString();
                // BRANCH_CODE = userInfo["REC_BRANCH_CODE"].ToString();
                //Boolean Is_Auto_Bcc = false;

                string[] IDS = null;

                if (SearchData.ContainsKey("to_ids"))
                    TO_IDS = SearchData["to_ids"].ToString().Replace(";", ",").ToLower();
                if (SearchData.ContainsKey("cc_ids"))
                    CC_IDS = SearchData["cc_ids"].ToString().Replace(";", ",").ToLower();
                if (SearchData.ContainsKey("bcc_ids"))
                    BCC_IDS = SearchData["bcc_ids"].ToString().Replace(";", ",").ToLower();
                if (SearchData.ContainsKey("subject"))
                    Subject = SearchData["subject"].ToString();
                if (SearchData.ContainsKey("message"))
                    Message = SearchData["message"].ToString().Replace("\n", "<br>");
                if (SearchData.ContainsKey("filedisplayname"))
                    Report_File_DispName = SearchData["filedisplayname"].ToString();
                if (SearchData.ContainsKey("filename"))
                    Report_File_Name = SearchData["filename"].ToString();
                if (SearchData.ContainsKey("user_pkid"))
                    User_Pkid = SearchData["user_pkid"].ToString();
                if (SearchData.ContainsKey("canftp"))
                    canftp = SearchData["canftp"].ToString() == "Y" ? true : false;

                if (SearchData.ContainsKey("iscommonid"))
                    isCommonId = (Boolean) SearchData["iscommonid"];


                if (canftp)
                {
                    Message += "<br>** EDI Files have been uploaded.";
                    Message += "<br>** This is a system generated email.";
                }

                if (SearchData.ContainsKey("read_receipt"))
                    ReadReceipt = SearchData["read_receipt"].ToString();
                if (SearchData.ContainsKey("delivery-receipt"))
                    DeliveryReceipt = SearchData["delivery_receipt"].ToString();

                string Email_Display_Name = "";
                string SMTP_SERVER = "";
                int SMTP_PORT = 587;
                string EmailID = "";
                string EmailPwd = "";
                Boolean SMTP_SSL = true;
                Boolean SMTP_Authentication = true;
                //Boolean SMTP_SPA = false;


                //DataTable Dt_Temp = new DataTable();
                //sql = " select param_name3 from mast_param  where param_type ='BRANCH SETTINGS' and param_name1 = 'AUTO-BCC-EMAIL-ID' and rec_branch_code = '" + BRANCH_CODE + "'";
                //Dt_Temp = DB.ExecuteQuery(sql);
                //if (Dt_Temp.Rows.Count > 0)
                //{
                //    Auto_Bcc_Email_ID_Branch = Dt_Temp.Rows[0]["param_name3"].ToString();
                //}


                DataTable Dt_User = new DataTable();
                if (isCommonId)
                {
                    
                    sql = "";
                    sql += " select mail_pkid, mail_name, mail_smtp_name, mail_smtp_port, mail_is_ssl_required, mail_is_auth_required, mail_is_spa_required, mail_id, mail_pwd from mailserver ";

                    DBConnection Con_Oracle = new DBConnection();
                    Dt_User = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    foreach (DataRow DrUser in Dt_User.Rows)
                    {
                        EmailID = DrUser["mail_id"].ToString().ToLower();
                        EmailPwd = DrUser["mail_pwd"].ToString();

                        SMTP_SERVER = DrUser["mail_smtp_name"].ToString();
                        SMTP_PORT = Lib.Conv2Integer(DrUser["mail_smtp_port"].ToString()) ;

                        SMTP_SSL = false;
                        if ( DrUser["mail_is_ssl_required"].ToString() == "Y")
                            SMTP_SSL = true;
                        SMTP_Authentication = false;
                        if (DrUser["mail_is_auth_required"].ToString() == "Y")
                            SMTP_Authentication = true;
                    }
                }
                else
                {
                    sql = "";
                    sql += " select user_email,user_email_pwd ";
                    sql += " from userm a ";
                    sql += " where user_pkid = '" + User_Pkid + "'";

                    DBConnection Con_Oracle = new DBConnection();
                    Dt_User = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow DrUser in Dt_User.Rows)
                    {
                        EmailID = DrUser["USER_EMAIL"].ToString().ToLower();
                        EmailPwd = DrUser["USER_EMAIL_PWD"].ToString();

                        //Email_Display_Name = DrUser["USER_EMAIL_DISPLAY_NAME"].ToString();

                        //if (DrUser["USR_EMAIL_AUTO_BCC"].ToString() == "Y")
                        //    Is_Auto_Bcc = true;
                        //else
                        //    Is_Auto_Bcc = false;

                        //SMTP_SERVER = DrUser["mail_smtp_name"].ToString();
                        //SMTP_PORT = Lib.Conv2Integer(DrUser["mail_smtp_port"].ToString());

                        //if (DrUser["MAIL_IS_SSL_REQUIRED"].ToString() == "Y")
                        //    SMTP_SSL = true;
                        //else
                        //    SMTP_SSL = false;

                        //if (DrUser["MAIL_IS_AUTH_REQUIRED"].ToString() == "Y")
                        //    SMTP_Authentication = true;
                        //else
                        //    SMTP_Authentication = false;

                        //if (DrUser["MAIL_IS_SPA_REQUIRED"].ToString() == "Y")
                        //    SMTP_SPA = true;
                        //else
                        //    SMTP_SPA = false;
                        break;
                    }
                }
                Dt_User.Rows.Clear();

                if (EmailID.Trim() == "" || EmailPwd.Trim() == "")
                {
                    bRet = false;
                    ErrorMessage = "From ID or Password cannot be blank.";
                    return bRet;
                }

                SmtpClient smtpClient = new SmtpClient(SMTP_SERVER, SMTP_PORT);
                smtpClient.EnableSsl = SMTP_SSL;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.UseDefaultCredentials = SMTP_Authentication;
                smtpClient.Credentials = new NetworkCredential(EmailID, EmailPwd);

                using (MailMessage message = new MailMessage())
                {
                    if (Email_Display_Name == "")
                        message.From = new MailAddress(EmailID);
                    else
                        message.From = new MailAddress(EmailID, Email_Display_Name);

                    message.Subject = Subject;
                    message.Body = Message;
                    message.IsBodyHtml = true;

                    /*
                    if (DeliveryReceipt == "YES")
                    {
                        message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure |
                                                                DeliveryNotificationOptions.OnSuccess |
                                                                DeliveryNotificationOptions.Delay;
                    }

                    if (ReadReceipt == "YES")
                        message.Headers.Add("Disposition-Notification-To", EmailID);
                        */

                    if (TO_IDS.Trim().Length > 0)
                    {
                        IDS = TO_IDS.Split(',');
                        foreach (string str in IDS)
                        {
                            if (str != "")
                                message.To.Add(str);
                        }
                    }
                    if (CC_IDS.Trim().Length > 0)
                    {
                        IDS = CC_IDS.Split(',');
                        foreach (string str in IDS)
                        {
                            if (str != "")
                                message.CC.Add(str);
                        }
                    }
                    if (BCC_IDS.Trim().Length > 0)
                    {
                        IDS = BCC_IDS.Split(',');
                        foreach (string str in IDS)
                        {
                            if (str != "")
                                message.Bcc.Add(str);
                        }
                    }

                    //if (Is_Auto_Bcc)
                    //    message.Bcc.Add(EmailID);

                    //if (Auto_Bcc_Email_ID_Branch != "")
                    //    message.Bcc.Add(Auto_Bcc_Email_ID_Branch);

                    if (Report_File_Name != "")
                    {
                        if (Report_File_Name.Contains(","))
                        {
                            string[] Fname = Report_File_Name.Split(',');
                            string[] Fname2 = null;
                            for (int i = 0; i < Fname.Length; i++)
                            {
                                Fname2 = Fname[i].Split('~');

                                Report_File_Name = Fname2[0].Trim();
                                Report_File_DispName = Fname2[1].Trim();
                                if (Report_File_Name != "")
                                {
                                    if (Report_File_DispName == "")
                                        Report_File_DispName = Report_File_Name;

                                    Attachment attachment = new Attachment(Report_File_Name, MediaTypeNames.Application.Octet);
                                    ContentDisposition disposition = attachment.ContentDisposition;
                                    disposition.FileName = Report_File_DispName;
                                    disposition.DispositionType = DispositionTypeNames.Attachment;
                                    message.Attachments.Add(attachment);
                                }
                            }
                        }
                        else
                        {
                            if (Report_File_DispName == "")
                                Report_File_DispName = Report_File_Name;

                            Attachment attachment = new Attachment(Report_File_Name, MediaTypeNames.Application.Octet);
                            ContentDisposition disposition = attachment.ContentDisposition;
                            disposition.FileName = Report_File_DispName;
                            disposition.DispositionType = DispositionTypeNames.Attachment;
                            message.Attachments.Add(attachment);
                        }
                    }

                    smtpClient.Send(message);
                    bRet = true;
                }

            }
            catch (Exception Ex)
            {
                bRet = false;
                ErrorMessage = Ex.Message.ToString();
            }

            
            /*
            DeleteFile(physicalPath + SPATH + "MAIL.XML");
            DeleteFile(physicalPath + SPATH + Report_File_Name);
            DeleteFolder(physicalPath + SPATH);
            */

            return bRet;
        }


    }
}
