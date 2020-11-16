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
    public class EdiService : BL_Base
    {
        string company_code = "";
        string messagedoc_type = "";

        // SERVER

        public IDictionary<string, object> TransferFiles2EdiFolder(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string Error = "";

            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string User_Code = SearchData["user_code"].ToString();


            DataTable Dt_Folders = new DataTable();

            Con_Oracle = new DBConnection();
            try
            {

                string sql = "";
                sql += " select param_pkid,param_code, param_name,";
                sql += " max(case when param_key = 'FTP-LOCAL-FOLDER' then param_value else null end) as ftp_folder, ";
                sql += " max(case when param_key = 'LOCAL-FOLDER' then param_value else null end) as local_folder ";
                sql += " from( ";
                sql += "  select a.param_pkid, a.param_code, a.param_id1, a.param_name, b.param_key, b.param_value ";
                sql += "  from param a inner ";
                sql += "  join paramvalues b on a.param_pkid = b.parent_id ";
                sql += "  where a.rec_company_code = '{comp_code}' and  b.param_key  in ('FTP-LOCAL-FOLDER', 'LOCAL-FOLDER') ";
                sql += " ) a ";
                sql += " group by param_code, param_name, param_pkid ";
                sql = sql.Replace("{comp_code}", company_code);
                Dt_Folders = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Folders.Rows)
                {

                    try
                    {
                        Transfer(SearchData, Dr);
                    }
                    catch (Exception Ex)
                    {
                        if (Error != "")
                            Error += "\n";
                        Error += Dr["param_code"].ToString() + " : " + Ex.Message;
                    }
                }
            }

            catch (Exception Ex)
            {
                Con_Oracle.CloseConnection();
                throw Ex;
            }
            Con_Oracle.CloseConnection();

            RetData.Add("error", Error);
            return RetData;
        }

        private void Transfer(Dictionary<string, object> SearchData, DataRow Dr)
        {
            string[] edifiles = null;
            int iCount = 0;

            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string User_Code = SearchData["user_code"].ToString();

            string headerid = "";
            string msgdate = "";
            string filename = "";
            string extension = "";
            string localfolder = "";
            string destfilename = "";
            string destfullfilename = "";
            string iSlNo = "";

            DataTable Dt_files = new DataTable();
            DBRecord Rec;

            try
            {

                sql = "select param_filetype,param_value from paramvalues where parent_id = '" + Dr["param_pkid"].ToString() + "' and param_impexp = 'IMPORT' and param_edifile ='Y'";
                Dt_files = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow DrFiles in Dt_files.Rows)
                {

                    localfolder = System.IO.Path.Combine(Dr["ftp_folder"].ToString().Replace("/", "\\") + DrFiles["param_value"].ToString().Replace("/", "\\"));
                    edifiles = System.IO.Directory.GetFiles(localfolder);
                    iCount += edifiles.Length;

                    foreach (string strfile in edifiles)
                    {
                        headerid = Guid.NewGuid().ToString().ToUpper();
                        msgdate = System.IO.File.GetCreationTime(strfile).ToString("yyyy-MM-dd hh:mm:ss");
                        filename = System.IO.Path.GetFileName(strfile);
                        extension = System.IO.Path.GetExtension(strfile).Replace(".", "");
                        localfolder = Dr["local_folder"].ToString();

                        destfilename = filename;
                        destfullfilename = System.IO.Path.Combine(localfolder.Replace("/", "\\") + DrFiles["param_value"].ToString().Replace("/", "\\") + destfilename.Replace("/", "\\"));

                        sql = "select messagefilename from edi_header where messagefilename = '" + destfullfilename.ToString().ToUpper() + "' ";
                        sql += " and messagesender = '" + Dr["param_code"].ToString() + "' and rec_company_code = '" + company_code + "'";
                        if (Con_Oracle.IsRowExists(sql))
                            continue;

                        sql = " select nvl(max(slno),0) + 1 as slno from edi_header where rec_company_code  = '" + company_code + "'";
                        iSlNo = Con_Oracle.ExecuteScalar(sql).ToString(); ;

                        // Cannot copy if file is already esists
                        // in the case of image files it will be moved to documents folder when the files are processed.

                        FileCopyOrDelete("COPY", strfile, destfullfilename);

                        Rec = new DBRecord();
                        Rec.CreateRow("edi_header", "ADD", "headerid", headerid);
                        Rec.InsertNumeric("slno", iSlNo.ToString());
                        Rec.InsertString("messagesender", Dr["param_code"].ToString());
                        Rec.InsertString("messagedate", msgdate);
                        Rec.InsertString("messagedoctype", DrFiles["param_filetype"].ToString());
                        Rec.InsertString("messagefilename", destfullfilename);
                        Rec.InsertString("messageextension", extension);
                        Rec.InsertString("messageprocessed", "N");
                        Rec.InsertString("rec_company_code", company_code);

                        sql = Rec.UpdateRow();
                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                        try
                        {
                            FileCopyOrDelete("DELETE", strfile, "");
                        }
                        catch (Exception) { }
                    }
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                }
                throw Ex;
            }


        }


        private void FileCopyOrDelete(string flag, string sourcefile, string destfile)
        {
            if (flag == "COPY")
                System.IO.File.Copy(sourcefile, destfile);
            if (flag == "DELETE")
                System.IO.File.Delete(sourcefile);
        }

        public Dictionary<string, object> ImportEdiFiles(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            if (SearchData.ContainsKey("company_code"))
                company_code = SearchData["company_code"].ToString();
            if (SearchData.ContainsKey("messagedoc_type"))
                messagedoc_type = SearchData["messagedoc_type"].ToString();

            try
            {
                sql = "select headerid, messagesender,messagedoctype,nvl(messageformat,'DEFAULT') as messageformat,messagefilename, messageextension ";
                sql += " from edi_header ";
                sql += " where rec_company_code='" + company_code + "'";
                if (messagedoc_type != "ALL")
                    sql += " and messagedoctype='" + messagedoc_type + "'";
                sql += " and messageprocessed = 'N' ";
                sql += " order by slno";
                Con_Oracle = new DBConnection();
                DataTable Dt_temp = new DataTable();
                Dt_temp = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow dr in Dt_temp.Rows)
                {
                    if ( dr["messagedoctype"].ToString() == "PO")
                    {
                        if (dr["messageextension"].ToString() == "XML" && dr["messageformat"].ToString() == "DEFAULT")
                        {
                            DefaultPOImport POImp = new DefaultPOImport();
                            POImp.HeaderID = dr["headerid"].ToString();
                            POImp.FilePathName = dr["messagefilename"].ToString();
                            POImp.company_code = company_code;
                            POImp.Sender = dr["messagesender"].ToString();
                            POImp.doctype = "PO";
                            POImp.ImportData();
                        }

                    }

                    if (dr["messagedoctype"].ToString() == "PO TRACKING")
                    {
                        if (dr["messageextension"].ToString() == "XML" && dr["messageformat"].ToString() == "DEFAULT")
                        {
                            DefaultPOImport POImp = new DefaultPOImport();
                            POImp.HeaderID = dr["headerid"].ToString();
                            POImp.FilePathName = dr["messagefilename"].ToString();
                            POImp.company_code = company_code;
                            POImp.Sender = dr["messagesender"].ToString();
                            POImp.doctype = "PO TRACKING";
                            POImp.ImportData();
                        }

                    }
                    if (dr["messagedoctype"].ToString() == "HBL")
                    {
                        if (dr["messageextension"].ToString() == "XML" && dr["messageformat"].ToString() == "DEFAULT")
                        {
                            DefaultBLImport BLImp = new DefaultBLImport();
                            BLImp.HeaderID = dr["headerid"].ToString();
                            BLImp.FilePathName = dr["messagefilename"].ToString();
                            BLImp.company_code = company_code;
                            BLImp.Sender = dr["messagesender"].ToString();
                            // BLImp.doctype = "BL";
                            BLImp.ImportData();
                        }

                    }

                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("status", "COMPLETE");
            return RetData;
        }





    }
}
