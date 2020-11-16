using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

using System.Drawing;
using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLAccounts

{
    public class AcctmService : BL_Base
    {
        ExcelFile WB;
        ExcelWorksheet WS = null;
        int iRow = 0;
        int iCol = 0;
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string report_folder = "";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Acctm> mList = new List<Acctm>();
            Acctm mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];

            report_folder = SearchData["report_folder"].ToString();
            string comp_code = SearchData["comp_code"].ToString();

            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  a.rec_company_code = '" + comp_code + "'";
                if (searchstring != "")
                {
                    sWhere += " and (";

                    sWhere += "  upper(a.acc_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.acc_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.acc_main_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.acc_main_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(b.acgrp_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(c.actype_name) like '%" + searchstring.ToUpper() + "%'";

                    sWhere += " )";
                }


                if (type == "EXCEL")
                {

                    DataTable Dt_List = new DataTable();
       
                    sql = " select  acc_pkid,acc_main_code, acc_code, acc_name,  ";
                    sql += "  pgr.acgrp_name as main_group,b.acgrp_name as sub_group,";
                    sql += "  c.actype_name, acc_against_invoice, acc_cost_centre,";
                    sql += "  d.note_no,d.main_head,d.sub_head , sub_note  ";
                    sql += "  from acctm a  ";
                    sql += "  left join acgroupm b on a.acc_group_id = b.acgrp_pkid";
                    sql += "  left join acgroupm pgr on b.acgrp_parent_id = pgr.acgrp_pkid  ";
                    sql += "  left join actypem  c on a.acc_type_id  = c.actype_pkid ";
                    sql += "  left join bshead d on a.acc_bs_id = d.pkid ";
                    sql += sWhere;
                    sql += " order by acc_name";

                    sql = sql.Replace("{startrow}", startrow.ToString());
                    sql = sql.Replace("{endrow}", endrow.ToString());

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mRow = new Acctm();
                        mRow.acc_pkid = Dr["acc_pkid"].ToString();
                        mRow.acc_code = Dr["acc_code"].ToString();
                        mRow.acc_name = Dr["acc_name"].ToString();
                        mRow.acc_main_code = Dr["acc_main_code"].ToString();
                        mRow.acc_group_name = Dr["sub_group"].ToString();
                        mRow.acc_main_group_name = Dr["main_group"].ToString();
                        mRow.acc_type_name = Dr["actype_name"].ToString();
                        mRow.acc_against_invoice = Dr["acc_against_invoice"].ToString();
                        mRow.acc_bs_note_no = Dr["note_no"].ToString();
                        mRow.acc_bs_main_head = Dr["main_head"].ToString();
                        mRow.acc_bs_sub_head = Dr["sub_head"].ToString();
                        mRow.acc_bs_sub_note = Dr["sub_note"].ToString();
                        mRow.acc_cost_centre = false;
                        if (Dr["acc_cost_centre"].ToString() == "Y")
                            mRow.acc_cost_centre = true;
                        mList.Add(mRow);
                    }

                    if (mList != null)
                        PrintAccMaster(mList, comp_code);
                }
                else
                {

                    if (type == "NEW")
                    {
                        sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM acctm a ";
                        sql += " left join acgroupm b on a.acc_group_id = b.acgrp_pkid ";
                        sql += " left join actypem  c on a.acc_type_id  = c.actype_pkid ";
                        sql += sWhere;
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
                    sql += "  select  acc_pkid,  acc_code, acc_name, acc_main_code ,";
                    sql += "  b.acgrp_name ,c.actype_name , acc_against_invoice, acc_cost_centre,";
                    sql += "  a.acc_bs_id, d.note_no ||' / '|| d.main_head || ' / ' || d.sub_head  as acc_bs_code, sub_note as acc_bs_name, ";
                    sql += "  row_number() over(order by acc_name) rn ";
                    sql += "  from acctm a  ";
                    sql += " left join acgroupm b on a.acc_group_id = b.acgrp_pkid ";
                    sql += " left join actypem  c on a.acc_type_id  = c.actype_pkid ";
                    sql += " left join bshead d on a.acc_bs_id = d.pkid ";
                    sql += sWhere;
                    sql += ") a ";
                    sql += " where rn between {startrow} and {endrow}";
                    sql += " order by acc_name";

                    sql = sql.Replace("{startrow}", startrow.ToString());
                    sql = sql.Replace("{endrow}", endrow.ToString());

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mRow = new Acctm();
                        mRow.acc_pkid = Dr["acc_pkid"].ToString();
                        mRow.acc_code = Dr["acc_code"].ToString();
                        mRow.acc_name = Dr["acc_name"].ToString();

                        mRow.acc_main_code = Dr["acc_main_code"].ToString();

                        mRow.acc_group_name = Dr["acgrp_name"].ToString();
                        mRow.acc_type_name = Dr["actype_name"].ToString();

                        mRow.acc_against_invoice = Dr["acc_against_invoice"].ToString();

                        mRow.acc_bs_id = Dr["acc_bs_id"].ToString();
                        if (Dr["acc_bs_name"].ToString().Length > 0)
                            mRow.acc_bs_code = Dr["acc_bs_code"].ToString();
                        else
                            mRow.acc_bs_code = "";
                        mRow.acc_bs_name = Dr["acc_bs_name"].ToString();

                        mRow.acc_cost_centre = false;
                        if (Dr["acc_cost_centre"].ToString() == "Y")
                            mRow.acc_cost_centre = true;
                        mList.Add(mRow);
                    }
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            return RetData;
        }

        private void PrintAccMaster(List<Acctm> mList, string comp_code)
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPADD3 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";
            string FolderId = "";
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = "ACCOUNT MASTER";

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "COMP_ADDRESS");
                mSearchData.Add("comp_code", comp_code);
                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPADD3 = Dr["COMP_ADDRESS3"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                FolderId = Guid.NewGuid().ToString().ToUpper();
                File_Display_Name = "AccMaster.xls";
                File_Name = Lib.GetFileName(report_folder, FolderId, File_Display_Name);

                WB = new ExcelFile();
                WB.Worksheets.Add("Report");
                WS = WB.Worksheets["Report"];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 20;
                WS.Columns[2].Width = 256 * 20;
                WS.Columns[3].Width = 256 * 35;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 35;
                WS.Columns[6].Width = 256 * 20;
                WS.Columns[7].Width = 256 * 12;
                WS.Columns[8].Width = 256 * 12;
                WS.Columns[9].Width = 256 * 10;
                WS.Columns[10].Width = 256 * 20;
                WS.Columns[11].Width = 256 * 35;
                WS.Columns[12].Width = 256 * 35;

                iRow = 0; iCol = 1;

                _Size = 14;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 12;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;
                if (str == "")
                    str = COMPADD3;
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, REPORT_CAPTION, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, "MAIN CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SUB CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MAIN GROUP", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SUB GROUP", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ALLOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COST CENTRE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOTE NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MAIN HEAD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SUB HEAD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SUB NOTE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                foreach (Acctm Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_main_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_main_group_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_group_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_type_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_against_invoice, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_cost_centre == true ? "Y" : "N", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_bs_note_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_bs_main_head, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_bs_sub_head, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_bs_sub_note, _Color, false, "", "L", "", _Size, false, 325, "", true);
                }

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Acctm mRow = new Acctm();

            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  acc_pkid,  acc_code, acc_name,acc_main_code, acc_main_id, ";
                sql += " acc_group_id, acc_type_id, acc_against_invoice, acc_cost_centre,acc_taxable,";
                sql += " b.param_name as acc_main_name,acc_branch_code, ";
                sql += " acc_sac_id,sac.param_code as acc_sac_code";
                sql += " ,acc_bs_id, c.note_no ||' / '|| c.main_head || ' / ' || c.sub_head  as acc_bs_code, sub_note as acc_bs_name ";
                sql += " from acctm a ";
                sql += " left join param b on a.acc_main_id = b.param_pkid ";
                sql += " left join param sac on a.acc_sac_id = sac.param_pkid ";
                sql += " left join bshead c on a.acc_bs_id = c.pkid ";
                sql += " where  a.acc_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Acctm();
                    mRow.acc_pkid = Dr["acc_pkid"].ToString();
                    mRow.acc_code = Dr["acc_code"].ToString();
                    mRow.acc_name = Dr["acc_name"].ToString();

                    mRow.acc_main_code = Dr["acc_main_code"].ToString();
                    mRow.acc_main_id = Dr["acc_main_id"].ToString();
                    mRow.acc_main_name = Dr["acc_main_name"].ToString();

                    mRow.acc_group_id = Dr["acc_group_id"].ToString();
                    mRow.acc_type_id = Dr["acc_type_id"].ToString();


                    mRow.acc_branch_code = Dr["acc_branch_code"].ToString();

                    mRow.acc_against_invoice = Dr["acc_against_invoice"].ToString();

                    mRow.acc_cost_centre = false;
                    if (Dr["acc_cost_centre"].ToString() == "Y")
                        mRow.acc_cost_centre = true;

                    mRow.acc_taxable = false;
                    if (Dr["acc_taxable"].ToString() == "Y")
                        mRow.acc_taxable = true;

                    mRow.acc_sac_id = Dr["acc_sac_id"].ToString();
                    mRow.acc_sac_code = Dr["acc_sac_code"].ToString();
                    mRow.acc_bs_id = Dr["acc_bs_id"].ToString();
                    mRow.acc_bs_code = Dr["acc_bs_code"].ToString();
                    mRow.acc_bs_name = Dr["acc_bs_name"].ToString();

                    break;
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            return RetData;
        }


        public string AllValid(Acctm Record)
        {
            string str = "";
            try
            {
                sql = "select acc_pkid from (";
                sql += "select acc_pkid  from acctm a where a.rec_company_code = '{COMPANY_CODE}' ";
                sql += " and (a.acc_code = '{CODE}' or a.acc_name = '{NAME}')  ";
                sql += ") a where acc_pkid <> '{PKID}'";

                sql = sql.Replace("{COMPANY_CODE}", Record._globalvariables.comp_code);
                sql = sql.Replace("{CODE}", Record.acc_code);
                sql = sql.Replace("{NAME}", Record.acc_name);
                sql = sql.Replace("{PKID}", Record.acc_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Acctm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if (Record.acc_code.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Code Cannot Be Empty");

                if (Record.acc_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");

                if (Record.acc_main_code.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Main Code Cannot Be Empty");

                if (Record.acc_group_id.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "A/c Group Cannot Be Empty");

                if (Record.acc_type_id.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "A/c Type Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("acctm", Record.rec_mode, "acc_pkid", Record.acc_pkid);
                Rec.InsertString("acc_code", Record.acc_code);
                Rec.InsertString("acc_name", Record.acc_name);
                Rec.InsertString("acc_group_id", Record.acc_group_id);
                Rec.InsertString("acc_type_id", Record.acc_type_id);
                Rec.InsertString("acc_main_code", Record.acc_main_code);
                if (Record.acc_code == Record.acc_main_code)
                {
                    Rec.InsertString("acc_main_id", Record.acc_pkid);
                    Rec.InsertString("acc_main_name", Record.acc_name);
                }
                else
                {
                    Rec.InsertString("acc_main_id", Record.acc_main_id);
                    Rec.InsertString("acc_main_name", Record.acc_main_name);
                }

                Rec.InsertString("acc_against_invoice", Record.acc_against_invoice);

                Rec.InsertString("acc_branch_code", Record.acc_branch_code);

                if (Record.acc_cost_centre)
                    Rec.InsertString("acc_cost_centre", "Y");
                else
                    Rec.InsertString("acc_cost_centre", "N");

                if (Record.acc_taxable)
                    Rec.InsertString("acc_taxable", "Y");
                else
                    Rec.InsertString("acc_taxable", "N");

                Rec.InsertString("acc_sac_id", Record.acc_sac_id);
                Rec.InsertString("acc_bs_id", Record.acc_bs_id);

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("rec_locked", "N");
                    Rec.InsertString("rec_deleted", "N");
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_created_date", "SYSDATE");
                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_edited_date", "SYSDATE");
                }

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);

                if (Record.acc_main_code.Trim().Length > 0)
                {
                    sql = "update acctm set acc_bs_id ='" + Record.acc_bs_id + "' where rec_company_code ='" + Record._globalvariables.comp_code + "'";
                    sql += " and acc_main_code ='" + Record.acc_main_code + "'";

                    Con_Oracle.ExecuteNonQuery(sql);
                }

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            return RetData;
        }


        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "actypem");
            parameter.Add("comp_code", SearchData["comp_code"].ToString());
            parameter.Add("branch_code", SearchData["branch_code"].ToString());
            RetData.Add("actypem", lovservice.Lov(parameter)["actypem"]);

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "acgroupm");
            parameter.Add("comp_code", SearchData["comp_code"].ToString());
            parameter.Add("branch_code", SearchData["branch_code"].ToString());
            RetData.Add("acgroupm", lovservice.Lov(parameter)["acgroupm"]);

            return RetData;
        }


    }
}
