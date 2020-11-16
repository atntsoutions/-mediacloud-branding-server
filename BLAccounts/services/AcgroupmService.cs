using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLAccounts
{
    public class AcgroupmService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Acgroupm> mList = new List<Acgroupm>();
            Acgroupm mRow;

            string type = SearchData["type"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  a.rec_company_code ='" + comp_code + "' " ;
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.acgrp_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }


                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM Acgroupm a "  ;
                    
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
                sql += "  select  a.acgrp_pkid,  a.acgrp_name, b.acgrp_name as acgrp_parent_name, ";
                sql += "  a.acgrp_order, a.acgrp_drcr,a.acgrp_fixedasset_code,a.acgrp_level, ";
                sql += "  a.acgrp_bs_id, c.note_no ||' / '|| c.main_head || ' / ' || c.sub_head  as acgrp_bs_code, sub_note as acgrp_bs_name, ";
                sql += "  row_number() over(order by a.acgrp_level) rn ";
                sql += "  from Acgroupm a ";
                sql += "  inner join Acgroupm b on a.acgrp_parent_id = b.acgrp_pkid ";
                sql += "  left join bshead c on a.acgrp_bs_id = c.pkid ";
                sql +=  sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by a.acgrp_level";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Acgroupm();
                    mRow.acgrp_pkid = Dr["acgrp_pkid"].ToString();
                    mRow.acgrp_name = Dr["acgrp_name"].ToString();
                    mRow.acgrp_parent_name = Dr["acgrp_parent_name"].ToString();

                    mRow.acgrp_drcr = Dr["acgrp_drcr"].ToString();
                    mRow.acgrp_fixedasset_code = Dr["acgrp_fixedasset_code"].ToString();
                    mRow.acgrp_order  = Lib.Conv2Integer(Dr["acgrp_order"].ToString());
                    mRow.acgrp_level = Lib.Conv2Integer(Dr["acgrp_level"].ToString());
                    mRow.acgrp_bs_id = Dr["acgrp_bs_id"].ToString();
                    if (Dr["acgrp_bs_name"].ToString().Length > 0)
                        mRow.acgrp_bs_code = Dr["acgrp_bs_code"].ToString();
                    else
                        mRow.acgrp_bs_code = "";
                    mRow.acgrp_bs_name = Dr["acgrp_bs_name"].ToString();

                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);

            return RetData;
        }


        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();
            List<Acgroupm> mList = new List<Acgroupm>();
            Acgroupm mRow;

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            try
            {
                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select acgrp_pkid, acgrp_name from acgroupm where rec_company_code = '" + comp_code + "'";
                sql += " and acgrp_parent_id is null order by acgrp_name";

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Acgroupm();
                    mRow.acgrp_pkid = Dr["acgrp_pkid"].ToString();
                    mRow.acgrp_name = Dr["acgrp_name"].ToString();
                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("list", mList);

            return RetData;
        }






        public Dictionary<string, object>  GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Acgroupm mRow =new Acgroupm();
            
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  acgrp_pkid,acgrp_name, acgrp_parent_id";
                sql += ",acgrp_order, acgrp_drcr,acgrp_fixedasset_code, acgrp_level ";
                sql += " ,acgrp_bs_id, b.note_no ||' / '|| b.main_head || ' / ' || b.sub_head  as acgrp_bs_code, sub_note as acgrp_bs_name ";
                sql += " from Acgroupm a  ";
                sql += " left join bshead b on a.acgrp_bs_id = b.pkid ";
                sql += " where  a.acgrp_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Acgroupm ();
                    mRow.acgrp_pkid = Dr["acgrp_pkid"].ToString();
                    mRow.acgrp_name = Dr["acgrp_name"].ToString();
                    mRow.acgrp_parent_id = Dr["acgrp_parent_id"].ToString();

                    mRow.acgrp_drcr = Dr["acgrp_drcr"].ToString();
                    mRow.acgrp_fixedasset_code = Dr["acgrp_fixedasset_code"].ToString();
                    mRow.acgrp_order = Lib.Conv2Integer(Dr["acgrp_order"].ToString());
                    mRow.acgrp_level = Lib.Conv2Integer(Dr["acgrp_level"].ToString());
                    mRow.acgrp_bs_id = Dr["acgrp_bs_id"].ToString();
                    mRow.acgrp_bs_code = Dr["acgrp_bs_code"].ToString();
                    mRow.acgrp_bs_name = Dr["acgrp_bs_name"].ToString();

                    break;
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            return RetData;
        }


        public string AllValid(Acgroupm Record)
        {
            string str = "";
            try
            {
                sql = "select acgrp_pkid from (";
                sql += "select acgrp_pkid  from Acgroupm a where a.rec_company_code = '" + Record._globalvariables.comp_code + "' ";
                sql += " and (a.acgrp_name = '{NAME}')  ";
                sql += ") a where acgrp_pkid <> '{PKID}'";

                sql = sql.Replace("{NAME}", Record.acgrp_name);
                sql = sql.Replace("{PKID}", Record.acgrp_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Acgroupm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();


                if (Record.acgrp_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("acgroupm", Record.rec_mode, "acgrp_pkid", Record.acgrp_pkid);
                Rec.InsertString("acgrp_name", Record.acgrp_name);
                Rec.InsertString("acgrp_parent_id", Record.acgrp_parent_id);
                Rec.InsertString("acgrp_drcr", Record.acgrp_drcr);
                Rec.InsertString("acgrp_fixedasset_code", Record.acgrp_fixedasset_code);
                Rec.InsertString("acgrp_bs_id", Record.acgrp_bs_id);

                Rec.InsertNumeric("acgrp_order", Record.acgrp_order.ToString());

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_deleted", "N");
                    Rec.InsertString("rec_locked", "N");
                    Rec.InsertNumeric("acgrp_level", Record.acgrp_level.ToString());
                }


                sql = Rec.UpdateRow();
                
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);

                if (Record.acgrp_acc_update)
                {
                    sql = "update acctm set ACC_BS_ID ='" + Record.acgrp_bs_id + "' where ACC_GROUP_ID='" + Record.acgrp_pkid + "'";
                    sql += " and rec_company_code='" + Record._globalvariables.comp_code + "'";
                   // sql += " and rec_branch_code='" + Record._globalvariables.branch_code + "'";
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



    }
}
