using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;


namespace BLReport1 
{
    public class MappingService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        List<Mappingm> mList = new List<Mappingm>();
        Mappingm mrow;
        string ErrorMessage = "";
        string cntr_no = "";
        string id = "";

        public IDictionary<string, object> MappingList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<Mappingm>();
            ErrorMessage = "";
            string Table_Name= "";
            string branch_code = "";
            try
            {
                Table_Name = SearchData["tablename"].ToString();
                branch_code = SearchData["branch_code"].ToString();

                Con_Oracle = new DBConnection();

                sql = " select pkid,br_code,table_name,source_col,target_col,slno ";
                sql += "  from mappingm a";
                sql += "  where table_name = '{TABLE}'";
                sql += "  and br_code = '{BRCODE}'";
                sql += "  order by slno";
                sql = sql.Replace("{TABLE}", Table_Name);
                sql = sql.Replace("{BRCODE}", branch_code);
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new Mappingm();
                    mrow.rec_mode = "EDIT";
                    mrow.pkid = Dr["pkid"].ToString();
                    mrow.br_code = Dr["br_code"].ToString();
                    mrow.table_name = Dr["table_name"].ToString();
                    mrow.source_col = Dr["source_col"].ToString();
                    mrow.target_col = Dr["target_col"].ToString();
                    mrow.slno = Lib.Conv2Integer(Dr["slno"].ToString());
                    mList.Add(mrow);
                }

                Dt_List.Rows.Clear();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("list", mList);
            return RetData;
        }


        public Dictionary<string, object> Save(Mappingm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {

                Con_Oracle = new DBConnection();

                if ((ErrorMessage = AllValid(Record)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                DBRecord Rec = new DBRecord();
                
                Con_Oracle.BeginTransaction();
                sql = "delete from  mappingm where table_name ='" + Record.table_name + "' and br_code = '" + Record._globalvariables.branch_code + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                string pcntr_pkid = "";
                foreach (Mappingm Row in Record.MappingList)
                {
                    pcntr_pkid = Guid.NewGuid().ToString().ToUpper();
                    Rec = new DBRecord();
                    Rec.CreateRow("mappingm", "ADD", "pkid", Row.pkid);
                    Rec.InsertString("br_code", Row.br_code);
                    Rec.InsertString("table_name", Row.table_name);
                    Rec.InsertString("source_col", Row.source_col);
                    Rec.InsertString("target_col", Row.target_col);
                    Rec.InsertNumeric("slno", Row.slno.ToString());
                    sql = Rec.UpdateRow();
                    Con_Oracle.ExecuteNonQuery(sql);
                }

                Con_Oracle.CommitTransaction();
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

            return RetData;
        }


        public string AllValid(Mappingm Record)
        {
            string str = "";
            try
            {
               
                /*
                if (Record.rec_mode == "ADD")
                {
                    if (!Lib.IsValidSalesman(Record.job_exp_id))
                    {
                        Lib.AddError(ref str, " Shipper Is Assigned With An Invalid Sales Person ");
                    }
                }
                */
            }
            catch (Exception Ex)
            {
                Lib.AddError(ref str, Ex.Message.ToString());
            }
            return str;
        }

    }
}
