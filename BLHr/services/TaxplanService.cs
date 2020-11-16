using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;


using DataBase_Oracle.Connections;

namespace BLHr
{
    public class TaxplanService : BL_Base
    {
      
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            string sWhere = "";
           
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<TaxPlan> mList = new List<TaxPlan>();
            TaxPlan mRow;

            string type = SearchData["type"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  a.rec_company_code = '{COMPCODE}' and a.tp_year =  {FYEAR}";

                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += " upper(a.tp_desc) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += " upper(a.tp_group_ctr) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{FYEAR}", year_code);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM taxplan  a ";                  
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
                sql += " select tp_pkid,tp_year,tp_group_ctr,tp_ctr,tp_desc,tp_limit,tp_editable,tp_bold, ";
                sql += " row_number() over(order by tp_group_ctr,tp_ctr) rn ";
                sql += " from taxplan a ";
               
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow} ";
                sql += " order by tp_group_ctr,tp_ctr";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new TaxPlan();
                    mRow.tp_pkid = Dr["tp_pkid"].ToString();
                    mRow.tp_year = Lib.Conv2Integer(Dr["tp_year"].ToString());
                    mRow.tp_group_ctr = Lib.Conv2Integer(Dr["tp_group_ctr"].ToString());
                    mRow.tp_ctr = Lib.Conv2Integer(Dr["tp_ctr"].ToString());
                    mRow.tp_desc = Dr["tp_desc"].ToString();
                    mRow.tp_limit = Lib.Conv2Integer(Dr["tp_limit"].ToString());

                    if (Dr["tp_editable"].ToString() == "Y")
                        mRow.tp_editable = true;
                    else
                        mRow.tp_editable = false;

                    if (Dr["tp_bold"].ToString() == "Y")
                        mRow.tp_bold = true;
                    else
                        mRow.tp_bold = false;

                    mList.Add(mRow);
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

            return RetData;
        }


        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            TaxPlan mRow = new TaxPlan();

            string id = SearchData["pkid"].ToString();
         

            try
            {
                Con_Oracle = new DBConnection();

                sql = " select tp_pkid,tp_year,tp_group_ctr,tp_ctr,tp_desc,tp_limit,tp_editable,";
                sql += " tp_bold from taxplan a ";              
                sql += " where a.tp_pkid = '" + id + "'";             
                sql += " order by a.tp_year,a.tp_group_ctr,a.tp_ctr ";
            
                DataTable Dt_Rec = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                   

                    mRow.tp_pkid = Dr["tp_pkid"].ToString();
                    mRow.tp_year = Lib.Conv2Integer(Dr["tp_year"].ToString());
                    mRow.tp_group_ctr = Lib.Conv2Integer(Dr["tp_group_ctr"].ToString());
                    mRow.tp_ctr = Lib.Conv2Integer(Dr["tp_ctr"].ToString());
                    mRow.tp_desc = Dr["tp_desc"].ToString();
                    mRow.tp_limit = Lib.Conv2Integer(Dr["tp_limit"].ToString());

                    if (Dr["tp_editable"].ToString() == "Y")
                        mRow.tp_editable = true;
                    else
                        mRow.tp_editable = false;
                    
                   
                    if (Dr["tp_bold"].ToString() == "Y") 
                        mRow.tp_bold = true;                    
                    else 
                        mRow.tp_bold = false;
                    
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


        public string AllValid(TaxPlan Record)
        { 
            string str = "";
            
            return str;
        }

        
        public Dictionary<string, object> Save(TaxPlan Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("taxplan", Record.rec_mode, "tp_pkid", Record.tp_pkid);
                Rec.InsertString("tp_desc", Record.tp_desc);
                Rec.InsertNumeric("tp_limit", Lib.Conv2Decimal(Record.tp_limit.ToString()).ToString());
                Rec.InsertNumeric("tp_group_ctr", Lib.Conv2Decimal(Record.tp_group_ctr.ToString()).ToString());
                Rec.InsertNumeric("tp_ctr", Lib.Conv2Decimal(Record.tp_ctr.ToString()).ToString());

                if (Record.tp_editable)
                    Rec.InsertString("tp_editable", "Y");
                else
                    Rec.InsertString("tp_editable", "N");

                if (Record.tp_bold)
                    Rec.InsertString("tp_bold", "Y");
                else
                    Rec.InsertString("tp_bold", "N");

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("tp_year", Record._globalvariables.year_code);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
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
            Con_Oracle.CloseConnection();
            return RetData;
        }
        
    }
}
