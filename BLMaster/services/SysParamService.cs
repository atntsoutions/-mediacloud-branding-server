using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase.Connections;

namespace BLMaster
{
    public class SysParamService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Param> mList = new List<Param>();
            Param mRow;

            string company_code = SearchData["company_code"].ToString();
            string type = SearchData["type"].ToString();
            string param_type = SearchData["rowtype"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string sortby = SearchData["sortby"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where a.rec_company_code = '{COMPANY_CODE}' and param_type = 'PARAM' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.param_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.param_name) like '%" + searchstring.ToUpper() + "%'";

                    sWhere += " )";
                }

                sWhere = sWhere.Replace("{COMPANY_CODE}", company_code);
                sWhere = sWhere.Replace("{PARAM_TYPE}", param_type);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM param  a ";
                    if (Con_Oracle.DB == "SQL")
                        sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM param  a ";

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
                sql += "  select  param_pkid, param_type, param_code, param_name,   row_number() over(order by param_name) rn ";
                sql += "  from param a " + sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                if (sortby == "CODE")
                    sql += " order by param_code";
                else
                    sql += " order by param_name";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Param();
                    mRow.param_pkid = Dr["param_pkid"].ToString();
                    mRow.param_type = Dr["param_type"].ToString();
                    mRow.param_code = Dr["param_code"].ToString();
                    mRow.param_name = Dr["param_name"].ToString();
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

            List<paramvalues> mList = new List<paramvalues>();
            paramvalues mRow = new paramvalues();

            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select keys, defvalue, filetype,impexp,edifile, param_value, param_format ";
                sql += " from paramkeys a left join ";
                sql += " (select param_key, param_value, param_format from paramvalues where parent_id = '" + id + "')  ";
                sql += " b on a.keys = b.param_key ";
                sql += " order by slno ";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new paramvalues();
                    mRow.param_key = Dr["keys"].ToString();
                    mRow.param_value = Dr["param_value"].ToString();
                    mRow.param_defvalue = Dr["defvalue"].ToString();

                    mRow.param_filetype  = Dr["filetype"].ToString();
                    mRow.param_impexp = Dr["impexp"].ToString();
                    mRow.param_edifile = Dr["edifile"].ToString();
                    mRow.param_format = Dr["param_format"].ToString();

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

        public Dictionary<string, object> Save(paramvalues_vm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {
                Con_Oracle = new DBConnection();

                DBRecord Rec = new DBRecord();

                Con_Oracle.BeginTransaction();

                sql = "delete from paramvalues where parent_id ='" + Record.param_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                foreach (paramvalues _Record in Record.RecordDet)
                {

                    if (_Record.param_value.ToString().Length > 0)
                    {
                        _Record.param_pkid = System.Guid.NewGuid().ToString().ToUpper();
                        Rec.CreateRow("paramvalues", "ADD", "param_pkid", _Record.param_pkid);
                        Rec.InsertString("parent_id", Record.param_pkid);
                        Rec.InsertString("param_key", _Record.param_key);
                        Rec.InsertString("param_value", _Record.param_value, "P");

                        Rec.InsertString("param_filetype", _Record.param_filetype);
                        Rec.InsertString("param_impexp", _Record.param_impexp);
                        Rec.InsertString("param_edifile", _Record.param_edifile);

                        Rec.InsertString("param_format", _Record.param_format);

                        Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                        sql = Rec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
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
