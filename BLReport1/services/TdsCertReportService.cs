using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;


namespace BLReport1
{
    public class TdsCertReportService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<TdsCertm> mList = new List<TdsCertm>();
            TdsCertm mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string company_code = SearchData["company_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where a.rec_company_code = '" + company_code + "'";
                sWhere += " and a.tds_year = " + year_code + "";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(tds_cert_no) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  tds_cert_qtr like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  b.param_code like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  b.param_name like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " ) ";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM tdscertm a ";
                    sql += " left join param b on a.tds_tan_id = b.param_pkid ";
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
                sql += " select tds_pkid,tds_cert_no,tds_cert_qtr,tds_cert_brcode";
                sql += " ,tds_gross,tds_amt,tds_doc_count";
                sql += " ,b.param_code as tds_tan_code,b.param_name as tds_tan_name";
                sql += " ,a.rec_created_by,a.rec_branch_code";
                sql += " ,row_number() over(order by tds_cert_no,tds_cert_qtr) rn ";
                sql += " from tdscertm a";
                sql += " left join param b on a.tds_tan_id = b.param_pkid";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by tds_tan_code,tds_cert_no,tds_cert_qtr";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new TdsCertm();
                    mRow.tds_pkid = Dr["tds_pkid"].ToString();
                    mRow.tds_cert_no = Dr["tds_cert_no"].ToString();
                    mRow.tds_cert_qtr = Dr["tds_cert_qtr"].ToString();
                    mRow.tds_cert_brcode = Dr["tds_cert_brcode"].ToString();
                    mRow.tds_tan_code = Dr["tds_tan_code"].ToString();
                    mRow.tds_tan_name = Dr["tds_tan_name"].ToString();
                    mRow.tds_doc_count = Lib.Conv2Integer(Dr["tds_doc_count"].ToString());
                    mRow.tds_gross = Lib.Conv2Decimal(Dr["tds_gross"].ToString());
                    mRow.tds_amt = Lib.Conv2Decimal(Dr["tds_amt"].ToString());
                    mRow.rec_branch_code = Dr["rec_branch_code"].ToString();
                    mRow.rec_created_by = Dr["rec_created_by"].ToString();
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
            TdsCertm mRow = new TdsCertm();
            string id = SearchData["pkid"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select tds_pkid,tds_cert_no,tds_cert_qtr,tds_cert_brcode ,c.comp_name as tds_cert_brname";
                sql += " ,tds_tan_id,tds_year,tds_gross,tds_amt,tds_doc_count ";
                sql += " ,b.param_code as tds_tan_code,b.param_name as tds_tan_name";
                sql += " from tdscertm a  ";
                sql += " left join param b on a.tds_tan_id = b.param_pkid";
                sql += " left join companym c on a.tds_cert_brcode = c.comp_code and c.comp_type='B'";
                sql += " where  a.tds_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new TdsCertm();
                    mRow.tds_pkid = Dr["tds_pkid"].ToString();
                    mRow.tds_cert_no = Dr["tds_cert_no"].ToString();
                    mRow.tds_cert_qtr = Dr["tds_cert_qtr"].ToString();
                    mRow.tds_cert_brcode = Dr["tds_cert_brcode"].ToString();
                    mRow.tds_cert_brname = Dr["tds_cert_brname"].ToString();
                    mRow.tds_tan_id = Dr["tds_tan_id"].ToString();
                    mRow.tds_tan_code = Dr["tds_tan_code"].ToString();
                    mRow.tds_tan_name = Dr["tds_tan_name"].ToString();
                    mRow.tds_year = Lib.Conv2Integer(Dr["tds_year"].ToString());
                    mRow.tds_doc_count = Lib.Conv2Integer(Dr["tds_doc_count"].ToString());
                    mRow.tds_gross = Lib.Conv2Decimal(Dr["tds_gross"].ToString());
                    mRow.tds_amt = Lib.Conv2Decimal(Dr["tds_amt"].ToString());
                    break;
                }
                mRow.TdsDetList = GetDetList(id, mRow.tds_tan_id, year_code);

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


        public string AllValid(TdsCertm Record)
        {
            string str = "";
            DateTime tdate = DateTime.Now;

            try
            {
                if (Record.tds_cert_no.Trim().Length <= 0)
                    Lib.AddError(ref str, " | Certificate Cannot Be Empty");

                sql = "select tds_pkid from (";
                sql += "select tds_pkid from tdscertm a where a.rec_company_code = '{COMPCODE}' ";
                sql += " and a.tds_cert_no = '{CERT_NO}' ";
                sql += " and a.tds_year = {TDSYEAR} ";
                sql += ") a where tds_pkid <> '{PKID}'";
                sql = sql.Replace("{CERT_NO}", Record.tds_cert_no);
                sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                sql = sql.Replace("{TDSYEAR}", Record._globalvariables.year_code);
                sql = sql.Replace("{PKID}", Record.tds_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " | This Certificate No already Exists");
              
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(TdsCertm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            List<TdsCertd> mList = new List<TdsCertd>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();
                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("tdscertm", Record.rec_mode, "tds_pkid", Record.tds_pkid);
                Rec.InsertString("tds_cert_no", Record.tds_cert_no);
                Rec.InsertString("tds_cert_qtr", Record.tds_cert_qtr);
                Rec.InsertString("tds_cert_brcode", Record.tds_cert_brcode);
                Rec.InsertString("tds_tan_id", Record.tds_tan_id);
                Rec.InsertNumeric("tds_gross", Record.tds_gross.ToString());
                Rec.InsertNumeric("tds_amt", Record.tds_amt.ToString());
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("tds_year", Record._globalvariables.year_code);
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

                sql = "delete from tdscertd  where tdsd_parent_id ='" + Record.tds_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                foreach (TdsCertd Row in Record.TdsDetList)
                {
                    if (Row.tdsd_amt > 0)
                    {
                        Row.tdsd_pkid = Guid.NewGuid().ToString().ToUpper();

                        Rec.CreateRow("tdscertd", "ADD", "tdsd_pkid", Row.tdsd_pkid);
                        Rec.InsertString("tdsd_parent_id", Record.tds_pkid);
                        Rec.InsertString("tdsd_tan_id", Record.tds_tan_id);
                        Rec.InsertString("tdsd_jv_id", Row.tdsd_jv_id);
                        Rec.InsertNumeric("tdsd_amt", Row.tdsd_amt.ToString());
                        sql = Rec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                }
                Con_Oracle.CommitTransaction();

                if (Record.rec_mode == "ADD")
                    mList = GetDetList(Record.tds_pkid, Record.tds_tan_id, Record._globalvariables.year_code);
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
            RetData.Add("list", mList);
            return RetData;
        }

        public IDictionary<string, object> TdsDetList(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Con_Oracle = new DBConnection();
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            List<TdsCertd> mList = new List<TdsCertd>();

            try
            {

                string id = SearchData["pkid"].ToString();
                string tanid = SearchData["tanid"].ToString();
                string year_code = SearchData["year_code"].ToString();
                mList = GetDetList(id, tanid, year_code);

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

        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //parameter.Add("comp_code", comp_code);
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }

        private List<TdsCertd> GetDetList(string parent_id, string tan_id, string year_code)
        {

            List<TdsCertd> mList = new List<TdsCertd>();
            TdsCertd Row;


            sql = " select a.rec_branch_code,jv_pkid,  jvh_vrno as jv_vrno, a.jvh_type as jv_type ,jvh_date as jv_date, ";
            sql += "  cust.cust_code as party_code, ";
            sql += "  cust.cust_name as party_name,";
            sql += "  tan.param_code as tan_code,";
            sql += "  tan.param_name as tan_name,";
            sql += "  a.jvh_reference as jv_reference,jvh_narration as jv_narration,jv_total , other_amt, tds_amt  ";
            sql += "  from ledgerh a inner join ledgert b on jvh_pkid = b.jv_parent_id ";
            sql += "  left join (";
            sql += "  select x.tdsd_jv_id,";
            sql += "  sum( case when tdsd_parent_id <> '" + parent_id + "' then tdsd_amt else 0 end ) as other_amt,";
            sql += "  sum( case when tdsd_parent_id = '" + parent_id + "' then tdsd_amt else 0 end ) as tds_amt";
            sql += "  from tdscertd x ";
            sql += "  where x.tdsd_tan_id ='" + tan_id + "' ";
            sql += "  group by tdsd_jv_id";
            sql += "  ) xref on ( b.jv_pkid = xref.tdsd_jv_id )";
            sql += " left join customerm cust on b.jv_tan_party_id = cust.cust_pkid";
            sql += " left join param tan on b.jv_tan_id = tan.param_pkid";
            sql += "  where  jv_tan_id = '" + tan_id + "' and a.jvh_year =" + year_code;
            sql += "  and b.jv_total != 0 and  (b.jv_total - nvl(other_amt,0)) >0 and a.jvh_type not in('OP','OB','OC') and jv_debit > 0";
            sql += "  order by a.jvh_date,a.jvh_type";

            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            decimal TdsAmt = 0, AllocAmt = 0, BalAmt = 0;

            foreach (DataRow Dr in Dt_Rec.Rows)
            {
                Row = new TdsCertd();
                Row.tdsd_jv_id = Dr["jv_pkid"].ToString();
                Row.tdsd_jv_branch = Dr["rec_branch_code"].ToString();
                Row.tdsd_jv_no = Dr["jv_vrno"].ToString();
                Row.tdsd_jv_type = Dr["jv_type"].ToString();
                Row.tdsd_jv_date = Lib.DatetoStringDisplayformat(Dr["jv_date"]);
                Row.tdsd_jv_ref = Dr["jv_reference"].ToString();
                Row.tdsd_jv_narration = Dr["jv_narration"].ToString();

                TdsAmt = Lib.Conv2Decimal(Dr["jv_total"].ToString());
                AllocAmt= Lib.Conv2Decimal(Dr["other_amt"].ToString());
                BalAmt = TdsAmt - AllocAmt;

                Row.tdsd_jv_total = TdsAmt;
                Row.tdsd_other_amt = AllocAmt;
                Row.tdsd_bal_amt = BalAmt;

                Row.tdsd_amt = Lib.Conv2Decimal(Dr["tds_amt"].ToString());//tds Cert Amt
                Row.tdsd_party_code = Dr["party_code"].ToString();
                Row.tdsd_party_name = Dr["party_name"].ToString();
                Row.tdsd_tan_code = Dr["tan_code"].ToString();
                Row.tdsd_tan_name = Dr["tan_name"].ToString();
                Row.tdsd_tan_id = tan_id;

                mList.Add(Row);
            }
            return mList;
        }

    }
}
