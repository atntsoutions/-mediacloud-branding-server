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
    public class TrackService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        List<TrackReport> mList = new List<TrackReport>();
        TrackReport mrow;
        string ErrorMessage = "";
        string cntr_no = "";
        string id = "";

        public IDictionary<string, object> TrackList(string id, string container_no, string house_bl_no)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<TrackReport>();
            ErrorMessage = "";
            string MBL_ID = "";
            try
            {
                if (id != "2018")
                {
                    RetData.Add("list", mList);
                    return RetData;
                }



                Con_Oracle = new DBConnection();
                if (container_no.Length <= 0 && house_bl_no.Length <= 0)
                    Lib.AddError(ref ErrorMessage, " | Either Container Number or House BL Number cannot be blank.");

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (container_no.Length > 10)
                {
                    cntr_no = container_no.Replace(" ", "");
                    cntr_no = cntr_no.Replace("-", "");
                    cntr_no = cntr_no.Trim();
                    cntr_no = cntr_no.Insert(10, "-");

                    sql = " select distinct h.hbl_mbl_id from jobm a";
                    sql += "   inner join hblm h on a.jobs_hbl_id = h.hbl_pkid ";
                    sql += "   inner join packingm b on a.job_pkid = b.pack_job_id ";
                    sql += "   inner join containerm c on b.pack_cntr_id = c.cntr_pkid ";
                    sql += "   where cntr_no = '{CNTRNO}' ";
                    sql = sql.Replace("{CNTRNO}", cntr_no);

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);

                    if (Dt_List.Rows.Count > 0)
                        MBL_ID = Dt_List.Rows[0]["hbl_mbl_id"].ToString();

                }

                if (MBL_ID == "" && house_bl_no.Length > 0)
                {
                    sql = "select hbl_mbl_id from hblm where hbl_type='HBL-SE' and trim(hbl_bl_no)='{BLNO}'";
                    sql = sql.Replace("{BLNO}", house_bl_no.Trim());
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_List.Rows.Count > 0)
                        MBL_ID = Dt_List.Rows[0]["hbl_mbl_id"].ToString();

                }

                if (MBL_ID != "")
                {
                    sql = " select vsl.param_name as vessel";
                    sql += "  ,trk_voyage as voyage";
                    sql += "  ,pol.param_name as pol_name";
                    sql += "  ,trk_pol_etd as pol_etd";
                    sql += "  ,trk_pol_etd_confirm as pol_etd_confirm";
                    sql += "  ,pod.param_name as pod_name";
                    sql += "  ,trk_pod_eta as pod_eta";
                    sql += "  ,trk_pod_eta_confirm as pod_eta_confirm";
                    sql += "  ,trk_order as trk_order ";
                    sql += "  ,row_number() over (order by trk_order) as slno";
                    sql += "  from trackingm t";
                    sql += "  inner join param  vsl on (t.trk_vsl_id=vsl.param_pkid)";
                    sql += "  inner join param pol on(t.trk_pol_id=pol.param_pkid)";
                    sql += "  inner join param pod on(t.trk_pod_id=pod.param_pkid)";
                    sql += "  where t.trk_parent_id ='{MBLID}'";
                    sql += "  order by trk_order";
                    sql = sql.Replace("{MBLID}", MBL_ID);
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TrackReport();
                        mrow.sl_no = Lib.Conv2Integer(Dr["slno"].ToString());
                        mrow.vessel_name = Dr["vessel"].ToString();
                        mrow.voyage = Dr["voyage"].ToString();
                        mrow.pol_name = Dr["pol_name"].ToString();
                        mrow.pol_etd = Lib.DatetoStringDisplayformat(Dr["pol_etd"]);
                        mrow.pol_etd_confirm = Dr["pol_etd_confirm"].ToString() == "Y" ? "YES" : "NO";
                        mrow.pod_name = Dr["pod_name"].ToString();
                        mrow.pod_eta = Lib.DatetoStringDisplayformat(Dr["pod_eta"]);
                        mrow.pod_eta_confirm = Dr["pod_eta_confirm"].ToString() == "Y" ? "YES" : "NO";
                        mList.Add(mrow);
                    }
                }
                else
                {
                    Lib.AddError(ref ErrorMessage, " | Tracking Deatils Not Found.");
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
    }
}

