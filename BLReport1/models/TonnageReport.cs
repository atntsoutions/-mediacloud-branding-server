using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class TonnageReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string mbl_pkid { get; set; }
        public string mbl_date { get; set; }
        public string mbl_no { get; set; }
        public Nullable<decimal> mbl_grwt { get; set; }
        public Nullable<decimal> mbl_chwt { get; set; }
        public Nullable<decimal> hbl_grwt { get; set; }
        public Nullable<decimal> hbl_chwt { get; set; }
        public string hbl_shipper_name { get; set; }
        public string hbl_consignee_name { get; set; }
        public string hbl_nomination { get; set; }
        public string mbl_agent_name { get; set; }
        public string mbl_airline_name { get; set; }
        public string hbl_pod_name { get; set; }
        public string hbl_pofd_name { get; set; }
        public string mbl_status_name { get; set; }
        public string sman_name { get; set; }
        
        public string hbl_pol_name { get; set; }
        public string branch { get; set; }
    }
}
