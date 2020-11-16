using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class BrokReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string branch { get; set; }
        public string jvh_vrno { get; set; }
        public string jvh_date { get; set; }
        public string hbl_no { get; set; }
        public string hbl_bl_no { get; set; }
        public string hbl_date { get; set; }
        public string vessel { get; set; }
        public string hbl_vessel_no { get; set; }
        public string hbl_terms { get; set; }
        public string hbl_nature { get; set; }
        public string carrier { get; set; }
        public string agent { get; set; }
        public string shipper { get; set; }
        public string consignee { get; set; }
        public string jvh_reference { get; set; }
        public string jvh_reference_date { get; set; }
        public string jvh_org_invno { get; set; }
        public string jvh_org_invdt { get; set; }
        public Nullable<decimal> jvh_basic_frt { get; set; }
        public Nullable<decimal> jvh_brok_per { get; set; }
        public Nullable<decimal> jvh_brok_amt { get; set; }
        public Nullable<decimal> jvh_brok_exrate { get; set; }
        public Nullable<decimal> jvh_brok_amt_inr { get; set; }
        
    }

}
