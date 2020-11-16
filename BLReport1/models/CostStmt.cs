using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class CostStmtReport : table_Base
    {

        public string rowtype { get; set; }
        public string rowcolor { get; set; }
        public string roworder2 { get; set; }
        public string jv_pkid { get; set; }
        public string jvh_date { get; set; }
        public string curr_code { get; set; }
        public string jvh_vrno { get; set; }
        public string jvh_type { get; set; }
        public string jvh_remarks { get; set; }
        public string jvh_reference { get; set; }
        public string reccategory { get; set; }

        public Nullable<decimal> jv_debit { get; set; }
        public Nullable<decimal> jv_credit { get; set; }
        public Nullable<decimal> jv_exrate { get; set; }
        public Nullable<decimal> inr { get; set; }
        public Nullable<decimal> opening { get; set; }

        public string branch { get; set; }

    }

}
