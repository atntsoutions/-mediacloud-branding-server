using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class TdsPay : table_Base
    {
        public string rowtype { get; set; }
        public string rowcolor { get; set; }

        public string acc_code { get; set; }
        public string jvh_date { get; set; }
        public string jvh_type { get; set; }
        public string jvh_vrno { get; set; }

        public string panno { get; set; }
        public string party_name { get; set; }
        public string location { get; set; }

        public Nullable<decimal> jv_tds_gross_amt { get; set; }


        public Nullable<decimal> jv_tds_rate { get; set; }
        public Nullable<decimal> interest { get; set; }

        public Nullable<decimal> commision { get; set; }

        public Nullable<decimal> contract { get; set; }
        public Nullable<decimal> rent { get; set; }
        public Nullable<decimal> building { get; set; }
        public Nullable<decimal> salary { get; set; }
        public Nullable<decimal> forgnpay { get; set; }
        public Nullable<decimal> ptax { get; set; }
        public Nullable<decimal> jv_credit { get; set; }
       

    }
}
