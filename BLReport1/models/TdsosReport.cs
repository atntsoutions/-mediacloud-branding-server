using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;


namespace BLReport1
{
    public class TdsosReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string branch { get; set; }
        public string party_name { get; set; }
        public string party_code { get; set; }
        public string jvh_docno { get; set; }
        public string jvh_date { get; set; }
        public string sman_name { get; set; }
        public string tan_code { get; set; }
        public string tan_name { get; set; }
        public string tds_cert_no { get; set; }
        public Nullable<decimal> tds_amt { get; set; }
        public Nullable<decimal> collected_amt { get; set; }
        public Nullable<decimal> pending_amt { get; set; }
        public Nullable<decimal> q1_amt { get; set; }
        public Nullable<decimal> q2_amt { get; set; }
        public Nullable<decimal> q3_amt { get; set; }
        public Nullable<decimal> q4_amt { get; set; }
    }
}
