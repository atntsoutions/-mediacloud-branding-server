using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class PayHistroyReport : table_Base
    {
        public string pkid { get; set; }
        public string rowtype { get; set; }
        public string rowcolor { get; set; }


        public string acc_name { get; set; }
        public string sl_no { get; set; }
        public string jvh_vrno { get; set; }
        public string jvh_type { get; set; }
        public string jvh_date { get; set; }
        public Nullable<decimal> jv_debit { get; set; }
        public string cr_date { get; set; }
        public Nullable<decimal> intrest { get; set; }
        public Nullable<decimal> days { get; set; }
        public Nullable<decimal> xref_amt { get; set; }
        public Nullable<decimal> cr_total { get; set; }
        public Nullable<decimal> bal_days { get; set; }
        public Nullable<decimal> balance { get; set; }
        public Nullable<decimal> int1 { get; set; }
        public Nullable<decimal> int2 { get; set; }
        public string branch { get; set; }


        public string acc_code { get; set; }
        public Nullable<decimal> jv_credit { get; set; }
        public string xref_crdate { get; set; }
        public Nullable<decimal> cust_crlimit { get; set; }
        public Nullable<decimal> cust_crdays { get; set; }
        public string sman_name { get; set; }
        public Nullable<decimal> pending { get; set; }
        public string status { get; set; }
        public Nullable<decimal> overdue { get; set; }

    }

    public class CollectionReport : table_Base
    {
        public string row_type { get; set; }
        public string row_color { get; set; }
        public string desc1 { get; set; }
        public string desc2 { get; set; }


        public string col1 { get; set; }
        public string col2 { get; set; }
        public string col3 { get; set; }
        public string col4 { get; set; }


    }

    }
