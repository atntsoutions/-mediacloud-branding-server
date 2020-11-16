using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class AcTransReport : table_Base
    {
        public string pkid { get; set; }
        public string rowtype { get; set; }
        public string rowcolor { get; set; }
        public string jvh_vrno { get; set; }
        public string jvh_type { get; set; }
        public string jvh_date { get; set; }
        public string acc_code { get; set; }
        public string acc_name { get; set; }
        public string narration { get; set; }
        public string rec_createdby { get; set; }
        public string rec_createddate { get; set; }
        public string rec_editedby { get; set; }
        public string rec_editeddate { get; set; }

        public Nullable<decimal> jv_debit { get; set; }
        public Nullable<decimal> jv_credit { get; set; }

        public string jvh_pkid { get; set; }
        public Nullable<int> slno { get; set; }
        public string jvh_docno { get; set; }
        public string xref_pkid { get; set; }
        public string xref_no { get; set; }
        public string xref_date { get; set; }
        public Nullable<decimal> xref_amt { get; set; }

    }
}
