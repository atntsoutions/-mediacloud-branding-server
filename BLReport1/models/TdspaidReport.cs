using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class TdspaidReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string branch_code { get; set; }
        public string jv_pkid { get; set; }
        public string jv_vrno { get; set; }
        public string jv_type { get; set; }
        public string jv_date { get; set; }
        public string party_code { get; set; }
        public string party_name { get; set; }
        public string tan_id { get; set; }
        public string tan_code { get; set; }
        public string tan_name { get; set; }
        public string tds_cert_no { get; set; }
        public string tds_cert_qtr { get; set; }
        public string cert_recvd_at { get; set; }
        public Nullable<decimal> gross_bill_amt { get; set; }
        public Nullable<decimal> gross_cert_amt { get; set; }
        public Nullable<decimal> jv_credit { get; set; }
        public Nullable<decimal> tds_amt { get; set; }
        public Nullable<decimal> cert_amt { get; set; }
        public Nullable<decimal> cert_alloc_amt { get; set; }
        public Nullable<decimal> pending_amt { get; set; }
        public string asd_section { get; set; }
        public string asd_trans_date { get; set; }
        public string asd_book_date { get; set; }
        public Nullable<decimal> asd_gross { get; set; }
        public Nullable<decimal> asd_deducted { get; set; }
        public Nullable<decimal> asd_tds { get; set; }
        public Nullable<decimal> asm_gross { get; set; }
        public Nullable<decimal> asm_deducted { get; set; }
        public Nullable<decimal> asm_tds { get; set; }
    }
}
