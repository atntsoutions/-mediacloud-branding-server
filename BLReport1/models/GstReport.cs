using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class GstReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string jvh_docno { get; set; }
        public string jvh_cc_category { get; set; }
        public string jvh_type { get; set; }
        public string jvh_date { get; set; }
        public string jvh_date_gstr1 { get; set; }
        public string jvh_gstin { get; set; }
        public string jvh_org_invno { get; set; }
        public string jvh_org_invdt { get; set; }
        public Nullable<decimal> jvh_net_amt { get; set; }
        public Nullable<decimal> jvh_tot_amt { get; set; }
        public Nullable<decimal> jv_credit { get; set; }
        public Nullable<decimal> jv_net_total { get; set; }
        public Nullable<decimal> jv_taxable_amt { get; set; }
        public string jvh_gst_type { get; set; }
        public Nullable<decimal> jv_gst_rate { get; set; }
        public Nullable<decimal> jv_gst_amt { get; set; }
        public Nullable<decimal> jv_cgst_amt { get; set; }
        public Nullable<decimal> jv_sgst_amt { get; set; }
        public Nullable<decimal> jv_igst_amt { get; set; }
        public Nullable<decimal> jv_cgst_rate { get; set; }
        public Nullable<decimal> jv_sgst_rate { get; set; }
        public Nullable<decimal> jv_igst_rate { get; set; }
        public string jvh_sez { get; set; }
        public string jvh_state_name { get; set; }
        public string rc { get; set; }
        public string jvh_invoice_type { get; set; }
        public string ecomgstn { get; set; }
        public Nullable<decimal> cess { get; set; }
        public string jvh_gst { get; set; }
        public string jvh_party_name { get; set; }
        public string jv_sac_code { get; set; }
        public string jv_acc_code { get; set; }
        public string jv_acc_name { get; set; }
        public Nullable<decimal> inv_amt { get; set; }
        public Nullable<decimal> taxable_amt { get; set; }

        public string branch { get; set; }
    }
}
