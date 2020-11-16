using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class BillWise : table_Base
    {
        public string rowtype { get; set; }
        public string rowcolor { get; set; }

        public string jvh_date { get; set; }
        public string jvh_vrno { get; set; }
        public string jvh_type { get; set; }
        public string jvh_sez { get; set; }
        public string acc_name { get; set; }

        public string jvh_gstin { get; set; }

        public string jvh_rc { get; set; }

        public string jvh_gst_type { get; set; }
        public string jvh_cc_category { get; set; }
        public string hbl_no { get; set; }



        public Nullable<decimal> jvh_tot_amt { get; set; }
        public Nullable<decimal> jvh_cgst_amt { get; set; }
        public Nullable<decimal> jvh_sgst_amt { get; set; }
        public Nullable<decimal> jvh_igst_amt { get; set; }
        public Nullable<decimal> jvh_gst_amt { get; set; }
        public Nullable<decimal> jvh_net_amt { get; set; }

        public string jvh_docno { get; set; }
        public string consignee { get; set; }
        public string pod { get; set; }
        public string volume { get; set; }

        public Nullable<decimal> hbl_ntwt { get; set; }
        public Nullable<decimal> hbl_grwt { get; set; }
        public Nullable<decimal> jv_frt { get; set; }
        public Nullable<decimal> jv_frt_gst { get; set; }
        public Nullable<decimal> jv_thc { get; set; }
        public Nullable<decimal> jv_thc_gst { get; set; }
        public Nullable<decimal> jv_detn { get; set; }
        public Nullable<decimal> jv_detn_gst { get; set; }
        public Nullable<decimal> jv_others { get; set; }
        public Nullable<decimal> jv_others_gst { get; set; }
        public Nullable<decimal> total_gst { get; set; }

        public string job_docno { get; set; }
        public string jexp_invoice_no { get; set; }
        public string jexp_comm_invoice_no { get; set; }


        public string job_sbno { get; set; }
        public string job_sbdt { get; set; }


        public Nullable<decimal> jv_dest_truck { get; set; }
        public Nullable<decimal> jv_bl_amend { get; set; }
        public Nullable<decimal> jv_bl_surr { get; set; }
        public Nullable<decimal> jv_bl_reissue { get; set; }
        public Nullable<decimal> jv_via_charge { get; set; }
        public Nullable<decimal> jv_detn2 { get; set; }

        public Nullable<decimal> hbl_chwt { get; set; }
        public Nullable<decimal> job_ntwt { get; set; }
        public Nullable<decimal> job_grwt { get; set; }

        public string branch { get; set; }

    }
}
