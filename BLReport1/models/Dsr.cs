using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class DsrReport : table_Base
    {
        public string job_pkid { get; set; }
        public string job_date { get; set; }
        public string job_docno { get; set; }
        public string job_invoice_nos { get; set; }
        public string job_prefix { get; set; }
        public string job_shipper { get; set; }
        public string job_consignee { get; set; }
        public string job_pol { get; set; }
        public string job_pod { get; set; }
        public string hbl_bl_no { get; set; }
        public string mbl_bl_no { get; set; }
        public Nullable<decimal> job_cbm { get; set; }
        public Nullable<decimal> job_pkg { get; set; }
        public Nullable<decimal> job_pcs { get; set; }
        public Nullable<decimal> job_ntwt { get; set; }
        public Nullable<decimal> job_grwt { get; set; }
        public string opr_sbill_no{ get; set; }
        public string opr_sbill_date { get; set; }
        public string opr_cargo_received_on { get; set; }
        public string forwarder_name { get; set; }
        public string mbl_vessel_name { get; set; }
        public string mbl_vessel_no { get; set; }
        public string opr_stuffed_at { get; set; }
        public string opr_stuffed_on { get; set; }
        public string hbl_book_cntr { get; set; }
        public string mbl_pol_etd { get; set; }
        public string mbl_pofd_eta { get; set; }
        public string job_type { get; set; }


        public string job_cha_name { get; set; }
        public string job_agent_name { get; set; }
        public string job_pofd_name { get; set; }
        public string salesman { get; set; }
        public string opr_cleared_date { get; set; }
        public string job_terms { get; set; }
        public string job_nomination { get; set; }
        public Nullable<decimal> job_chwt { get; set; }
        public string hbl_date { get; set; }
        public string hbl_invoice_nos { get; set; }
        public string mbl_date { get; set; }
        public string liner_name { get; set; }

        public Nullable<decimal> mbl_grwt { get; set; }
        public Nullable<decimal> mbl_chwt { get; set; }
        public string mbl_folder_no { get; set; }
        public string mbl_folder_sent_date { get; set; }
        public string opr_drawback_slno { get; set; }
        public string opr_drawback_date { get; set; }
        public Nullable<decimal> opr_drawback_amt { get; set; }
        public string hbl_no { get; set; }
        public string mbl_no { get; set; }

        public string hbl_exporter_name  { get; set; }
        public string hbl_importer_name { get; set; }
        public string hbl_agent_name { get; set; }
        public string impj_edi_no { get; set; }
        public string mbl_pol { get; set; }

        public Nullable<decimal> hbl_pkg { get; set; }
        public Nullable<decimal> hbl_cbm { get; set; }
        public Nullable<decimal> hbl_grwt { get; set; }
        public Nullable<decimal> hbl_chwt { get; set; }
        public string cha_name { get; set; }
        public string hbl_pod_eta { get; set; }
        public string impj_be_type { get; set; }
        public string impj_docs_required { get; set; }
        public string impj_edichklst_sent_on { get; set; }
        public string hbl_beno { get; set; }
        public string hbl_bedate { get; set; }
        public string impj_status { get; set; }
        public string impj_status_date { get; set; }
        public string impj_cleared_on { get; set; }
        public string hbl_remarks { get; set; }
        public string impj_doc_recvd_date { get; set; }
        public string impj_doc_send_date { get; set; }
        public string impj_waybill_no { get; set; }
        public string impj_waybill_date { get; set; }
        public string impj_sbno { get; set; }
        public string impj_sbdate { get; set; }
        public string job_remarks { get; set; }
        public string job_commodity { get; set; }
      
        public string job_status { get; set; }


        public string job_sman { get; set; }
        public string mbl_pod { get; set; }
        public string branch { get; set; }

        public string opr_ep_rec_date { get; set; }

        public string mbl_book_no { get; set; }
        public string mbl_book_date { get; set; }
        public string mbl_prealert_date { get; set; }
        public string hbl_ar_invnos { get; set; }
        public string job_nature { get; set; }

        public string job_cntr_type { get; set; }
        public Nullable<decimal> job_cntr_teu { get; set; }
        public string job_billtype_id { get; set; }

        public string hbl_released_date { get; set; }
        public string mbl_released_date { get; set; }

        public Nullable<decimal> hbl_ar_invamt { get; set; }
        public Nullable<decimal> hbl_ar_gstamt { get; set; }

       
        public string job_liner_agent { get; set; }
        public string job_cntr { get; set; }

        public string mbl_vessel2_name { get; set; }

    }
    
}
