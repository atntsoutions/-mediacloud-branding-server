using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class InvoiceReport : table_Base
    {
        public string inv_comp_name { get; set; }
        public string inv_comp_add1 { get; set; }
        public string inv_comp_add2 { get; set; }
        public string inv_comp_add3 { get; set; }
        public string inv_comp_tel { get; set; }
        public string inv_comp_fax { get; set; }
        public string inv_comp_web { get; set; }
        public string inv_comp_email { get; set; }
        public string inv_comp_panno { get; set; }
        public string inv_comp_cinno { get; set; }
        public string inv_comp_gstin { get; set; }
        public string inv_comp_reg_address { get; set; }
        public string inv_comp_uamno { get; set; }

        public string jvh_docno { get; set; }
        public string jvh_date { get; set; }
        public string jvh_gstin { get; set; }
        public string jvh_state_code { get; set; }
        public string jvh_state_name { get; set; }
        public string jvh_cc_category { get; set; }
        public string jvh_party_name { get; set; }
        public string jvh_party_addr1 { get; set; }
        public string jvh_party_addr2 { get; set; }
        public string jvh_party_addr3 { get; set; }
        public Nullable<decimal> jvh_tot_amt { get; set; }
        public Nullable<decimal> jvh_cgst_amt { get; set; }
        public Nullable<decimal> jvh_sgst_amt { get; set; }
        public Nullable<decimal> jvh_igst_amt { get; set; }
        public Nullable<decimal> jvh_net_amt { get; set; }
        public string jvh_curr_code { get; set; }
        public Nullable<decimal> jvh_tot_famt { get; set; }
        public Nullable<decimal> jvh_cgst_famt { get; set; }
        public Nullable<decimal> jvh_sgst_famt { get; set; }
        public Nullable<decimal> jvh_igst_famt { get; set; }
        public Nullable<decimal> jvh_net_famt { get; set; }
        public string jvh_narration { get; set; }
        public string jvh_type { get; set; }
        public string jvh_reference { get; set; }
        public string jvh_acc_id { get; set; }
        public string jvh_org_invno { get; set; }
        

        public string hbl_exp_id { get; set; }
        public string hbl_exp_name { get; set; }
        public string hbl_imp_id { get; set; }
        public string hbl_bl_no { get; set; }
        public string hbl_no { get; set; }
        public string hbl_date { get; set; }
        public string hbl_pol_name { get; set; }
        public string hbl_pod_name { get; set; }
        public string hbl_pofd_name { get; set; }
        public string hbl_freight_status { get; set; }
        public Nullable<decimal> hbl_pkg { get; set; }
        public string hbl_pkg_unit { get; set; }
        public Nullable<decimal> hbl_cbm { get; set; }
        public string hbl_consignee_name { get; set; }
        public string hbl_containers { get; set; }
        public string hbl_beno { get; set; }
        public string hbl_bedate { get; set; }
        public Nullable<decimal> hbl_ntwt { get; set; }
        public Nullable<decimal> hbl_grwt { get; set; }
        public Nullable<decimal> hbl_chwt { get; set; }
        public string hbl_wt_unit { get; set; }
        public string hbl_genjob_no { get; set; }
        public string hbl_genigm_no { get; set; }
        public string hbl_carton_nos { get; set; }
        public string hbl_invoice_nos { get; set; }
        public string hbl_genjobtype_code { get; set; }
        public int hbl_invnos_prncount { get; set; }

        public string mbl_bkno { get; set; }
        public string mbl_no { get; set; }
        public string mbl_date { get; set; }
        public string mbl_vessel_name { get; set; }
        public string mbl_vessel_voyage { get; set; }
        public string mbl_pol_etd { get; set; }
        public string mbl_cha_name { get; set; }
        public string mbl_igmno { get; set; }
        public string mbl_igmdate { get; set; }

        public int jv_ctr { get; set; }
        public string jv_sac_code { get; set; }
        public string jv_acc_name { get; set; }
        public Nullable<decimal> jv_total { get; set; }
        public Nullable<decimal> jv_cgst_rate { get; set; }
        public Nullable<decimal> jv_cgst_amt { get; set; }
        public Nullable<decimal> jv_sgst_rate { get; set; }
        public Nullable<decimal> jv_sgst_amt { get; set; }
        public Nullable<decimal> jv_igst_rate { get; set; }
        public Nullable<decimal> jv_igst_amt { get; set; }
        public Nullable<decimal> jv_net_total { get; set; }
        public Nullable<decimal> jv_qty { get; set; }
        public Nullable<decimal> jv_rate { get; set; }
        public Nullable<decimal> jv_ftotal { get; set; }
        public Nullable<decimal> jv_exrate { get; set; }
        public string jv_curr_code { get; set; }
        public Nullable<decimal> jv_total_fc { get; set; }
        public Nullable<decimal> jv_cgst_famt { get; set; }
        public Nullable<decimal> jv_sgst_famt { get; set; }
        public Nullable<decimal> jv_igst_famt { get; set; }
        public Nullable<decimal> jv_net_ftotal { get; set; }


        public string job_sbnos { get; set; }
        public string job_invnos { get; set; }
        public string job_commodity { get; set; }
        public string job_comm_invnos { get; set; }

        public string gj_seal_no { get; set; }
        public string gj_loaded_on { get; set; }
        public string gj_unloaded_on { get; set; }
        public string gj_cfs { get; set; }
        public string gj_from { get; set; }
        public string gj_to1 { get; set; }
        public string gj_to2 { get; set; }
        public string gj_shipper_inv_no { get; set; }

        public string gj_vessel { get; set; }
        public string gj_mbl_no { get; set; }
        public string gj_hbl_no { get; set; }
        public string gj_frt_status { get; set; }
        public string gj_cha_name { get; set; }
        public string gj_consignee_name { get; set; }
        public string gj_sb_no { get; set; }
        public string gj_lr_no { get; set; }
        public string gj_pack_list_no { get; set; }

    }
}
