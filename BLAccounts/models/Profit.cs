using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class Profit : table_Base
    {
        public string rowtype { get; set; }
        public string rowcolor { get; set; }


        public string type { get; set; }
        public string reportdate { get; set; }

        public string mbl_date { get; set; }
        public string hbl_date { get; set; }
        public string mbl_no { get; set; }
        public string mbl_bl_no { get; set; }

        public string hbl_bl_no { get; set; }

        public string buy_date { get; set; }
        public string sell_date { get; set; }
        public string exporter { get; set; }
        public string consignee { get; set; }

        public string agent { get; set; }

        public string sman { get; set; }
        public string nomination { get; set; }
        public string mbl_terms { get; set; }
        public string hbl_terms { get; set; }


        public Nullable<decimal> mbl_chwt { get; set; }
        public Nullable<decimal> hbl_chwt { get; set; }
        public Nullable<decimal> mbl_grwt { get; set; }
        public Nullable<decimal> hbl_grwt { get; set; }
        public Nullable<decimal> hbl_cbm { get; set; }
        public Nullable<decimal> hbl_ntwt { get; set; }

        public Nullable<decimal> frt_dr { get; set; }
        public Nullable<decimal> fsc_dr { get; set; }
        public Nullable<decimal> wrs_dr { get; set; }
        public Nullable<decimal> mcc_dr { get; set; }
        public Nullable<decimal> oth_dr { get; set; }
        public Nullable<decimal> frt_cr { get; set; }
        public Nullable<decimal> fsc_cr { get; set; }
        public Nullable<decimal> wrs_cr { get; set; }
        public Nullable<decimal> mcc_cr { get; set; }
        public Nullable<decimal> oth_cr { get; set; }


        public Nullable<decimal> ex_01 { get; set; }
        public Nullable<decimal> ex_02 { get; set; }
        public Nullable<decimal> ex_03 { get; set; }
        public Nullable<decimal> ex_04 { get; set; }
        public Nullable<decimal> ex_05 { get; set; }
        public Nullable<decimal> ex_06 { get; set; }
        public Nullable<decimal> ex_07 { get; set; }
        public Nullable<decimal> ex_10 { get; set; }
        public Nullable<decimal> ex_17 { get; set; }
        public Nullable<decimal> ex_ot { get; set; }

        public Nullable<decimal> in_01 { get; set; }
        public Nullable<decimal> in_02 { get; set; }
        public Nullable<decimal> in_03 { get; set; }
        public Nullable<decimal> in_04 { get; set; }
        public Nullable<decimal> in_05 { get; set; }
        public Nullable<decimal> in_06 { get; set; }
        public Nullable<decimal> in_07 { get; set; }
        public Nullable<decimal> in_10 { get; set; }
        public Nullable<decimal> in_17 { get; set; }
        public Nullable<decimal> in_ot { get; set; }




        public Nullable<decimal> margin_cr { get; set; }
        public Nullable<decimal> frt_ho_dr { get; set; }
        public Nullable<decimal> frt_ho_cr { get; set; }
        public Nullable<decimal> buy { get; set; }
        public Nullable<decimal> sell { get; set; }




        public Nullable<decimal> in_1401 { get; set; }
        public Nullable<decimal> in_1402 { get; set; }

        public Nullable<decimal> in_1403 { get; set; }
        public Nullable<decimal> in_1404 { get; set; }
        public Nullable<decimal> in_1405 { get; set; }
        public Nullable<decimal> ex_1401 { get; set; }
        public Nullable<decimal> ex_1402 { get; set; }
        public Nullable<decimal> ex_1403 { get; set; }
        public Nullable<decimal> ex_1404 { get; set; }
        public Nullable<decimal> ex_1405 { get; set; }
        public Nullable<decimal> cost_dr { get; set; }
        public Nullable<decimal> cost_cr { get; set; }
        public Nullable<decimal> income { get; set; }
        public Nullable<decimal> expense { get; set; }

        public Nullable<decimal> profit { get; set; }
        public Nullable<decimal> total { get; set; }
        public Nullable<decimal> roi { get; set; }



        public string mawb_date { get; set; }
        public string hawb_date { get; set; }
        public string jvh_year { get; set; }
        public string hbl_no { get; set; }
        public string hbl_rec_creared_date { get; set; }
        public string liner { get; set; }
        public string mawb_no { get; set; }
        public string hawb_no { get; set; }
        public string discription { get; set; }
        public string consignee_city { get; set; }
        public string consignee_state { get; set; }
        

        public Nullable<decimal> ex_1105 { get; set; }
        public Nullable<decimal> ex_1106 { get; set; }
        public Nullable<decimal> ex_1107 { get; set; }
        public Nullable<decimal> in_1105 { get; set; }
        public Nullable<decimal> in_1106 { get; set; }
        public Nullable<decimal> in_1107 { get; set; }

        public Nullable<decimal> ex_1301 { get; set; }
        public Nullable<decimal> ex_1302 { get; set; }
        public Nullable<decimal> ex_1303 { get; set; }
        public Nullable<decimal> ex_1304 { get; set; }
        public Nullable<decimal> ex_1305 { get; set; }
        public Nullable<decimal> ex_1306 { get; set; }
        public Nullable<decimal> ex_1307 { get; set; }
        public Nullable<decimal> in_1301 { get; set; }
        public Nullable<decimal> in_1302 { get; set; }
        public Nullable<decimal> in_1303 { get; set; }
        public Nullable<decimal> in_1304 { get; set; }
        public Nullable<decimal> in_1305 { get; set; }
        public Nullable<decimal> in_1306 { get; set; }
        public Nullable<decimal> in_1307 { get; set; }

        public string branch { get; set; }

        public string job_no { get; set; }
        public string job_date { get; set; }
        public string job_terms { get; set; }

        public Nullable<decimal> job_ntwt { get; set; }
        public Nullable<decimal> job_grwt { get; set; }
        public Nullable<decimal> job_cbm { get; set; }
        public Nullable<decimal> job_chwt { get; set; }

        public Nullable<decimal> ex_1101 { get; set; }
        public Nullable<decimal> ex_1102 { get; set; }
        public Nullable<decimal> ex_1103 { get; set; }
        public Nullable<decimal> in_1101 { get; set; }
        public Nullable<decimal> in_1102 { get; set; }
        public Nullable<decimal> in_1103 { get; set; }

        public Nullable<decimal> ex_1201 { get; set; }
        public Nullable<decimal> ex_1202 { get; set; }
        public Nullable<decimal> ex_1203 { get; set; }
        public Nullable<decimal> in_1201 { get; set; }
        public Nullable<decimal> in_1202 { get; set; }
        public Nullable<decimal> in_1203 { get; set; }

        public string pol { get; set; }
        public string pod { get; set; }
        public string pofd { get; set; }

        public string job_type { get; set; }
        public string bl_notify_name { get; set; }
        public string exp_city { get; set; }
        public string exp_state { get; set; }
        public string job_created_date { get; set; }
        public string job_commodity { get; set; }
        public string job_invoice_nos { get; set; }
        public string job_cntr_type { get; set; }

        public string mbl_folder_no { get; set; }
        public string hbl_created_date { get; set; }
        public string mbl_status { get; set; }
        public string hbl_ar_invnos { get; set; }
        public string hbl_orgin_country { get; set; }
        public string hbl_book_cntr { get; set; }
        public Nullable<decimal> hbl_book_cntr_teu { get; set; }
        public string hbl_nature { get; set; }
        public string mbl_jobtype { get; set; }


        public string mbl_pol_etd { get; set; }
        public string mbl_commodity { get; set; }
        public string mbl_liner { get; set; }
        public string hbl_ddp_ddu_exwork { get; set; }

        public string exp_created { get; set; }
        public string mbl_shipment_type { get; set; }

        public string org_country { get; set; }
        public string pod_country { get; set; }
        public string pofd_country { get; set; }
        public string buyer_name { get; set; }
        

        public Nullable<decimal> rebate_dr { get; set; }
    }
}
