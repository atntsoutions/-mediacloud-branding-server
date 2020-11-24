using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{

    public class pim_spot : table_Base
    {
        public string spot_pkid { get; set; }
        public string spot_date { get; set; }
        public int spot_slno { get; set; }

        public string spot_store_id { get; set; }
        public string spot_store_name { get; set; }

        public string spot_vendor_id { get; set; }
        public string spot_vendor_name { get; set; }

        public string spot_region_id { get; set; }
        public string spot_region_name { get; set; }

        public string spot_recce_id { get; set; }
        public string spot_recce_name { get; set; }


        public string spot_executive_name { get; set; }
        public string spot_store_contact_name { get; set; }
        public string spot_store_contact_tel { get; set; }

        public string spot_job_remarks { get; set; }

        public string spot_server_folder { get; set; }

        public string spot_store_view { get; set; }

        public bool spot_store_view_file_uploaded { get; set; }

        public string spot_installation_view { get; set; }
        public bool spot_installation_view_file_uploaded { get; set; }

    }


    public class pim_spotd : table_Base
    {
        public string spotd_pkid { get; set; }
        public string spotd_parent_id { get; set; }
        public int spotd_slno { get; set; }
        public string spotd_name { get; set; }

        public string spotd_uom { get; set; }
        public decimal spotd_wd { get; set; }
        public decimal spotd_ht { get; set; }

        public string spotd_artwork_id { get; set; }
        public string spotd_artwork_name { get; set; }
        public string spotd_artwork_file_name { get; set; }
        public string spotd_artwork_folder { get; set; }

        public string spotd_product_id { get; set; }
        public string spotd_product_name { get; set; }

        public string spotd_artwork_server_folder { get; set; }

        public string spotd_server_folder { get; set; }
        public string spotd_close_view { get; set; }
        public bool spotd_close_view_file_uploaded { get; set; }

        public string spotd_long_view { get; set; }
        public bool spotd_long_view_file_uploaded { get; set; }

        public string spotd_final_view { get; set; }
        public bool spotd_final_view_file_uploaded { get; set; }


    }


}