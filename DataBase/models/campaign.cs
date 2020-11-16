using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{


    public class Campaign : table_Base
    {
        public string cam_pkid { get; set; }
        public int cam_slno { get; set; }
        public string cam_name { get; set; }
        public string cam_tab_id { get; set; }
        public string cam_tab_name { get; set; }

        public string cam_table_name { get; set; }

        public string user_is_admin { get; set; }

        public string cam_type { get; set; }
        public string cam_store { get; set; }
        public string cam_product_name { get; set; }
        public string cam_size { get; set; }
        public string cam_aep { get; set; }
        public string cam_output { get; set; }
        public string cam_approver { get; set; }
        public string cam_receiver { get; set; }
        public string cam_logo { get; set; }
        public string cam_image1 { get; set; }
        public string cam_image2 { get; set; }
        public string cam_image3 { get; set; }
        public string cam_image4 { get; set; }
        public string cam_image5 { get; set; }
        public string cam_text1 { get; set; }
        public string cam_text2 { get; set; }
        public string cam_text3 { get; set; }
        public string cam_text4 { get; set; }
        public string cam_text5 { get; set; }


        public Boolean tab_campaign_table { get; set; }

        public string cam_product_name_values { get; set; }

        public string cam_size_values { get; set; }
        public string cam_aep_values { get; set; }
        public string cam_output_values { get; set; }
        public string cam_text1_values { get; set; }
        public string cam_text2_values { get; set; }
        public string cam_text3_values { get; set; }
        public string cam_text4_values { get; set; }
        public string cam_text5_values { get; set; }


    }

}