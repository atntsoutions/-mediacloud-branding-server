using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class Companym : table_Base
    {
        public string comp_pkid { get; set; }
        public string comp_type { get; set; }
        public string comp_parent_id { get; set; }
        public string comp_parent_name { get; set; }
        public string comp_code { get; set; }
        public string comp_name { get; set; }

        public string comp_short_name { get; set; }

        public string comp_address1 { get; set; }
        public string comp_address2 { get; set; }
        public string comp_address3 { get; set; }

        public string comp_district { get; set; }
        public string comp_state { get; set; }

        public string comp_region_id { get; set; }
        public string comp_region_name { get; set; }

        public string comp_tel { get; set; }
        public string comp_fax { get; set; }
        public string comp_web { get; set; }

        

        public string comp_email { get; set; }
        public string comp_ptc { get; set; }
        public string comp_mobile { get; set; }
        public string comp_prefix { get; set; }
        public string comp_panno { get; set; }
        public string comp_cinno { get; set; }
        public string comp_gstin { get; set; }
        public string comp_reg_address { get; set; }
        public string comp_iata_code { get; set; }
        public string comp_location { get; set; }


        public string comp_approver_email { get; set; }
        public string comp_receiver_email { get; set; }

        public string comp_logo_name { get; set; }
        public Boolean comp_logo_uploaded { get; set; }

        public string comp_image_name { get; set; }
        public Boolean comp_image_uploaded { get; set; }

        public string comp_branch_type { get; set; }

        public string comp_country_code { get; set; }
        public string comp_pol_code { get; set; }

        public int comp_order { get; set; }
        public string comp_uamno { get; set; }


        public string pkid { get; set; }
        public string user_id { get; set; }
        public Boolean selected { get; set; }

    }

}