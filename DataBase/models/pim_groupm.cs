using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class pim_groupm : table_Base
    {
        public string grp_pkid { get; set; }
        public string grp_parent_id { get; set; }
        public int grp_level { get; set; }

        public string grp_name { get; set; }
        public int grp_level_slno { get; set; }
        public string grp_level_id { get; set; }
        public string grp_level_name { get; set; }
        public string grp_table_name { get; set; }


        public string rec_type { get; set; }

    }

}