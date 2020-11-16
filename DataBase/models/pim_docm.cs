using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class pim_docm : table_Base
    {
        public string doc_pkid { get; set; }


        public string doc_store_id { get; set; }
        public string doc_store_name { get; set; }

        public string doc_grp_id { get; set; }
        public string doc_grp_name { get; set; }
        public string doc_grp_level_name { get; set; }
        public string doc_table_name { get; set; }
        public int doc_slno { get; set; }
        public string doc_name { get; set; }
        public string doc_file_name { get; set; }
        public Boolean doc_file_uploaded { get; set; }
        public string doc_thumbnail { get; set; }

        public string doc_server_folder { get; set; }
    }
    
}