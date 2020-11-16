using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    
    public class tablesm : table_Base
    {
        public string tab_pkid { get; set; }
        public string tab_name { get; set; }
        public string tab_table_name { get; set; }
        public string tab_caption { get; set; }

        public string tab_id { get; set; }
        public string tab_store { get; set; }
        public string tab_group { get; set; }
        public string tab_sku { get; set; }
        public string tab_file { get; set; }

        public Boolean tab_sku_duplication { get; set; }
        public Boolean tab_store_duplication { get; set; }
        public Boolean tab_campaign_table { get; set; }
    }

    public class tablesd : table_Base
    {
        public string tabd_pkid { get; set; }
        public string tabd_parent_id { get; set; }

        public string tabd_thumbnail { get; set; }

        public string tabd_tab_name { get; set; }
        public string tabd_table_name { get; set; }
        public string tabd_col_name { get; set; }
        public string tabd_col_caption { get; set; }
        public string tabd_col_type { get; set; }
        public int tabd_col_rows { get; set; }
        public int tabd_col_len { get; set; }
        public int tabd_col_dec { get; set; }
        public string tabd_col_case { get; set; }
        public string tabd_col_mandatory { get; set; }
        public string tabd_col_id { get; set; }
        public string tabd_col_value { get; set; }
        public string tabd_col_list { get; set; }
        public int tabd_col_order { get; set; }


        public Boolean tabd_col_file_uploaded { get; set; }


    }

}