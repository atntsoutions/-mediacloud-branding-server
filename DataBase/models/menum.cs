using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;
namespace DataBase
{

    public class Menum : table_Base
    {
        public string menu_pkid { get; set; }
        public string menu_code { get; set; }
        public string menu_name { get; set; }
        public string menu_route1 { get; set; }
        public string menu_route2 { get; set; }
        public string menu_type { get; set; }
        public Boolean menu_sep { get; set; }
        public int menu_order { get; set; }

        public string menu_module_id { get; set; }
        public string menu_module_name { get; set; }
        public int menu_module_order { get; set; }


        public Boolean menu_displayed { get; set; }

        public Boolean rights_company { get; set; }
        public Boolean rights_admin { get; set; }
        public Boolean rights_restricted { get; set; }
        public Boolean rights_add { get; set; }
        public Boolean rights_edit { get; set; }
        public Boolean rights_delete { get; set; }
        public Boolean rights_print { get; set; }
        public Boolean rights_email { get; set; }
        public Boolean rights_docs { get; set; }
        public Boolean rights_docs_upload { get; set; }

        public Boolean rights_view { get; set; }
        public string rights_approval { get; set; }
    }

}