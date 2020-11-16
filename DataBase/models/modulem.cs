using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{


    public class Modulem : table_Base
    {
        public string module_pkid { get; set; }
        public string module_name { get; set; }
        public int module_order { get; set; }

    }

}