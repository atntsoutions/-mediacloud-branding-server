using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class Settings : table_Base
    {
        public string parentid { get; set; }
        public string tablename { get; set; }
        public string caption { get; set; }
        public string id { get; set; }
        public string code { get; set; }
        public string name { get; set; }
        public string tabletype { get; set; }
    }

    public class Settings_VM : table_Base
    {
        public List<Settings> RecordDet { get; set; }
    }

    public class Lockingm : table_Base
    {
        public string lock_pkid { get; set; }
        public int lock_year { get; set; }
        public string lock_ar { get; set; }
        public string lock_ap { get; set; }
        public string lock_drn { get; set; }
        public string lock_crn { get; set; }
        public string lock_dri { get; set; }
        public string lock_cri { get; set; }
        public string lock_cr { get; set; }
        public string lock_cp { get; set; }
        public string lock_br { get; set; }
        public string lock_bp { get; set; }
        public string lock_jv { get; set; }
        public string lock_cjv { get; set; }

    }
}