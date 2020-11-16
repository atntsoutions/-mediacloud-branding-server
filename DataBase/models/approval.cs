using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{

    public class approvald : table_Base
    {
        public string ad_pkid { get; set; }
        public string ad_parent_id { get; set; }
        public string ad_by { get; set; }
        public string ad_remarks { get; set; }
        public string ad_status { get; set; }
        public string ad_date { get; set; }
    }

}