using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{

    public class mailhistory : table_Base
    {
        public string mail_pkid { get; set; }
        public string mail_date { get; set; }
        public string mail_source { get; set; }
        public string mail_source_id { get; set; }
        public string mail_send_by { get; set; }
        public string mail_send_to { get; set; }
        public string mail_send_cc { get; set; }
        public string mail_refno { get; set; }
        public string mail_comments { get; set; }
        public string mail_files { get; set; }

        public string mail_subject { get; set; }
        public string mail_message { get; set; }

    }

}

