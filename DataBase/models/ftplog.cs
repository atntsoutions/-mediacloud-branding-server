using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class Ftplog : table_Base
    {
        public string ftp_pkid { get; set; }
        public string ftp_from { get; set; }
        public string ftp_to { get; set; }
        public string ftp_date { get; set; }
        public string ftp_action { get; set; }
        public string ftp_module { get; set; }
        public string ftp_module_pkid { get; set; }
        public string ftp_subject { get; set; }
        public string ftp_is_ack { get; set; }
        public string ftp_process_id { get; set; }
        public string ftp_comp_code { get; set; }
        public string ftp_branch_code { get; set; }
        public string ftp_user_code { get; set; }
        public string ftp_remarks { get; set; }
        public string ftp_isread { get; set; }
        public string ftp_file_path { get; set; }
    }
}
