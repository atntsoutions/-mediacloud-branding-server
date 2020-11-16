using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{

    public class Param : table_Base
    {
        public string param_pkid { get; set; }
        public string param_type { get; set; }
        public string param_code { get; set; }
        public string param_name { get; set; }
        public string param_id1 { get; set; }
        public string param_id2 { get; set; }
        public string param_id3 { get; set; }
        public string param_id4 { get; set; }
        public string param_email { get; set; }
        public decimal param_rate { get; set; }

        public int param_slno { get; set; }

        public string param_file_name { get; set; }
        
        public bool param_file_uploaded { get; set; }

        public string param_server_folder { get; set; }
    }

    
    public class paramvalues : table_Base
    {
        public string param_pkid { get; set; }
        public string parent_id { get; set; }
        public string param_key { get; set; }
        public string param_value { get; set; }
        public string param_defvalue { get; set; }
        public string param_filetype { get; set; }
        public string param_impexp { get; set; }
        public string param_format { get; set; }
        public string param_edifile { get; set; }
    }

    public class paramvalues_vm : table_Base
    {
        public string param_pkid { get; set; }
        public List<paramvalues> RecordDet { get; set; }
    }

}