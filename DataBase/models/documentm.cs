


namespace DataBase
{
    public class documentm : table_Base
    {
        public string doc_pkid { get; set; }
        public string doc_catg_name { get; set; }
        public string doc_file_name { get; set; }
        public string doc_full_name { get; set; }
        public string doc_group_id { get; set; }
        public string doc_catg_id { get; set; }
        public string doc_file_size { get; set; }
        public string rec_deleted_by { get; set; }
        public bool row_displayed { get; set; }
    }
}