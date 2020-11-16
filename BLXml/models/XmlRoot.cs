namespace BLXml.models
{
    public abstract class XmlRoot
    {
        public string ActualShipper = "Y";
        public string HBLNO ="";
       // public string HBL_IDS = "";
        public int Total_Records;
        public string MessageNumber = "";
        public string FileName = "VS";
        public string MODULE_ID = "";
        public abstract void Generate();
    }
}