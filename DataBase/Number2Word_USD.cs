using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataBase
{
    public static class Number2Word_USD
    {
        public static string Convert(string Number, string Currency, string Dec_Name)
        {
            if (Number == "0" || Number == "" || Number == ".")
                return "Zero";
            try
            {
                string sData = Number;
                string[] str;
                if (sData.IndexOf('.') >= 0)
                {
                    str = sData.Split('.');
                    sData = Num2Words(str[0], "LEFT PORTION");

                    if (str[1].ToString() != "00")
                    {
                        string Str = Num2Words(str[1], "RIGHT PORTION");
                        if (Str != "")
                            sData += " And " + Dec_Name + Str;
                    }
                }
                else
                    sData = Num2Words(sData, "LEFT PORTION");
                sData += " Only ";
                if (Currency != "")
                    sData = Currency + " : " + sData;
                return sData;
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
        private static string Num2Words(string sData, string sType)
        {
            string Billion = "";
            string Hundred_Billion = "";

            string Hundred_Million = "";
            string Million = "";
            string Hundred_Thousand = "";
            string Thousand = "", Hundred = "", Tens = "";
            string Str = sData.PadLeft(12, ' ');

            //211234567891

            Hundred_Billion = Str.Substring(0, 1).Trim();
            Billion = Str.Substring(1, 2).Trim();

            Hundred_Million = Str.Substring(3, 1).Trim();
            Million = Str.Substring(4, 2).Trim();

            Hundred_Thousand = Str.Substring(6, 1).Trim();
            Thousand = Str.Substring(7, 2).Trim();
            Hundred = Str.Substring(9, 1).Trim();
            Tens = Str.Substring(10, 2).Trim();

            Str = "";
            Str += Convert(Hundred_Billion, "Hundred");
            Str += Convert(Billion, "Billion");
            if (Hundred_Billion != "" && Billion == "00")
                Str += " Billion ";

            Str += Convert(Hundred_Million, "Hundred");
            Str += Convert(Million, "Million");
            if (Hundred_Million != "" && Million == "00")
                Str += " Million ";

            Str += Convert(Hundred_Thousand, "Hundred");
            Str += Convert(Thousand, "Thousand");
            if (Hundred_Thousand != "" && Thousand == "00")
                Str += " Thousand ";
            Str += Convert(Hundred, "Hundred");
            Str += Convert(Tens, "");
            return Str;
        }
        private static string Convert(string sData, string Suffix)
        {
            string Str = "";
            if (sData == "")
                return "";
            int iPos = int.Parse(sData);
            if (iPos <= 0)
                return "";
            string[] sArray = new string[] {
                "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
                "Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen", "Seventeen","Eighteen","Nineteen",
                "Twenty", "Twenty One" , "Twenty Two" , "Twenty Three" , "Twenty Four" ,"Twenty Five" , "Twenty Six" ,"Twenty Seven","Twenty Eight","Twenty Nine" ,
                "Thirty", "Thirty One" , "Thirty Two" , "Thirty Three" , "Thirty Four" ,"Thirty Five" , "Thirty Six" ,"Thirty Seven","Thirty Eight","Thirty Nine" ,
                "Forty", "Forty One" , "Forty Two" , "Forty Three" , "Forty Four" ,"Forty Five" , "Forty Six","Forty Seven","Forty Eight","Forty Nine" ,
                "Fifty", "Fifty One" , "Fifty Two" , "Fifty Three" , "Fifty Four" ,"Fifty Five" , "Fifty Six","Fifty Seven","Fifty Eight","Fifty Nine" ,
                "Sixty", "Sixty One" , "Sixty Two" , "Sixty Three" , "Sixty Four" ,"Sixty Five" , "Sixty Six","Sixty Seven","Sixty Eight","Sixty Nine" ,
                "Seventy", "Seventy One" , "Seventy Two" , "Seventy Three" , "Seventy Four" ,"Seventy Five","Seventy Six","Seventy Seven","Seventy Eight" ,"Seventy Nine" ,
                "Eighty", "Eighty One" , "Eighty Two" , "Eighty Three" , "Eighty Four" ,"Eighty Five","Eighty Six","Eighty Seven","Eighty Eight" ,"Eighty Nine" ,
                "Ninety", "Ninety One" , "Ninety Two" , "Ninety Three" , "Ninety Four" ,"Ninety Five","Ninety Six","Ninety Seven","Ninety Eight" ,"Ninety Nine" ,
            };
            if (sData != "")
                Str = " " + sArray[iPos].ToString();
            return Str + " " + Suffix;
        }
    }
}
