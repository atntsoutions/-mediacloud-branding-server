using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataBase
{
    public static class Number2Word_RS
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
                    sData = Num2Words(str[0]);
                    if (Lib.Convert2Decimal(str[1]) != 0)
                    {
                        sData += " And Paise " + Num2Words(str[1]);
                        sData = sData.Replace("Paise", Dec_Name);
                    }
                }
                else
                    sData = Num2Words(sData);
                sData += " Only";
                if (Currency != "")
                    sData = Currency + " : " + sData;
                return sData;
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
        private static string Num2Words(string sData)
        {
            string Crore_Thousand = "", Crore_Hundred = "", Crore = "", Lakhs = "";
            string Thousand = "", Hundred = "", Tens = "";
            string Str = sData.PadLeft(12, ' ');
            Crore_Thousand = Str.Substring(0, 2).Trim();
            Crore_Hundred = Str.Substring(2, 1).Trim();
            Crore = Str.Substring(3, 2).Trim();
            Lakhs = Str.Substring(5, 2).Trim();
            Thousand = Str.Substring(7, 2).Trim();
            Hundred = Str.Substring(9, 1).Trim();
            Tens = Str.Substring(10, 2).Trim();
            Str = "";
            Str += Convert(Crore_Thousand, "Thousand");
            Str += Convert(Crore_Hundred, "Hundred");
            Str += Convert(Crore, "");
            if (Str != "")
                Str += " Crore ";
            Str += Convert(Lakhs, "Lakhs");
            Str += Convert(Thousand, "Thousand");
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
