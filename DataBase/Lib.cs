using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Security.Cryptography;
using System.Drawing;
using XL.XSheet;
using DataBase.Connections;



namespace DataBase
{

    public static class Lib
    {
        private static decimal TotEarn = 0;
        private static decimal TotDeduct = 0;
        private static decimal TotLopAmt = 0;
        //Testing Editing

        public static string BACK_END_DATE_FORMAT = "dd-MMM-yyyy";
        public static string FRONT_END_DATE_FORMAT = "yyyy-MM-dd";
        public static string FRONT_END_DATE_DISPLAY_FORMAT = "dd/MM/yyyy";

        public static string REPORT_FOLDER = @"c:\report";



        public static void CreateErrorLog(string str, string fName = "")
        {
            //string FileName = FolderName +  "errorlog.txt";

            string FileName = @"c:\\log\errorlog2.txt";
            if (fName != "")
                FileName = fName;

            StreamWriter sw = new StreamWriter(FileName, true);
            sw.WriteLine(str);
            sw.Flush();
            sw.Close();
        }


        public static string DatetoString(object mDate)
        {
            string str = "";
            if (mDate == DBNull.Value)
                str = "";
            else
            {
                str = ((DateTime)mDate).ToString(Lib.FRONT_END_DATE_FORMAT);
            }
            return str;
        }

        public static string DatetoStringDisplayformat(object mDate)
        {
            string str = "";
            if (mDate == DBNull.Value)
                str = "";
            else
            {
                str = ((DateTime)mDate).ToString(Lib.FRONT_END_DATE_DISPLAY_FORMAT);
            }
            return str;
        }


        public static string DatetoStringDisplayformatWithTime(object mDate)
        {
            string str = "";
            if (mDate == DBNull.Value)
                str = "";
            else
            {
                str = ((DateTime)mDate).ToString(Lib.FRONT_END_DATE_DISPLAY_FORMAT + " hh:mm:ss");
            }
            return str;
        }


        public static object ExcelCompatibleDate(object mDate, string format)
        {
            DateTime Dt;
            string str = "";
            if (mDate == DBNull.Value)
                return "";
            else
            {
                str = ((DateTime)mDate).ToString(format);
                Dt = DateTime.ParseExact(str.ToString(), format, System.Globalization.CultureInfo.InvariantCulture);
                return Dt;
            }
        }


        public static string GetFormNumber(string sBrCode, string sCategory)
        {
            string sFrmNum = "";
            if (sCategory == "WAGES")
            {
                if (sBrCode == "HOCPL" || sBrCode == "COKSF" ||
                    sBrCode == "COKAF" || sBrCode == "SEZSF")
                    sFrmNum = "FORM No. XI";
            }
            else if (sCategory == "PAYSLIP")
            {
                if (sBrCode == "HOCPL" || sBrCode == "COKSF" ||
                    sBrCode == "COKAF" || sBrCode == "SEZSF")
                    sFrmNum = "FORM NO. XIII";
            }
            return sFrmNum;
        }


        public static string StringToDate(Object Data)
        {
            string sData = "";
            DateTime Dt;
            if (Data == null || Data.ToString() == "")
                sData = "NULL";
            else
            {
                Dt = DateTime.Parse(Data.ToString());
                sData = Dt.ToString(Lib.BACK_END_DATE_FORMAT);
            }
            return sData;
        }
        public static string StringToDate(Object Data, string DateFormat)
        {
            string sData = "";
            DateTime Dt;
            if (Data == null || Data.ToString() == "")
                sData = "NULL";
            else
            {
                int dd = 0, mm = 0, yy = 0;
                string[] ThisDate = null;

                if (Data.ToString().Contains("/"))
                    ThisDate = Data.ToString().Split('/');
                else if (Data.ToString().Contains("-"))
                    ThisDate = Data.ToString().Split('-');
                else if (Data.ToString().Contains("."))
                    ThisDate = Data.ToString().Split('.');
                if (ThisDate != null)
                {
                    if (ThisDate.Length == 3)
                    {
                        if (DateFormat == "DD-MM-YYYY")
                        {
                            dd = Lib.Conv2Integer(ThisDate[0]);
                            mm = Lib.Conv2Integer(ThisDate[1]);
                            yy = Lib.Conv2Integer(ThisDate[2]);
                        }
                        else if (DateFormat == "MM-DD-YYYY")
                        {
                            mm = Lib.Conv2Integer(ThisDate[0]);
                            dd = Lib.Conv2Integer(ThisDate[1]);
                            yy = Lib.Conv2Integer(ThisDate[2]);
                        }
                        else if (DateFormat == "YYYY-MM-DD")
                        {
                            yy = Lib.Conv2Integer(ThisDate[0]);
                            mm = Lib.Conv2Integer(ThisDate[1]);
                            dd = Lib.Conv2Integer(ThisDate[2]);
                        }
                    }
                }
                if (mm > 0 && dd > 0)
                {
                    if (yy < 100)
                        yy = yy + 2000;
                    Dt = new DateTime(yy, mm, dd);
                    sData = Dt.ToString(Lib.BACK_END_DATE_FORMAT);
                }
                else
                    sData = "NULL";
            }
            return sData;
        }
        public static bool IsInFinYear(Object sDate, Object YearStartDate, Object YearEndDate, Boolean ChkFutureDate = false)
        {
            bool bOk = false;
            try
            {
                DateTime Date1 = DateTime.Parse(sDate.ToString());
                DateTime StartDate = DateTime.Parse(YearStartDate.ToString());
                DateTime EndDate = DateTime.Parse(YearEndDate.ToString());

                if (Date1 >= StartDate && Date1 <= EndDate)
                {
                    bOk = true;
                }
                if (bOk && ChkFutureDate)
                {
                    DateTime SERVERDATE = DateTime.Now;
                    int SERV_DATE_YY = SERVERDATE.Year;
                    int SERV_DATE_MM = SERVERDATE.Month;
                    int SERV_DATE_DD = SERVERDATE.Day;
                    SERVERDATE = new DateTime(SERV_DATE_YY, SERV_DATE_MM, SERV_DATE_DD);
                    if (Date1 > SERVERDATE)
                    {
                        bOk = false;
                    }
                }
            }
            catch (Exception)
            {
                bOk = false;
            }
            return bOk;
        }
        public static bool IsBeforeFinYear(Object sDate, Object YearStartDate)
        {
            bool bOk = false;
            try
            {
                DateTime StartDate = DateTime.Parse(YearStartDate.ToString());
                DateTime Date1 = DateTime.Parse(sDate.ToString());
                if (Date1 < StartDate)
                {
                    bOk = true;
                }
            }
            catch (Exception)
            {
                bOk = false;
            }
            return bOk;
        }

        public static bool IsFutureDate(Object sDate)
        {
            bool bOk = false;
            try
            {
                DateTime Date1 = DateTime.Parse(sDate.ToString());
                DateTime SERVERDATE = DateTime.Now;
                int SERV_DATE_YY = SERVERDATE.Year;
                int SERV_DATE_MM = SERVERDATE.Month;
                int SERV_DATE_DD = SERVERDATE.Day;
                SERVERDATE = new DateTime(SERV_DATE_YY, SERV_DATE_MM, SERV_DATE_DD);
                if (Date1 > SERVERDATE)
                {
                    bOk = true;
                }
            }
            catch (Exception)
            {
                bOk = false;
            }
            return bOk;
        }
        // Please make sure Conv2Decimal and Conver2Decimal is same
        public static decimal Conv2Decimal(string sText)
        {
            decimal nData;
            try
            {
                nData = decimal.Parse(sText);
            }
            catch (Exception)
            {
                nData = 0;
            }
            return nData;
        }
        // Please make sure Conv2Decimal and Conver2Decimal is same
        public static decimal Convert2Decimal(string sText)
        {
            decimal nData;
            try
            {
                nData = decimal.Parse(sText);
            }
            catch (Exception)
            {
                nData = 0;
            }
            return nData;
        }

        public static string[] ConvertString2Lines(string Text1, int width, string SplitType = "CHAR", string sfontName = "Calibri", float sfontSize = 10, string sfontStyle = "R")
        {
            string[] wLines = null;
            try
            {
                string Sentence = "";
                int itot = 0;
                if (SplitType == "CHAR")
                {
                    for (int i = 0; i < Text1.Length; i++)
                    {
                        itot++;
                        Sentence += Text1[i].ToString();
                        if ((itot % width) == 0)
                            Sentence += "\n";
                    }
                }
                if (SplitType == "WORD")
                {
                    FontStyle fStyle = FontStyle.Regular;
                    if (sfontStyle.Contains("B"))
                        fStyle = FontStyle.Bold;
                    if (sfontStyle.Contains("I"))
                        fStyle = FontStyle.Italic;

                    string[] WrdArry = Text1.Split(' ');
                    Sentence = "";
                    float LineWidth = 0;
                    for (int i = 0; i < WrdArry.Length; i++)
                    {
                        LineWidth += GetWordWidth(WrdArry[i], sfontName, sfontSize, fStyle);
                        if (LineWidth > width)
                        {
                            LineWidth = GetWordWidth(WrdArry[i], sfontName, sfontSize, fStyle);
                            Sentence += "\n";
                        }
                        if (Sentence != "")
                            Sentence += " ";
                        Sentence += WrdArry[i].ToString();
                    }
                }

                wLines = Sentence.Split('\n');
            }
            catch (Exception)
            {
                wLines = null;
            }
            return wLines;
        }
        public static string NumericFormat(String sNum, int DPlaces)
        {
            string sData;
            Decimal nAmt;
            string sFmt;
            try
            {
                sFmt = "{0:F" + DPlaces + "}";
                nAmt = Decimal.Parse(sNum);
                sData = String.Format(sFmt, nAmt);
            }
            catch (Exception)
            {
                sData = "0";
            }
            return sData;
        }



        public static Nullable<decimal> Conv2Decimal(string sText, string RETVAL)
        {
            decimal nData;
            try
            {
                nData = decimal.Parse(sText);
            }
            catch (Exception)
            {
                nData = 0;
            }
            if (nData == 0)
            {
                if (RETVAL == "NULL")
                    return null;
                else
                    return 0;
            }
            else
                return nData;
        }



        public static int Conv2Integer(string sText)
        {
            int nData;
            try
            {
                nData = int.Parse(sText);
            }
            catch (Exception)
            {
                nData = 0;
            }
            return nData;
        }


        public static Boolean AddError(ref string ErrorTarget, string ErrorMsg)
        {
            if (ErrorTarget.Length > 0 && ErrorMsg.Length > 0)
                ErrorTarget += "\n" + ErrorMsg;
            else
                ErrorTarget += ErrorMsg;

            return true;
        }
        public static string NumFormat(String sNum, int DPlaces, Boolean ThousandSep = false)
        {
            string sData;
            Decimal nAmt;
            string sFmt = "";
            string DecFmt = "";
            try
            {
                nAmt = Decimal.Parse(sNum);
                for (int i = 1; i <= DPlaces; i++)
                    DecFmt += "0";

                if (DecFmt != "")
                    sFmt = "0." + DecFmt;

                if (ThousandSep)
                {
                    if (sFmt != "")
                        sFmt = "0," + sFmt;
                    else
                        sFmt = "0,0";
                }

                if (sFmt != "")
                {
                    sFmt = "{0:" + sFmt + "}";
                    sData = String.Format(sFmt, nAmt);
                }
                else
                    sData = nAmt.ToString();
            }
            catch (Exception)
            {
                sData = "0";
            }
            return sData;
        }

        public static string NumFormat_Old(String sNum, int DPlaces, Boolean ThousandSep = false)
        {
            string sData;
            Decimal nAmt;
            string sFmt;
            try
            {
                nAmt = Decimal.Parse(sNum);


                if (DPlaces != 0)
                    sFmt = "{0:Z" + DPlaces + "}";
                else
                    sFmt = "{0:0}";

                if (ThousandSep)
                    sFmt = sFmt.Replace("Z", "N");

                sData = String.Format(sFmt, nAmt);
            }
            catch (Exception)
            {
                sData = "0";
            }
            return sData;
        }



        public static string SpellNumber(string MyNumber, string Currency, string Fraction)
        {
            return SpellNumber(MyNumber, Currency, Fraction, "");
        }

        public static string SpellNumber(string MyNumber, string Currency, string Fraction, string DecName)
        {
            string Dollars = "", Cents = "", Temp = "";
            int DecimalPlace, Count;
            string str;
            string[] Place = { "", "", "", "", "", "", "", "", "", "" };

            Place[2] = " Thousand ";
            Place[3] = " Million ";
            Place[4] = " Billion ";
            Place[5] = " Trillion ";

            //String representation of amount.
            MyNumber = MyNumber.Trim();
            //Position of decimal place 0 if none.
            DecimalPlace = MyNumber.IndexOf('.');
            //Convert cents and set MyNumber to dollar amount.
            if (DecimalPlace > 0)
            {
                str = MyNumber.Substring(DecimalPlace + 1) + "00";
                Cents = GetTens(str.Substring(0, 2));
                MyNumber = MyNumber.Substring(0, DecimalPlace);
            }
            Count = 1;
            while (MyNumber != "")
            {
                str = "000" + MyNumber;
                Temp = GetHundreds(str.Substring(str.Length - 3, 3));
                if (Temp != "")
                    Dollars = Temp + Place[Count] + Dollars;
                if (MyNumber.Length > 3)
                    MyNumber = MyNumber.Substring(0, MyNumber.Length - 3);
                else
                    MyNumber = "";
                Count = Count + 1;
            }

            switch (Dollars)
            {
                case "": Dollars = ""; break;
                case "One": Dollars = Currency + ((Currency != "") ? "One " : ""); break;
                default: Dollars = Currency + " " + Dollars; break;
            }
            switch (Cents)
            {
                /*
                case "":Cents = " " ;break;
                case "One":Cents = ((Fraction !="") ? " and One " : "") + Fraction  ;break;
                default:Cents = ((Fraction !="") ? " and " : "")  + Cents + " " + Fraction  ;break;
                */
                case "": Cents = ""; break;
                case "One": Cents = ((Fraction != "") ? " and One " : ""); break;
                default: Cents = ((Fraction != "") ? " and " : "") + Cents; break;
            }
            if (Cents != " ")
            {
                Cents = Cents + DecName;
            }
            if (Dollars != "" || Cents != "")
                Cents = Cents + " Only";
            return Dollars + Cents;
        }
        //'*******************************************
        //' Converts a number from 100-999 into text *
        //'*******************************************

        public static string GetHundreds(string MyNumber)
        {
            string Result = "";
            int iStart = 0;

            if (MyNumber == "")
                return "";
            if (MyNumber == "0")
                return "";

            MyNumber = "000" + MyNumber;
            iStart = MyNumber.Length - 3;
            MyNumber = MyNumber.Substring(iStart, 3);

            // Convert the hundreds place.
            if (MyNumber.Substring(0, 1) != "0")
                Result = GetDigit(MyNumber.Substring(0, 1)) + " Hundred ";
            // Convert the tens and ones place.

            if (MyNumber.Substring(1, 1) != "0")
                Result = Result + GetTens(MyNumber.Substring(1));
            else
                Result = Result + GetDigit(MyNumber.Substring(2));
            return Result;
        }

        //*********************************************
        //Converts a number from 10 to 99 into text. *
        //*********************************************

        public static string GetTens(string TensText)
        {
            string Result = "";
            if (TensText.Substring(0, 1) == "1")
            {
                switch (TensText)
                {
                    case "10": Result = "Ten"; break;
                    case "11": Result = "Eleven"; break;
                    case "12": Result = "Twelve"; break;
                    case "13": Result = "Thirteen"; break;
                    case "14": Result = "Fourteen"; break;
                    case "15": Result = "Fifteen"; break;
                    case "16": Result = "Sixteen"; break;
                    case "17": Result = "Seventeen"; break;
                    case "18": Result = "Eighteen"; break;
                    case "19": Result = "Nineteen"; break;
                }
            }
            else
            {
                switch (TensText.Substring(0, 1))
                {
                    case "2": Result = "Twenty "; break;
                    case "3": Result = "Thirty "; break;
                    case "4": Result = "Forty "; break;
                    case "5": Result = "Fifty "; break;
                    case "6": Result = "Sixty "; break;
                    case "7": Result = "Seventy "; break;
                    case "8": Result = "Eighty "; break;
                    case "9": Result = "Ninety "; break;
                }
                Result = Result + GetDigit(TensText.Substring(TensText.Length - 1, 1));
            }
            return Result;
        }

        //*******************************************
        // Converts a number from 1 to 9 into text. *
        //*******************************************
        public static string GetDigit(string Digit)
        {
            string sWords = "";
            switch (Digit)
            {
                case "1": sWords = "One"; break;
                case "2": sWords = "Two"; break;
                case "3": sWords = "Three"; break;
                case "4": sWords = "Four"; break;
                case "5": sWords = "Five"; break;
                case "6": sWords = "Six"; break;
                case "7": sWords = "Seven"; break;
                case "8": sWords = "Eight"; break;
                case "9": sWords = "Nine"; break;
            }
            return sWords;
        }


        public static string GetAttention(string Attention)
        {
            string str = "";
            if (Attention.ToString().Trim().Length > 0)
                str = "ATTN : " + Attention.ToString();
            return str;
        }

        public static string GetTelFax(string Tel, string Fax)
        {
            string str = "";
            if (Tel.ToString().Trim().Length > 0)
                str = "TEL : " + Tel.ToString();
            if (Fax.ToString().Trim().Length > 0)
            {
                if (str != "")
                    str += " ";
                str += "FAX : " + Fax.ToString();
            }
            return str;
        }

        public static string GetSqlFormatted(string str)
        {
            return str.Replace("'", "''");
        }


        public static void SaveXmlFile(string FolderName, string PKID, Dictionary<string, string> myData)
        {
            DataSet Dt_Set = new DataSet();
            DataTable Dt_Temp = new DataTable();
            Dt_Temp.Columns.Add("CODE", typeof(System.String));
            Dt_Temp.Columns.Add("DESC", typeof(System.String));
            foreach (KeyValuePair<string, string> pair in myData)
                Dt_Temp.Rows.Add(pair.Key, pair.Value);
            Dt_Set.Tables.Add(Dt_Temp);

            var physicalPath = AppDomain.CurrentDomain.BaseDirectory;
            string sFileName = physicalPath + FolderName + PKID.Replace("-", "") + ".xml";
            Dt_Set.WriteXml(sFileName);
        }

        public static Dictionary<string, string> ReadXmlFile(string FolderName, string PKID)
        {
            var physicalPath = AppDomain.CurrentDomain.BaseDirectory;
            string sFileName = physicalPath + FolderName + PKID.Replace("-", "") + ".xml";

            Dictionary<string, string> myData = null;
            if (System.IO.File.Exists(sFileName))
            {
                DataSet Dt_Set = new DataSet();
                Dt_Set.ReadXml(sFileName);
                foreach (DataRow Dr1 in Dt_Set.Tables[0].Rows)
                {
                    if (myData == null)
                        myData = new Dictionary<string, string>();
                    myData.Add(Dr1["CODE"].ToString(), Dr1["DESC"].ToString());
                }
            }
            return myData;
        }

        public static DataTable ReadXmlFile(int iVer, string FolderName, string PKID)
        {
            var physicalPath = AppDomain.CurrentDomain.BaseDirectory;
            string sFileName = physicalPath + FolderName + PKID.Replace("-", "") + ".xml";
            DataTable myData = new DataTable();

            if (System.IO.File.Exists(sFileName))
            {
                DataSet Dt_Set = new DataSet();
                Dt_Set.ReadXml(sFileName);
                myData = Dt_Set.Tables[0];
            }
            return myData;
        }




        public static string Encrypt(string originalString)
        {
            byte[] bytes = ASCIIEncoding.ASCII.GetBytes("ZeroCool");
            if (String.IsNullOrEmpty(originalString))
            {
                return "";
                //throw new ArgumentNullException("The string which needs to be encrypted can not be null.");
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream();
            CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateEncryptor(bytes, bytes), CryptoStreamMode.Write);

            StreamWriter writer = new StreamWriter(cryptoStream);
            writer.Write(originalString);
            writer.Flush();
            cryptoStream.FlushFinalBlock();
            writer.Flush();

            return Convert.ToBase64String(memoryStream.GetBuffer(), 0, (int)memoryStream.Length);
        }

        public static string Decrypt(string cryptedString)
        {
            byte[] bytes = ASCIIEncoding.ASCII.GetBytes("ZeroCool");
            if (String.IsNullOrEmpty(cryptedString))
            {
                return "";
                // throw new ArgumentNullException("The string which needs to be decrypted can not be null.");
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(cryptedString));
            CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);

            return reader.ReadToEnd();
        }



        public static string[] GetXmlSplitLines(string XmlData)
        {
            string[] sArray = null;
            try
            {
                if (XmlData.Trim().Length > 0)
                {
                    XmlData = XmlData.Replace("\r\n", "\n");
                    XmlData = XmlData.Replace("\r", "\n");
                    sArray = XmlData.Split('\n');
                }
            }
            catch (Exception)
            {
            }
            return sArray;
        }

        public static string Convert_Weight(string sType, string data, int iDec)
        {
            decimal iData = 0;
            try
            {
                if (sType == "KG2LBS")
                    iData = Conv2Decimal(data) * (decimal)2.2046;
                if (sType == "CBM2CFT")
                    iData = Conv2Decimal(data) * (decimal)35.314;
                if (sType == "LBS2KG")
                    iData = Conv2Decimal(data) / (decimal)2.2046;
                if (sType == "CFT2CBM")
                    iData = Conv2Decimal(data) / (decimal)35.314;
            }
            catch (Exception Ex)
            {
                iData = 0;
                throw Ex;
            }
            return NumFormat(iData.ToString(), iDec);
        }


        public static string getFrontEndDate(string sDate)
        {
            string sRet = "";
            if (sDate != "")
            {
                DateTime Dt = DateTime.Parse(sDate.ToString());
                sRet = Dt.ToString("dd-MM-yyyy");
            }
            return sRet;
        }

        public static void WriteData(ExcelWorksheet eWS, int nRow, int nCol, object sValue, Color nColor, Boolean bBold = false, string bBorder = "", string HAlignment = "LEFT", string FontName = "", int FontSize = 9, Boolean WrapText = false, int iRowHt = 325, string dFormat = "", Boolean BlankIfZero = true)
        {

            string str = "";
            str = sValue.GetType().ToString();


            if (FontName == "")
                FontName = "Calibri";

            if (str == "System.Decimal")
            {
                if (BlankIfZero)
                {
                    if ((decimal)sValue != 0)
                        eWS.Cells[nRow, nCol].Value = sValue;
                    else
                        eWS.Cells[nRow, nCol].Value = "";

                }
                else
                    eWS.Cells[nRow, nCol].Value = sValue;
            }
            else
                eWS.Cells[nRow, nCol].Value = sValue;

            eWS.Cells[nRow, nCol].Style.Font.Color = nColor;
            if (bBold)
                eWS.Cells[nRow, nCol].Style.Font.Weight = ExcelFont.BoldWeight;


            if (dFormat != "")
            {
                eWS.Cells[nRow, nCol].Style.NumberFormat = dFormat;
            }

            if (bBorder != "")
            {
                if (bBorder.Contains("L"))
                    eWS.Cells[nRow, nCol].Style.Borders.SetBorders(MultipleBorders.Left, Color.Black, LineStyle.Thin);
                if (bBorder.Contains("R"))
                    eWS.Cells[nRow, nCol].Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, LineStyle.Thin);
                if (bBorder.Contains("B"))
                    eWS.Cells[nRow, nCol].Style.Borders.SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
                if (bBorder.Contains("T"))
                    eWS.Cells[nRow, nCol].Style.Borders.SetBorders(MultipleBorders.Top, Color.Black, LineStyle.Thin);
                if (bBorder.Contains("A"))
                    eWS.Cells[nRow, nCol].Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);

                if (bBorder.Contains("U"))
                {
                    eWS.Cells[nRow, nCol].Style.Font.UnderlineStyle = UnderlineStyle.Single;
                }


            }
            if (HAlignment.ToUpper().Contains("L"))
                eWS.Cells[nRow, nCol].Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
            if (HAlignment.ToUpper().Contains("C"))
                eWS.Cells[nRow, nCol].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            if (HAlignment.ToUpper().Contains("R"))
                eWS.Cells[nRow, nCol].Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;

            if (HAlignment.ToUpper().Contains("T"))
                eWS.Cells[nRow, nCol].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
            else if (HAlignment.ToUpper().Contains("B"))
                eWS.Cells[nRow, nCol].Style.VerticalAlignment = VerticalAlignmentStyle.Bottom;
            else
                eWS.Cells[nRow, nCol].Style.VerticalAlignment = VerticalAlignmentStyle.Center;


            eWS.Rows[nRow].Height = iRowHt; //325;

            eWS.Cells[nRow, nCol].Style.Font.Name = FontName;
            eWS.Cells[nRow, nCol].Style.Font.Size = FontSize * 20;

            eWS.Cells[nRow, nCol].Style.WrapText = WrapText;
        }

        public static void WriteMergeCell(ExcelWorksheet eWS, int rRow, int rCol, int _width,
            int _height, object sData, string _fontName,
            int _fontSize, bool _fontbold, Color _fontcolor, string _halignment, string _valignment, string _mborders, string _linestyle, bool _wraptext = false, bool _UnderLine = false)
        {
            CellRange myCell;
            myCell = eWS.Cells.GetSubrangeRelative(rRow, rCol, _width, _height);
            myCell.Merged = true;
            myCell.Style.WrapText = _wraptext;
            LineStyle _lstyle = LineStyle.None;
            if (_linestyle == "THICK")
                _lstyle = LineStyle.Thick;
            if (_linestyle == "THIN")
                _lstyle = LineStyle.Thin;
            myCell.Style.Font.Color = _fontcolor;
            if (_fontbold)
                myCell.Style.Font.Weight = ExcelFont.BoldWeight;
            myCell.Style.Font.Name = _fontName;
            myCell.Style.Font.Size = _fontSize * 20;
            myCell.Value = sData;
            if (_UnderLine)
                myCell.Style.Font.UnderlineStyle = UnderlineStyle.Single;

            if (_mborders != "")
            {
                if (_mborders.Contains("L"))
                    myCell.Style.Borders.SetBorders(MultipleBorders.Left, Color.Black, _lstyle);
                if (_mborders.Contains("R"))
                    myCell.Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, _lstyle);
                if (_mborders.Contains("B"))
                    myCell.Style.Borders.SetBorders(MultipleBorders.Bottom, Color.Black, _lstyle);
                if (_mborders.Contains("T"))
                    myCell.Style.Borders.SetBorders(MultipleBorders.Top, Color.Black, _lstyle);
                if (_mborders.Contains("A"))
                    myCell.Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, _lstyle);

            }

            if (_halignment.ToUpper() == "L")
                myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
            if (_halignment.ToUpper() == "C")
                myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            if (_halignment.ToUpper() == "R")
                myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;

            if (_valignment.ToUpper() == "T")
                myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Top;
            if (_valignment.ToUpper() == "C")
                myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            if (_valignment.ToUpper() == "B")
                myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Bottom;
            if (_valignment.ToUpper() == "J")
                myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Justify;


        }

        public static string CreateReportFile(string fName)
        {
            string RootPath = System.Web.HttpContext.Current.Server.MapPath(".");
            string Folder = fName;
            string fullName = System.IO.Path.Combine(RootPath, "Reports", Folder);
            CreateFolder(fullName);
            string Fname = fName;
            fullName = System.IO.Path.Combine(RootPath, "Reports", Folder, Fname);
            return fullName;
        }


        public static string OLD_CreateReportFile(string fName)
        {
            string RootPath = System.Web.HttpContext.Current.Server.MapPath(".");
            string Folder = System.Guid.NewGuid().ToString().ToUpper();
            string fullName = System.IO.Path.Combine(RootPath, "Reports", Folder);
            CreateFolder(fullName);
            string Fname = fName;
            fullName = System.IO.Path.Combine(RootPath, "Reports", Folder, Fname);
            return fullName;
        }


        public static Boolean CreateFolder(string Folder, Boolean RaiseErorr = false)
        {
            Boolean bRet = false;
            try
            {
                System.IO.Directory.CreateDirectory(Folder);
                bRet = true;
            }
            catch (Exception ex)
            {
                bRet = false;
                if ( RaiseErorr )
                    throw new Exception(ex.Message);
            }
            return bRet;
        }

        public static Boolean RemoveFolder(string Folder, Boolean RaiseErorr = false)
        {
            Boolean bRet = false;
            try
            {
                if ( System.IO.Directory.Exists(Folder))
                    System.IO.Directory.Delete(Folder,true);
                bRet = true;
            }
            catch (Exception ex)
            {
                bRet = false;
                if (RaiseErorr)
                    throw new Exception(ex.Message);
            }
            return bRet;
        }

        public static Boolean RemoveFile(string sFile, Boolean RaiseErorr =false)
        {
            Boolean bRet = false;
            try
            {
                if ( System.IO.File.Exists(sFile) )
                    System.IO.File.Delete(sFile);
                bRet = true;
            }
            catch (Exception ex)
            {
                bRet = false;
                if (RaiseErorr)
                    throw new Exception(ex.Message);
            }
            return bRet;
        }

        public static string CopyFile(string sFile, string tFile, Boolean RaiseErorr = false)
        {
            string sValue = "";
            try
            {
                if (System.IO.File.Exists(sFile))
                    System.IO.File.Copy(sFile, tFile,true) ;
                sValue = "";
            }
            catch (Exception ex)
            {
                sValue = ex.Message;
                if (RaiseErorr)
                    throw new Exception(ex.Message);
            }
            return sValue;
        }

        public static string CopyFilesFromFolder(string sFolder, string tFolder, Boolean RaiseErorr = false)
        {
            string sValue = "";
            try
            {
                string[] fileEntries = Directory.GetFiles(sFolder);
                foreach (string fileName in fileEntries)
                    System.IO.File.Copy(fileName, Path.Combine(tFolder, Path.GetFileName(fileName)) , true);
                sValue = "";
            }
            catch (Exception ex)
            {
                sValue = ex.Message;
                if (RaiseErorr)
                    throw new Exception(ex.Message);
            }
            return sValue;
        }


        public static string GetFileName(string Report_Folder, string Sub_Folder, string File_Name)
        {
            try
            {
                Report_Folder = System.IO.Path.Combine(Report_Folder, Sub_Folder);
                System.IO.Directory.CreateDirectory(Report_Folder);
                File_Name = System.IO.Path.Combine(Report_Folder, File_Name);
            }
            catch (Exception)
            {
                File_Name = "";
            }
            return File_Name;
        }

        public static string ProperFileName(string str)
        {
            string sRet = str;
            try
            {
                sRet = sRet.Replace("\\", "");
                sRet = sRet.Replace("/", "");
                sRet = sRet.Replace(":", "");
                sRet = sRet.Replace("*", "");
                sRet = sRet.Replace("?", "");
                sRet = sRet.Replace("<", "");
                sRet = sRet.Replace(">", "");
                sRet = sRet.Replace("|", "");
                sRet = sRet.Replace("'", "");
                sRet = sRet.Replace("#", "");
                sRet = sRet.Replace("&", "");
                sRet = sRet.Replace("%", "");
                sRet = sRet.Replace("+", "");
            }
            catch (Exception)
            {
            }
            return sRet;

        }

               

        public static Color GetColor(string sColorCode)
        {
            string str = GetColorName(sColorCode);
            Color clr = Color.Black;

            if (str.ToUpper() == "GREEN")
                clr = Color.Green;
            if (str.ToUpper() == "RED")
                clr = Color.Red;
            if (str.ToUpper() == "GRAY")
                clr = Color.Gray;
            if (str.ToUpper() == "BLUE")
                clr = Color.Blue;
            if (str.ToUpper() == "WHITE")
                clr = Color.White;
            if (str.ToUpper() == "DEEPSKYBLUE")
                clr = Color.DeepSkyBlue;

            return clr;
        }
        public static string GetColorName(string sColorCode)
        {
            sColorCode = sColorCode.Trim();
            string sCname = "BLACK";
            try
            {
                if (sColorCode == "0")
                    sCname = "BLACK";
                else if (sColorCode == "1")
                    sCname = "RED";
                else if (sColorCode == "2")
                    sCname = "DEEPSKYBLUE";
                else if (sColorCode == "3")
                    sCname = "YELLOW";
            }
            catch (Exception)
            {
            }
            return sCname;
        }
        public static string GetRootPath()
        {
            string rPath = "";
            try
            {
                rPath = System.Web.HttpContext.Current.Server.MapPath(".");
            }
            catch (Exception)
            {
                rPath = "";
            }
            return rPath;
        }

        public static decimal RoundNumber_Latest(string sData, int dPlaces, Boolean bRound)
        {
            Decimal nAmt = 0;
            string sFmt;
            try
            {
                sFmt = "{0:F" + dPlaces + "}";
                sData = String.Format(sFmt, Conv2Decimal(sData));
                nAmt = Conv2Decimal(sData);
            }
            catch (Exception)
            {
                nAmt = 0;
            }
            return nAmt;
        }

        public static DataRow ProcessFobInvoice(string invid, Boolean bUpdate = true)
        {
            DataTable DT_INV = new DataTable();
            DataRow DR_IN = null;
            string sql = "";
            DBConnection Con = new DBConnection();
            Boolean IsTrans = false;
            try
            {

                sql = "";
                sql += " select jexp_curr_id, jexp_exrate,jexp_inv_amt, ";
                sql += " 0 as jexp_freight_rate,jexp_freight_amount, jexp_freight_curr_id, ";
                sql += " 0 as jexp_packing_rate,jexp_packing_amount, jexp_packing_curr_id, ";
                sql += " jexp_insurance_rate, jexp_insurance_amount, jexp_insurance_curr_id, ";
                sql += " jexp_commission_rate, jexp_commission_amount, jexp_commission_curr_id, ";
                sql += " jexp_fobdiscount_rate, jexp_fobdiscount_amount, jexp_fobdiscount_curr_id, ";
                sql += " jexp_otherded_rate, jexp_otherded_amount, jexp_otherded_curr_id, ";
                sql += " jexp_add ";
                sql += " from jobexpm where jexp_pkid  = '" + invid + "'";


                DT_INV = Con.ExecuteQuery(sql);

                if (DT_INV.Rows.Count <= 0)
                    return DR_IN;

                DR_IN = DT_INV.Rows[0];

                string curr_id = DR_IN["jexp_curr_id"].ToString();
                decimal exrate = Conv2Decimal(DR_IN["jexp_exrate"].ToString());
                string _inclusive = DR_IN["jexp_add"].ToString();

                decimal _FREIGHT_FC = Conv2Decimal(DR_IN["jexp_freight_amount"].ToString());
                decimal _FREIGHT_RATE_FC = 0;
                if (DR_IN["jexp_freight_curr_id"].ToString() != curr_id)
                    _FREIGHT_FC = _FREIGHT_FC / exrate;

                decimal _INSURANCE_FC = Conv2Decimal(DR_IN["jexp_insurance_amount"].ToString());
                decimal _INSURANCE_RATE_FC = Conv2Decimal(DR_IN["jexp_insurance_rate"].ToString());
                if (DR_IN["jexp_insurance_curr_id"].ToString() != curr_id)
                    _INSURANCE_FC = _INSURANCE_FC / exrate;

                decimal _PACK_CHARGES_FC = Conv2Decimal(DR_IN["jexp_packing_amount"].ToString());
                decimal _PACK_CHARGES_RATE_FC = 0;
                if (DR_IN["jexp_packing_curr_id"].ToString() != curr_id)
                    _PACK_CHARGES_FC = _PACK_CHARGES_FC / exrate;

                decimal _COMMISSION_FC = Conv2Decimal(DR_IN["jexp_commission_amount"].ToString());
                decimal _COMMISSION_RATE_FC = Conv2Decimal(DR_IN["jexp_commission_rate"].ToString());
                if (DR_IN["jexp_commission_curr_id"].ToString() != curr_id)
                    _COMMISSION_FC = _COMMISSION_FC / exrate;

                decimal _FOB_DISC_AMOUNT_FC = Conv2Decimal(DR_IN["jexp_fobdiscount_amount"].ToString());
                decimal _FOB_DISC_AMOUNT_RATE_FC = Conv2Decimal(DR_IN["jexp_fobdiscount_rate"].ToString());
                if (DR_IN["jexp_fobdiscount_curr_id"].ToString() != curr_id)
                    _FOB_DISC_AMOUNT_FC = _FOB_DISC_AMOUNT_FC / exrate;

                decimal _ODED_AMOUNT_FC = Conv2Decimal(DR_IN["jexp_otherded_amount"].ToString());
                decimal _ODED_AMOUNT_RATE_FC = Conv2Decimal(DR_IN["jexp_otherded_rate"].ToString());
                if (DR_IN["jexp_otherded_curr_id"].ToString() != curr_id)
                    _ODED_AMOUNT_FC = _ODED_AMOUNT_FC / exrate;

                decimal _INVOICE_AMOUNT_FC = Conv2Decimal(DR_IN["jexp_inv_amt"].ToString());

                if (_FREIGHT_RATE_FC <= 0 && _INVOICE_AMOUNT_FC > 0)
                    _FREIGHT_RATE_FC = _FREIGHT_FC / _INVOICE_AMOUNT_FC * 100;
                if (_INSURANCE_FC <= 0)
                {
                    _INSURANCE_FC = _INVOICE_AMOUNT_FC * _INSURANCE_RATE_FC / 100;
                    _INSURANCE_FC = RoundNumber_Latest(_INSURANCE_FC.ToString(), 2, true);
                }
                if (_INSURANCE_RATE_FC <= 0 && _INVOICE_AMOUNT_FC > 0)
                    _INSURANCE_RATE_FC = _INSURANCE_FC / _INVOICE_AMOUNT_FC * 100;
                if (_PACK_CHARGES_RATE_FC <= 0 && _INVOICE_AMOUNT_FC > 0)
                    _PACK_CHARGES_RATE_FC = _PACK_CHARGES_FC / _INVOICE_AMOUNT_FC * 100;
                if (_COMMISSION_FC <= 0)
                {
                    _COMMISSION_FC = _INVOICE_AMOUNT_FC * _COMMISSION_RATE_FC / 100;
                    _COMMISSION_FC = RoundNumber_Latest(_COMMISSION_FC.ToString(), 2, true);
                }
                if (_COMMISSION_RATE_FC <= 0 && _INVOICE_AMOUNT_FC > 0)
                    _COMMISSION_RATE_FC = _COMMISSION_FC / _INVOICE_AMOUNT_FC * 100;
                if (_FOB_DISC_AMOUNT_FC <= 0)
                {
                    _FOB_DISC_AMOUNT_FC = _INVOICE_AMOUNT_FC * _FOB_DISC_AMOUNT_RATE_FC / 100;
                    _FOB_DISC_AMOUNT_FC = RoundNumber_Latest(_FOB_DISC_AMOUNT_FC.ToString(), 2, true);
                }
                if (_FOB_DISC_AMOUNT_RATE_FC <= 0 && _INVOICE_AMOUNT_FC > 0)
                    _FOB_DISC_AMOUNT_RATE_FC = _FOB_DISC_AMOUNT_FC / _INVOICE_AMOUNT_FC * 100;
                if (_ODED_AMOUNT_FC <= 0)
                {
                    _ODED_AMOUNT_FC = _INVOICE_AMOUNT_FC * _ODED_AMOUNT_RATE_FC / 100;
                    _ODED_AMOUNT_FC = RoundNumber_Latest(_ODED_AMOUNT_FC.ToString(), 2, true);
                }
                if (_ODED_AMOUNT_RATE_FC <= 0 && _INVOICE_AMOUNT_FC > 0)
                    _ODED_AMOUNT_RATE_FC = _ODED_AMOUNT_FC / _INVOICE_AMOUNT_FC * 100;

                decimal _Inv_Fc = _INVOICE_AMOUNT_FC + _PACK_CHARGES_FC - _FOB_DISC_AMOUNT_FC - _ODED_AMOUNT_FC;
                decimal _Fob_Fc = _Inv_Fc;

                if (_inclusive == "BOTH")
                    _Fob_Fc = _Inv_Fc - _FREIGHT_FC - _INSURANCE_FC;
                if (_inclusive == "FREIGHT")
                    _Fob_Fc = _Inv_Fc - _FREIGHT_FC;
                if (_inclusive == "INSURANCE")
                    _Fob_Fc = _Inv_Fc - _INSURANCE_FC;
                if (_inclusive == "NO")
                    _Fob_Fc = _Inv_Fc;

                if (_inclusive == "FREIGHT")
                    _Inv_Fc += _INSURANCE_FC;
                if (_inclusive == "INSURANCE")
                    _Inv_Fc += _FREIGHT_FC;
                if (_inclusive == "NO")
                    _Inv_Fc += _FREIGHT_FC + _INSURANCE_FC;

                decimal _Inv_Inr = _Inv_Fc * exrate;
                decimal _Fob_Inr = _Fob_Fc * exrate;

                DR_IN["jexp_freight_amount"] = _FREIGHT_FC;
                DR_IN["jexp_freight_rate"] = _FREIGHT_RATE_FC;
                DR_IN["jexp_packing_amount"] = _PACK_CHARGES_FC;
                DR_IN["jexp_packing_rate"] = _PACK_CHARGES_RATE_FC;
                DR_IN["jexp_insurance_amount"] = _INSURANCE_FC;
                DR_IN["jexp_insurance_rate"] = _INSURANCE_RATE_FC;
                DR_IN["jexp_commission_amount"] = _COMMISSION_FC;
                DR_IN["jexp_commission_rate"] = _COMMISSION_RATE_FC;
                DR_IN["jexp_fobdiscount_amount"] = _FOB_DISC_AMOUNT_FC;
                DR_IN["jexp_fobdiscount_rate"] = _FOB_DISC_AMOUNT_RATE_FC;
                DR_IN["jexp_otherded_amount"] = _ODED_AMOUNT_FC;
                DR_IN["jexp_otherded_rate"] = _ODED_AMOUNT_RATE_FC;

                if (bUpdate)
                {
                    sql = "";
                    sql = "update jobexpm set ";
                    sql += " jexp_fob =" + _Fob_Inr.ToString();
                    sql += " ,jexp_fob_fc =" + _Fob_Fc.ToString();
                    sql += " ,jexp_inv_total    =" + _Inv_Inr.ToString();
                    sql += " ,jexp_inv_total_fc =" + _Inv_Fc.ToString();

                    sql += " ,jexp_insurance_amount_4rpt =" + _INSURANCE_FC.ToString();
                    sql += " ,jexp_commission_amount_4rpt =" + _COMMISSION_FC.ToString();
                    sql += " ,jexp_fobdiscount_amount_4rpt =" + _FOB_DISC_AMOUNT_FC.ToString();
                    sql += " ,jexp_otherded_amount_4rpt =" + _ODED_AMOUNT_FC.ToString();

                    sql += " where jexp_pkid='" + invid + "'";

                    Con.BeginTransaction();
                    IsTrans = true;
                    Con.ExecuteNonQuery(sql);
                    Con.CommitTransaction();
                    IsTrans = false;
                }
            }
            catch (Exception)
            {
                if (IsTrans)
                    Con.RollbackTransaction();
            }
            Con.CloseConnection();
            return DR_IN;
        }

        public static void ProcessFobItem(string invid = "", string jobid = "", string itmid = "")
        {
            string id = "";
            string sql = "";
            DBConnection Con = new DBConnection();
            Boolean IsTrans = false;
            try
            {

                sql = "";
                sql += "select itm_pkid, itm_amount, itm_strrefund_no, itm_strrefund_rate, ";
                sql += "itm_dbk_code,itm_dbk_qty, itm_dbk_rate, itm_dbk_valuecap, b.param_code as itm_scheme_code,";
                sql += " itm_rosl_rate,itm_rosl_valuecap,itm_rosl_ctl_rate,itm_rosl_ctl_valuecap ";
                sql += "from itemm a left join param b on itm_scheme_id = b.param_pkid ";
                sql += " where 1=1 ";
                if (invid != "")
                    sql += " and itm_invoice_id = '" + invid + "'";
                if (jobid != "")
                    sql += " and itm_job_id = '" + jobid + "'";
                if (itmid != "")
                    sql += " and itm_pkid = '" + itmid + "'";

                DataTable Dt_Itm = new DataTable();
                Dt_Itm = Con.ExecuteQuery(sql);

                DataRow DR_IN = ProcessFobInvoice(invid, false);

                if (DR_IN == null)
                    return;

                string curr_id = DR_IN["jexp_curr_id"].ToString();
                decimal exrate = Conv2Decimal(DR_IN["jexp_exrate"].ToString());
                string _inclusive = DR_IN["jexp_add"].ToString();
                decimal _FREIGHT_RATE_FC = Conv2Decimal(DR_IN["jexp_freight_rate"].ToString());
                decimal _INSURANCE_RATE_FC = Conv2Decimal(DR_IN["jexp_insurance_rate"].ToString());
                decimal _COMMISSION_RATE_FC = Conv2Decimal(DR_IN["jexp_commission_rate"].ToString());
                decimal _PACK_CHARGES_RATE_FC = Conv2Decimal(DR_IN["jexp_packing_rate"].ToString());
                decimal _FOB_DISC_AMOUNT_RATE_FC = Conv2Decimal(DR_IN["jexp_fobdiscount_rate"].ToString());
                decimal _ODED_AMOUNT_RATE_FC = Conv2Decimal(DR_IN["jexp_otherded_rate"].ToString());

                decimal _RowFob_Inr = 0;

                foreach (DataRow Dr in Dt_Itm.Rows)
                {
                    id = Dr["itm_pkid"].ToString();
                    decimal _RowAmt_Fc = Conv2Decimal(Dr["itm_amount"].ToString());
                    decimal _RowFrt_Fc = _RowAmt_Fc * _FREIGHT_RATE_FC / 100;
                    decimal _RowPacking_Fc = _RowAmt_Fc * _PACK_CHARGES_RATE_FC / 100;
                    decimal _RowIns_Fc = _RowAmt_Fc * _INSURANCE_RATE_FC / 100;
                    decimal _RowDiscount_Fc = _RowAmt_Fc * _FOB_DISC_AMOUNT_RATE_FC / 100;
                    decimal _RowOther_Fc = _RowAmt_Fc * _ODED_AMOUNT_RATE_FC / 100;
                    decimal _RowComm_Fc = _RowAmt_Fc * _COMMISSION_RATE_FC / 100;

                    // invamt + packing  - fob disc - other ded
                    decimal _RowInv_Fc = _RowAmt_Fc + _RowPacking_Fc - _RowDiscount_Fc - _RowOther_Fc;
                    decimal _RowFob_Fc = _RowInv_Fc;

                    if (_inclusive == "BOTH")
                        _RowFob_Fc = _RowInv_Fc - _RowFrt_Fc - _RowIns_Fc;
                    if (_inclusive == "FREIGHT")
                        _RowFob_Fc = _RowInv_Fc - _RowFrt_Fc;
                    if (_inclusive == "INSURANCE")
                        _RowFob_Fc = _RowInv_Fc - _RowIns_Fc;
                    if (_inclusive == "NO")
                        _RowFob_Fc = _RowInv_Fc;

                    _RowFob_Fc = RoundNumber_Latest(_RowFob_Fc.ToString(), 2, true);
                    _RowFob_Inr = _RowFob_Fc * exrate;
                    _RowFob_Inr = RoundNumber_Latest(_RowFob_Inr.ToString(), 2, true);


                    decimal _iRowStrAmt = 0;
                    if (Dr["itm_strrefund_no"].ToString().Trim().Length > 0)
                    {
                        _iRowStrAmt = _RowFob_Inr * Conv2Decimal(Dr["itm_strrefund_rate"].ToString()) / 100;
                        _iRowStrAmt = RoundNumber_Latest(_iRowStrAmt.ToString(), 2, true);
                    }

                    decimal _iRowDbkAmt = 0;
                    decimal _iRowRoslAmt = 0;
                    decimal _nTotRate = 0;
                    decimal _nDBKValue = 0;
                    decimal _nValueCap = 0;

                    decimal _nRoslValue = 0;
                    decimal _nRoslValueCap = 0;



                    if (!Dr["itm_dbk_code"].Equals(DBNull.Value))
                    {
                        _nTotRate = Conv2Decimal(Dr["itm_dbk_rate"].ToString());
                        _nDBKValue = _RowFob_Inr * _nTotRate / 100;
                        _nValueCap = Conv2Decimal(Dr["itm_dbk_qty"].ToString()) * Conv2Decimal(Dr["itm_dbk_valuecap"].ToString());

                        if (_nValueCap == 0 || (_nDBKValue > 0 && _nDBKValue < _nValueCap))
                        {
                            //Dr["DBK_CALCON"] = "R";
                            _iRowDbkAmt = _nDBKValue;
                        }
                        else
                        {
                            //Dr["DBK_CALCON"] = "V";
                            _iRowDbkAmt = _nValueCap;
                        }
                        //ISROSL // 62 and 63 have been added after roslctl
                        /*
                        if (Dr["itm_scheme_code"].ToString() == "60" || Dr["itm_scheme_code"].ToString() == "61" ||
                            Dr["itm_scheme_code"].ToString() == "62" || Dr["itm_scheme_code"].ToString() == "63" ||
                            Dr["itm_scheme_code"].ToString() == "64" || Dr["itm_scheme_code"].ToString() == "65")
                        {
                        */
                        // State Rosl 
                        _nRoslValue = _RowFob_Inr * Conv2Decimal(Dr["itm_rosl_rate"].ToString()) / 100;
                        _nRoslValueCap = Conv2Decimal(Dr["itm_dbk_qty"].ToString()) * Conv2Decimal(Dr["itm_rosl_valuecap"].ToString());
                        if (_nRoslValueCap == 0 || (_nRoslValue > 0 && _nRoslValue < _nRoslValueCap))
                            _iRowRoslAmt = _nRoslValue;
                        else
                            _iRowRoslAmt = _nRoslValueCap;

                        // Central Rosl
                        _nRoslValue = _RowFob_Inr * Conv2Decimal(Dr["itm_rosl_ctl_rate"].ToString()) / 100;
                        _nRoslValueCap = Conv2Decimal(Dr["itm_dbk_qty"].ToString()) * Conv2Decimal(Dr["itm_rosl_ctl_valuecap"].ToString());
                        if (_nRoslValueCap == 0 || (_nRoslValue > 0 && _nRoslValue < _nRoslValueCap))
                            _iRowRoslAmt += _nRoslValue;
                        else
                            _iRowRoslAmt += _nRoslValueCap;


                        //}
                        _iRowDbkAmt = RoundNumber_Latest(_iRowDbkAmt.ToString(), 2, true);
                        _iRowRoslAmt = RoundNumber_Latest(_iRowRoslAmt.ToString(), 2, true);
                    }

                    sql = " update itemm set ";
                    sql += " itm_fob = " + _RowFob_Inr.ToString();
                    sql += " ,itm_fob_fc = " + _RowFob_Fc.ToString();
                    sql += " ,itm_dbk_value = " + _iRowDbkAmt.ToString();
                    sql += " ,itm_rosl_value = " + _iRowRoslAmt.ToString();
                    sql += " ,itm_str_value = " + _iRowStrAmt.ToString();
                    sql += " where itm_pkid = '" + id + "'";


                    Con.BeginTransaction();
                    IsTrans = true;
                    Con.ExecuteQuery(sql);
                    Con.CommitTransaction();
                    IsTrans = false;
                }
            }
            catch (Exception)
            {
                if (IsTrans)
                    Con.RollbackTransaction();
            }
            Con.CloseConnection();
        }

        public static void UpdateFobsummary(string invid = "", string jobid = "")
        {
            DataTable DT_INV = new DataTable();
            string sql = "";
            string sql2 = "";
            DBConnection Con = new DBConnection();
            Boolean IsTrans = false;
            try
            {
                sql = "";
                if (invid != "")
                {
                    sql += " update jobexpm  set ";
                    sql += " (JEXP_ITM_AMOUNT, JEXP_PMV_TOTAL, JEXP_STR_VALUE, JEXP_DBK_VALUE, JEXP_ROSL_VALUE) = ";
                    sql += " (select sum(ITM_AMOUNT),sum(ITM_PMV_TOTAL),sum(ITM_STR_VALUE),sum(ITM_DBK_VALUE),sum(ITM_ROSL_VALUE) ";
                    sql += " from itemm   where jexp_pkid = itm_invoice_id and itm_invoice_id = '" + invid + "') ";
                    sql += " where jexp_pkid = '" + invid + "' ";
                }
                sql2 = "";
                if (jobid != "")
                {
                    sql2 += " update jobm set ";
                    sql2 += " (JOB_FOB, JOB_FOB_FC, JOB_PMV_TOTAL, JOB_STR_VALUE, JOB_DBK_VALUE, JOB_ROSL_VALUE) = ";
                    sql2 += " (select ";
                    sql2 += " sum(jexp_fob), sum(jexp_fob_fc),sum(jexp_pmv_total), sum(jexp_str_value), sum(jexp_dbk_value), sum(jexp_rosl_value) ";
                    sql2 += " from jobexpm where job_pkid = jexp_job_id  and jexp_job_id = '" + jobid + "') ";
                    sql2 += " where job_pkid = '" + jobid + "'";
                }
                Con.BeginTransaction();
                IsTrans = true;
                if (invid != "")
                    Con.ExecuteNonQuery(sql);
                if (jobid != "")
                    Con.ExecuteNonQuery(sql2);
                Con.CommitTransaction();
                IsTrans = false;
            }
            catch (Exception)
            {
                if (IsTrans)
                    Con.RollbackTransaction();
            }
            Con.CloseConnection();
        }

        public static string GetDecimalName(string CurrencyCode)
        {
            if (CurrencyCode == "INR")
                return "Paise";
            else if (CurrencyCode == "USD")
                return "Cents";
            else if (CurrencyCode == "AED")
                return "Fils";
            else
                return "";
        }
        public static float GetWordWidth(string sWord, string fontName, float fsize, FontStyle fStyle)
        {
            Font sfont = new Font(fontName, fsize, fStyle);
            Bitmap b = new Bitmap(2200, 2200);
            Graphics g = Graphics.FromImage(b);
            SizeF sizeOfString = new SizeF();
            sizeOfString = g.MeasureString(sWord, sfont);
            b.Dispose();
            g.Dispose();
            return sizeOfString.Width;
        }


        public static string getCCType(string main_code, Boolean JobWise = false)
        {
            string sdata = "HBL WISE";
            if (main_code == "1101" || main_code == "1102" || main_code == "1103" || main_code == "1104" || main_code == "1201" || main_code == "1202" || main_code == "1203" || main_code == "1204")
                sdata = "JOB WISE";
            if (main_code == "1107" && JobWise)
                sdata = "JOB WISE";
            return sdata;
        }

        public static DataTable getCCJOBS(string cc_category, string cc_id)
        {
            DataTable Dt_cc = new DataTable();
            string sql = "";

            string scatg = "";
            if (cc_category == "SI SEA EXPORT")
                scatg = "JOB SEA EXPORT";
            if (cc_category == "SI AIR EXPORT")
                scatg = "JOB AIR EXPORT";

            if (cc_category == "MBL SEA EXPORT")
                scatg = "JOB SEA EXPORT";
            if (cc_category == "MAWB AIR EXPORT")
                scatg = "JOB AIR EXPORT";



            if (cc_category == "SI SEA EXPORT" || cc_category == "SI AIR EXPORT")
            {
                sql = "select '" + scatg + "' as cc_category, job_pkid as id, count(*) over() as tot ";
                sql += " from jobm where jobs_hbl_id = '" + cc_id + "'";
            }
            if (cc_category == "MBL SEA EXPORT" || cc_category == "MAWB AIR EXPORT")
            {
                sql = "select '" + scatg + "' as cc_category, job_pkid as id, count(*) over() as tot ";
                sql += " from hblm a inner join jobm on hbl_pkid = jobs_hbl_id ";
                sql += " where hbl_mbl_id = '" + cc_id + "'";
            }


            if (sql != "")
            {
                DBConnection Con = new DBConnection();
                Dt_cc = Con.ExecuteQuery(sql);
                Con.CloseConnection();
            }

            return Dt_cc;
        }

        public static DataTable getCCHBLS(string cc_category, string cc_id)
        {
            DataTable Dt_cc = new DataTable();
            DataTable Dt_cost = new DataTable();
            string sql = "";
            string sql1 = "";
            string scatg = cc_category;
            if (cc_category == "SI SEA EXPORT" || cc_category == "SI AIR EXPORT" || cc_category == "SI SEA IMPORT" || cc_category == "SI AIR IMPORT" || cc_category == "GENERAL JOB")
            {
                sql = "select '" + scatg + "' as cc_category, ";
                sql += "'" + cc_category + "' as ccgroup, ";
                sql += " hbl_pkid as id, count(*) over() as tot, hbl_chwt, hbl_chwt as tot_chwt,hbl_cbm, hbl_cbm as tot_cbm ";
                sql += " from hblm where hbl_pkid = '" + cc_id + "'";
            }


            if (cc_category == "MBL SEA EXPORT" || cc_category == "MBL SEA IMPORT" || cc_category == "MAWB AIR EXPORT" || cc_category == "MAWB AIR IMPORT")
            {
                if ( scatg.Contains ("MBL") )
                    scatg = scatg.Replace("MBL", "SI");
                if (scatg.Contains("MAWB"))
                    scatg = scatg.Replace("MAWB", "SI");
                sql = "select '" + scatg + "' as cc_category,";
                sql += "'" + cc_category + "' as ccgroup,hbl_ddp, '{TYPE}' hbl_shipment_type, '{DDP}' as costing_ddp, ";
                sql += " hbl_pkid as id,count(*) over() as tot, ";
                sql += " hbl_chwt,sum(hbl_chwt) over(partition by hbl_mbl_id) as tot_chwt, ";
                sql += " hbl_cbm,sum(hbl_cbm) over(partition by hbl_mbl_id) as tot_cbm, ";
                sql += " sum(case when hbl_ddp = 'YES' then 1 else 0 end) over(partition by hbl_mbl_id) as tot_ddp ";
                sql += " from hblm a ";
                sql += " where a.hbl_mbl_id = '" + cc_id + "'";

                sql1 = " select hbl_shipment_type, cost_ddp from hblm a left ";
                sql1 += " join costingm b on a.hbl_pkid = cost_pkid ";
                sql1 += " where hbl_pkid = '" + cc_id + "' ";

            }

            if (sql != "")
            {
                DBConnection Con = new DBConnection();

                if (sql1 != "")
                {
                    Dt_cost = Con.ExecuteQuery(sql1);
                    foreach (DataRow Dr in Dt_cost.Rows)
                    {
                        sql = sql.Replace("{TYPE}", Dr["hbl_shipment_type"].ToString());
                        sql = sql.Replace("{DDP}", Dr["cost_ddp"].ToString());
                        break;
                    }
                }

                Dt_cc = Con.ExecuteQuery(sql);
                Con.CloseConnection();
            }
            return Dt_cc;
        }


        public static DataTable getCCCntrs(string cc_category, string cc_id)
        {
            DataTable Dt_cc = new DataTable();
            string sql = "";
            string scatg = cc_category;
            if (cc_category == "SI SEA EXPORT")
            {
                sql = " select 'CNTR SEA EXPORT' as cc_category,hbl_cntr_id as id, count(*) over() as tot from hblcontainer where hbl_id = '" + cc_id + "'";
            }
            if (cc_category == "MBL SEA EXPORT")
            {
                sql += " select 'CNTR SEA EXPORT' as cc_category,hbl_cntr_id as id , count(*) over() as tot from( ";
                sql += " select distinct hbl_cntr_id from hblcontainer where hbl_id in ";
                sql += " (select hbl_pkid from hblm where hbl_mbl_id = '" + cc_id + "') ";
                sql += " ) ";
            }
            if (sql != "")
            {
                DBConnection Con = new DBConnection();
                Dt_cc = Con.ExecuteQuery(sql);
                Con.CloseConnection();
            }
            return Dt_cc;
        }


        public static Boolean IsValidSalesman(string ShipperID)
        {
            Boolean bRet = false;
            if (ShipperID != "")
            {
                DBConnection Con_Oracle = new DBConnection();
                string sql = "";
                try
                {
                    DataTable Dt_Temp = new DataTable();
                    sql = "  select nvl(b.rec_locked,'N') as rec_locked from customerm a inner join ";
                    sql += " param b on cust_sman_id = param_pkid   where cust_pkid = '{CUSTID}' ";
                    sql = sql.Replace("{CUSTID}", ShipperID);
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        if (Dt_Temp.Rows[0]["rec_locked"].ToString() == "N")
                        {
                            bRet = true;
                        }
                    }
                    Con_Oracle.CloseConnection();
                }
                catch (Exception)
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                }
            }
            return bRet;
        }



        public static string getSalesmanID(string ShipperID, Boolean OnlyValid = true)
        {
            string sID = "";
            if (ShipperID != "")
            {
                DBConnection Con_Oracle = new DBConnection();
                string sql = "";
                try
                {
                    DataTable Dt_Temp = new DataTable();
                    sql = "  select cust_sman_id from customerm where cust_pkid = '{CUSTID}'";
                    sql = sql.Replace("{CUSTID}", ShipperID);
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                        sID = Dt_Temp.Rows[0]["cust_sman_id"].ToString();

                    if (sID.ToString() == "")
                    {
                        Dt_Temp = new DataTable();
                        sql = "select param_pkid from param where param_type = 'SALESMAN' and param_code = 'NA'";
                        Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                        if (Dt_Temp.Rows.Count > 0)
                            sID = Dt_Temp.Rows[0]["param_pkid"].ToString();
                    }
                    Con_Oracle.CloseConnection();
                }
                catch (Exception)
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    //throw Ex;

                }
            }
            return sID;
        }
        public static string getNomination(string ConsigneeID)
        {
            string sNom = "";
            if (ConsigneeID != "")
            {
                DBConnection Con_Oracle = new DBConnection();
                string sql = "";
                try
                {
                    DataTable Dt_Temp = new DataTable();
                    sql = " select cust_nomination from customerm where cust_pkid = '{CUSTID}' ";
                    sql = sql.Replace("{CUSTID}", ConsigneeID);
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                        sNom = Dt_Temp.Rows[0]["cust_nomination"].ToString();

                    if (sNom.ToString().Trim() == "")
                    {
                        sNom = "NA";
                    }
                    Con_Oracle.CloseConnection();
                }
                catch (Exception)
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    //throw Ex;
                }
            }
            return sNom;
        }

        public static string GetPortCode(string sCode)
        {
            if (sCode.Length >= 5)
                sCode = sCode.Substring(2, 3);
            return sCode;
        }
        public static string GetTruncated(string str, int slen)
        {
            if (str.Length > slen)
                str = str.Substring(0, slen);
            return str;
        }
        public static Boolean IsJobLocked(string JobID)
        {
            bool sLock = false;
            bool sRowExist = false;
            string editcode = "";
            if (JobID != "")
            {
                DBConnection Con_Oracle = new DBConnection();
                string sql = "";
                try
                {
                    sql = " select m.hbl_edit_code as mbl_edit_code from jobm a";
                    sql += "  inner join hblm h on a.jobs_hbl_id = h.hbl_pkid";
                    sql += "  inner join hblm m on h.hbl_mbl_id = m.hbl_pkid";
                    sql += "  where a.job_pkid ='{JOBID}'";
                    sql = sql.Replace("{JOBID}", JobID);
                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        sRowExist = true;
                        editcode = Dt_Temp.Rows[0]["mbl_edit_code"].ToString();
                    }
                    if (editcode.Trim() == "")
                    {

                        sql = " select h.hbl_edit_code as hbl_edit_code from jobm a";
                        sql += "  inner join hblm h on a.jobs_hbl_id = h.hbl_pkid";
                        sql += "  where a.job_pkid ='{JOBID}'";
                        sql = sql.Replace("{JOBID}", JobID);
                        Dt_Temp = new DataTable();
                        Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                        if (Dt_Temp.Rows.Count > 0)
                        {
                            sRowExist = true;
                            editcode = Dt_Temp.Rows[0]["hbl_edit_code"].ToString();
                        }
                    }
                    Dt_Temp.Rows.Clear();
                    Con_Oracle.CloseConnection();

                    if (sRowExist)
                    {
                        sLock = true;
                        if (editcode.IndexOf("{S}") >= 0)
                            sLock = false;
                    }
                }
                catch (Exception Ex)
                {
                    sLock = false;
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    Con_Oracle.CreateErrorLog("IsJobLockedError " + Ex.Message.ToString());
                }
            }
            return sLock;
        }

        public static string IsDateLocked(string JvhDate, string JvhType, string JvhCompCode, string JvhBranchCode, string JvhYearCode)
        {
            string sLock = "";
            if (JvhDate == null || JvhDate.ToString() == "")
                sLock = "";
            else
            {
                DBConnection Con_Oracle = new DBConnection();
                string sql = "";
                try
                {

                    JvhType = JvhType.Replace("-", "_");

                    string sLockField = "LOCK_" + JvhType;

                    sql = " select lock_pkid," + sLockField + " from lockingm";
                    sql += " where rec_company_code = '{COMPCODE}' ";
                    sql += " and rec_branch_code = '{BRANCHCODE}' ";
                    sql += " and lock_year = {YEARCODE}";
                    sql += " and {LOCKTYPE} >= '{JVDATE}' ";

                    sql = sql.Replace("{COMPCODE}", JvhCompCode);
                    sql = sql.Replace("{BRANCHCODE}", JvhBranchCode);
                    sql = sql.Replace("{YEARCODE}", JvhYearCode);
                    sql = sql.Replace("{JVDATE}", JvhDate);
                    sql = sql.Replace("{LOCKTYPE}", sLockField);

                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        sLock = " Transactions Locked Upto " + Lib.DatetoStringDisplayformat(Dt_Temp.Rows[0][sLockField]);
                    }
                    Dt_Temp.Rows.Clear();
                    Con_Oracle.CloseConnection();
                }
                catch (Exception Ex)
                {
                    sLock = Ex.Message;
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    Con_Oracle.CreateErrorLog("IsDateLockedError " + Ex.Message.ToString());
                }
            }
            return sLock;
        }
        public static void AuditLog(string module, string type, string action, string comp_code, string branch_code, string user_code,
            string pkey, string refno = "", string new_remarks = "",
            decimal old_amt = 0, decimal new_amt = 0, string old_remarks = "", string computer = ""
            )
        {
            string sql = "";
            DBConnection Con_Oracle = new DBConnection();
            try
            {
                DBRecord Rec = new DBRecord();
                Rec.CreateRow("auditlog", "ADD", "audit_pkey", pkey);
                Rec.InsertString("audit_module", module);
                Rec.InsertString("audit_type", type);
                Rec.InsertString("audit_user_code", user_code);
                Rec.InsertString("audit_comp_code", comp_code);
                Rec.InsertString("audit_branch_code", branch_code);
                Rec.InsertString("audit_action", action);
                Rec.InsertString("audit_refno", refno);
                Rec.InsertString("audit_computer", computer);
                Rec.InsertNumeric("audit_old_amt", old_amt.ToString());
                Rec.InsertNumeric("audit_new_amt", new_amt.ToString());
                if (old_remarks.ToString().Length > 250)
                    Rec.InsertString("audit_old_remarks", old_remarks.Substring(0, 250));
                else
                    Rec.InsertString("audit_old_remarks", old_remarks);
                if (new_remarks.ToString().Length > 250)
                    Rec.InsertString("audit_remarks", new_remarks.Substring(0, 250));
                else
                    Rec.InsertString("audit_remarks", new_remarks);
                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                    Con_Oracle.CreateErrorLog("AuditLogError " + Ex.Message.ToString());
                }
            }
        }

        public static void FtpLog(string from, string to, string action, string module, string moduleid, string subject, bool isack,
            string processid, string company_code, string branch_code, string user_code, string remarks, string file_path = "", string UpdateSql = "")
        {
            string sql = "";
            DBConnection Con_Oracle = new DBConnection();
            try
            {
                if (subject.Length > 100)
                    subject = subject.Substring(0, 100);

                if (processid.Length > 100)
                    processid = processid.Substring(0, 100);

                if (file_path.Length > 100)
                    file_path = file_path.Substring(0, 100);

                if (remarks.Length > 250)
                    remarks = remarks.Substring(0, 250);

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("ftplog", "ADD", "ftp_pkid", Guid.NewGuid().ToString().ToUpper());
                Rec.InsertString("ftp_from", from);
                Rec.InsertString("ftp_to", to);
                Rec.InsertFunction("ftp_date", "SYSDATE");
                Rec.InsertString("ftp_action", action);
                Rec.InsertString("ftp_module", module);
                Rec.InsertString("ftp_module_pkid", moduleid);
                Rec.InsertString("ftp_subject", subject);
                Rec.InsertString("ftp_is_ack", isack == true ? "Y" : "N");
                Rec.InsertString("ftp_process_id", processid);
                Rec.InsertString("ftp_comp_code", company_code);
                Rec.InsertString("ftp_branch_code", branch_code);
                Rec.InsertString("ftp_user_code", user_code);
                Rec.InsertString("ftp_remarks", remarks);
                Rec.InsertString("ftp_isread", "N");
                Rec.InsertString("ftp_file_path", file_path);

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();

                if (UpdateSql != "")
                {
                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(UpdateSql);
                    Con_Oracle.CommitTransaction();
                }
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                    Con_Oracle.CreateErrorLog("FtpLogError " + Ex.Message.ToString());
                }
            }
        }


        public static void UpdateCCusingMBLID(string mblid)
        {
            string sql = "";
                sql = "select jvh_pkid from ledgerh where jvh_type in('PN','HO','IN-ES') and jvh_cc_id = '" + mblid + "'";

            DBConnection Con_Oracle = new DBConnection();
            Con_Oracle = new DBConnection();
            DataTable Dt_test = new DataTable();
            Dt_test = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();

            foreach (DataRow Dr in Dt_test.Rows)
            {
                Lib.UpdateCC(Dr["jvh_pkid"].ToString());
            }
            Dt_test.Rows.Clear();

            DataTable Dt_House = new DataTable();
            sql = " select jvh_pkid from hblm a ";
            sql += " inner join ledgerh c on a.hbl_pkid = c.jvh_cc_id ";
            sql += " where hbl_mbl_id = '" + mblid + "'";
            Dt_House = new DataTable();
            Dt_House = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow Drh in Dt_House.Rows)
            {
                Lib.UpdateCC(Drh["jvh_pkid"].ToString());
            }
            Dt_House.Rows.Clear();
            Dt_House = null;


        }


        public static void UpdateCCusingHBLID(string hblid)
        {
            string sql = "";
            DataTable Dt_House = new DataTable();
            sql = " select jvh_pkid from hblm a ";
            sql += " inner join ledgerh c on a.hbl_pkid = c.jvh_cc_id ";
            sql += " where hbl_pkid = '" + hblid + "'";
            Dt_House = new DataTable();

            DBConnection Con_Oracle = new DBConnection();
            Con_Oracle = new DBConnection();
            Dt_House = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow Drh in Dt_House.Rows)
            {
                Lib.UpdateCC(Drh["jvh_pkid"].ToString());
            }
            Dt_House.Rows.Clear();
            Dt_House = null;
        }


        public static void UpdateCCusingJOBID(string jobid)
        {
            string sql = "";
            DataTable Dt_House = new DataTable();
            sql = " select jvh_pkid from jobm a ";
            sql += " inner join ledgerh b on a.jobs_hbl_id = b.jvh_cc_id ";
            sql += " where job_pkid = '" + jobid + "'";
            Dt_House = new DataTable();

            DBConnection Con_Oracle = new DBConnection();
            Con_Oracle = new DBConnection();
            Dt_House = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow Drh in Dt_House.Rows)
            {
                Lib.UpdateCC(Drh["jvh_pkid"].ToString());
            }
            Dt_House.Rows.Clear();
            Dt_House = null;
        }




        public static void UpdateCC(string jvh_pkid)
        {
            string sql = "";

            Boolean bCntrExists = false;

            Boolean jobwise = false;

            DataTable Dt_cc_jobs = new DataTable();
            DataTable Dt_cc_hbls = new DataTable();
            DataTable Dt_cc_cntr = new DataTable();

            DataRow Record = null;
            int iCtr = 0;
            decimal cc_amt = 0;
            DBConnection Con_Oracle = new DBConnection();
            DBRecord Rec = null;

            try
            {
                sql = "";
                sql += " select jvh_pkid, jvh_year, jvh_date, jvh_cc_category, jvh_cc_id, hbl_mbl_id, ";
                sql += " jv_pkid, jv_acc_id,acc_main_code as jv_acc_main_code, acc_code as jv_acc_code, (nvl(jv_debit,0) + nvl(jv_credit,0)) as jv_total, ";
                sql += " a.rec_company_code, a.rec_branch_code ";
                sql += " from ledgerh a inner ";
                sql += " join ledgert b on jvh_pkid = jv_parent_id ";
                sql += " inner join acctm c on jv_acc_id = acc_pkid ";
                sql += " left join hblm on jvh_cc_id = hbl_pkid";
                sql += " where jvh_pkid = '" + jvh_pkid + "' and acc_cost_centre = 'Y' ";
                sql += " order by jv_ctr ";

                DataTable Dt_LedgerList = new DataTable();

                Dt_LedgerList = Con_Oracle.ExecuteQuery(sql);

                if (Dt_LedgerList.Rows.Count <= 0)
                {
                    Con_Oracle.CloseConnection();
                    return;
                }

                Record = Dt_LedgerList.Rows[0];

                Dt_cc_jobs = Lib.getCCJOBS(Record["jvh_cc_category"].ToString(), Record["jvh_cc_id"].ToString());
                Dt_cc_hbls = Lib.getCCHBLS(Record["jvh_cc_category"].ToString(), Record["jvh_cc_id"].ToString());
                Dt_cc_cntr = Lib.getCCCntrs(Record["jvh_cc_category"].ToString(), Record["jvh_cc_id"].ToString());

                /*
                jobwise = false;
                if (Record["jvh_cc_category"].ToString() == "SI SEA EXPORT" && Record["jv_acc_main_code"].ToString() == "1107")
                {
                    if (Record["hbl_mbl_id"].Equals(DBNull.Value))
                    {
                        jobwise = true;
                    }
                }
                */

                Con_Oracle.BeginTransaction();

                sql = "delete from  costcentert where ct_jvh_id ='" + jvh_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);



                foreach (DataRow Row in Dt_LedgerList.Rows)
                {
                    bCntrExists = false;


                    jobwise = false;
                    if (Record["jvh_cc_category"].ToString() == "SI SEA EXPORT" && Row["jv_acc_main_code"].ToString() == "1107")
                    {
                        if (Record["hbl_mbl_id"].Equals(DBNull.Value))
                        {
                            jobwise = true;
                        }
                    }

                    if (Record["jvh_cc_category"].ToString() == "GENERAL JOB")
                    {
                        foreach (DataRow Dr in Dt_cc_hbls.Rows)
                        {
                            iCtr++;
                            cc_amt = Lib.Conv2Decimal(Row["jv_total"].ToString()) / Lib.Conv2Decimal(Dr["tot"].ToString());
                            cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);

                            Rec = new DBRecord();

                            Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                            Rec.InsertString("ct_jvh_id", jvh_pkid);

                            Rec.InsertNumeric("ct_year", Record["jvh_year"].ToString());
                            Rec.InsertString("ct_jv_id", Row["jv_pkid"].ToString());
                            Rec.InsertString("ct_acc_id", Row["jv_acc_id"].ToString());
                            Rec.InsertString("ct_category", Dr["cc_category"].ToString());
                            Rec.InsertString("ct_cost_id", Dr["id"].ToString());
                            Rec.InsertNumeric("ct_cost_year", Record["jvh_year"].ToString());
                            Rec.InsertNumeric("ct_amount", cc_amt.ToString());

                            Rec.InsertString("ct_type", "M");
                            Rec.InsertString("ct_posted", "Y");
                            Rec.InsertString("ct_remarks", "");
                            Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                            Rec.InsertString("rec_company_code", Record["rec_company_code"].ToString());
                            Rec.InsertString("rec_branch_code", Record["rec_branch_code"].ToString());

                            sql = Rec.UpdateRow();
                            Con_Oracle.ExecuteNonQuery(sql);
                        }

                    }
                    else
                    {
                        if (Lib.getCCType(Row["jv_acc_main_code"].ToString(), jobwise) == "JOB WISE")
                        {
                            foreach (DataRow Dr in Dt_cc_jobs.Rows)
                            {
                                iCtr++;
                                cc_amt = Lib.Conv2Decimal(Row["jv_total"].ToString()) / Lib.Conv2Decimal(Dr["tot"].ToString());
                                cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);

                                Rec = new DBRecord();

                                Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                                Rec.InsertString("ct_jvh_id", jvh_pkid);
                                Rec.InsertNumeric("ct_year", Record["jvh_year"].ToString());
                                Rec.InsertString("ct_jv_id", Row["jv_pkid"].ToString());
                                Rec.InsertString("ct_acc_id", Row["jv_acc_id"].ToString());
                                Rec.InsertString("ct_category", Dr["cc_category"].ToString());
                                Rec.InsertString("ct_cost_id", Dr["id"].ToString());
                                Rec.InsertNumeric("ct_cost_year", Record["jvh_year"].ToString());
                                Rec.InsertNumeric("ct_amount", cc_amt.ToString());
                                Rec.InsertString("ct_type", "M");
                                Rec.InsertString("ct_posted", "Y");
                                Rec.InsertString("ct_remarks", "");
                                Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                                Rec.InsertString("rec_company_code", Record["rec_company_code"].ToString());
                                Rec.InsertString("rec_branch_code", Record["rec_branch_code"].ToString());
                                sql = Rec.UpdateRow();
                                Con_Oracle.ExecuteNonQuery(sql);
                            }
                        }
                        else if (Lib.getCCType(Row["jv_acc_main_code"].ToString()) == "HBL WISE")
                        {
                            foreach (DataRow Dr in Dt_cc_cntr.Rows)
                            {

                                bCntrExists = true;

                                iCtr++;

                                cc_amt = Lib.Conv2Decimal(Row["jv_total"].ToString()) / Lib.Conv2Decimal(Dr["tot"].ToString());
                                cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);

                                Rec = new DBRecord();

                                Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                                Rec.InsertString("ct_jvh_id", jvh_pkid);
                                Rec.InsertNumeric("ct_year", Record["jvh_year"].ToString());
                                Rec.InsertString("ct_jv_id", Row["jv_pkid"].ToString());
                                Rec.InsertString("ct_acc_id", Row["jv_acc_id"].ToString());
                                Rec.InsertString("ct_category", Dr["cc_category"].ToString());
                                Rec.InsertString("ct_cost_id", Dr["id"].ToString());
                                Rec.InsertNumeric("ct_cost_year", Record["jvh_year"].ToString());
                                Rec.InsertNumeric("ct_amount", cc_amt.ToString());
                                Rec.InsertString("ct_type", "M");
                                Rec.InsertString("ct_posted", "N");
                                Rec.InsertString("ct_remarks", "");
                                Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                                Rec.InsertString("rec_company_code", Record["rec_company_code"].ToString());
                                Rec.InsertString("rec_branch_code", Record["rec_branch_code"].ToString());
                                sql = Rec.UpdateRow();
                                Con_Oracle.ExecuteNonQuery(sql);
                            }
                            foreach (DataRow Dr in Dt_cc_hbls.Rows)
                            {
                                iCtr++;
                                cc_amt = Lib.Conv2Decimal(Row["jv_total"].ToString()) / Lib.Conv2Decimal(Dr["tot"].ToString());
                                cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                                cc_amt = getCCAmt(Dr, Lib.Conv2Decimal(Row["jv_total"].ToString()), cc_amt);

                                Rec = new DBRecord();

                                Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                                Rec.InsertString("ct_jvh_id", jvh_pkid);

                                Rec.InsertNumeric("ct_year", Record["jvh_year"].ToString());
                                Rec.InsertString("ct_jv_id", Row["jv_pkid"].ToString());
                                Rec.InsertString("ct_acc_id", Row["jv_acc_id"].ToString());
                                Rec.InsertString("ct_category", Dr["cc_category"].ToString());
                                Rec.InsertString("ct_cost_id", Dr["id"].ToString());
                                Rec.InsertNumeric("ct_cost_year", Record["jvh_year"].ToString());
                                Rec.InsertNumeric("ct_amount", cc_amt.ToString());
                                if (bCntrExists)
                                    Rec.InsertString("ct_type", "S");
                                else
                                    Rec.InsertString("ct_type", "M");

                                Rec.InsertString("ct_posted", "Y");
                                Rec.InsertString("ct_remarks", "");
                                Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                                Rec.InsertString("rec_company_code", Record["rec_company_code"].ToString());
                                Rec.InsertString("rec_branch_code", Record["rec_branch_code"].ToString());

                                sql = Rec.UpdateRow();
                                Con_Oracle.ExecuteNonQuery(sql);
                            }
                        }
                    }
                }

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                    Con_Oracle.CreateErrorLog("CCUpdateError " + Ex.Message.ToString());
                }
            }
        }


        public static decimal getCCAmt(DataRow Dr, decimal jv_total, decimal cc_amt)
        {

            try
            {
                if (Dr["ccgroup"].ToString() == "MAWB AIR EXPORT" || Dr["ccgroup"].ToString() == "MAWB AIR IMPORT")
                {
                    // DDP shipment of agent invoice is allocated to only one house
                    if (Dr["costing_ddp"].ToString() == "Y")
                    {
                        if (Lib.Conv2Decimal(Dr["tot_ddp"].ToString()) > 0)
                        {
                            cc_amt = jv_total / Lib.Conv2Decimal(Dr["tot_ddp"].ToString());
                            cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                        }
                    }
                    if (Lib.Conv2Decimal(Dr["tot_chwt"].ToString()) > 0)
                    {
                        cc_amt = jv_total * Lib.Conv2Decimal(Dr["hbl_chwt"].ToString()) / Lib.Conv2Decimal(Dr["tot_chwt"].ToString());
                        cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                    }
                }
                if (Dr["ccgroup"].ToString() == "MBL SEA EXPORT")
                {
                    // DDP shipment of agent invoice is allocated to only one house
                    if (Dr["costing_ddp"].ToString() == "Y")
                    {
                        if (Lib.Conv2Decimal(Dr["tot_ddp"].ToString()) > 0)
                        {
                            cc_amt = 0;
                            if (Dr["hbl_ddp"].ToString() == "YES")
                            {
                                cc_amt = jv_total / Lib.Conv2Decimal(Dr["tot_ddp"].ToString());
                                cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                            }
                        }
                    }
                    else if (Dr["hbl_shipment_type"].ToString() == "BUYERS CONSOLE" || Dr["hbl_shipment_type"].ToString() == "CONSOLE")
                    {
                        if (Lib.Conv2Decimal(Dr["tot_cbm"].ToString()) > 0)
                        {
                            cc_amt = jv_total * Lib.Conv2Decimal(Dr["hbl_cbm"].ToString()) / Lib.Conv2Decimal(Dr["tot_cbm"].ToString());
                            cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CreateErrorLog(ex.Message.ToString());
            }
            return cc_amt;
        }

        public static object nvlDate(object svalue, object sret, string format)
        {
            DateTime Dt;
            if (svalue == null)
                return "";
            else if (svalue.ToString() == "")
                return "";
            else
            {
                Dt = DateTime.ParseExact(svalue.ToString(), format, System.Globalization.CultureInfo.InvariantCulture);
                return Dt;
            }
        }

        public static void FindLoPAmount(string EMP_ID, int Year, int Month, decimal LpDays, decimal DaysWorked)
        {
            DBConnection Con_Oracle = null;
            bool bESIGovShare = false;
            string sql = "";
            DataTable Dt_SalMonth;
            DataTable Dt_SalMaster;
            DataRow Dr = null;
            int DaysInMonth = 30;
            decimal nPF_Amt = 0, nPF_BaseAmt = 0, TempAmt = 0, PF_Special_LimitAmt = 0;
            bool ESIAllowed = false;
            decimal ESILimit = 15000;
            decimal pf_excluded_Amt = 0;
            string pf_excluded_Cols = "";
            bool bTrans = false;
            decimal PF_percent = 12;
            decimal PF_CeilLimit = 15000;
            decimal ESI_percent = 1;
            decimal PF_Emplr_Pension_Per = 0;
            decimal PensionAmt = 0;
            decimal TotEarn_master = 0;
            decimal nTot = 0, AdminPercent = 0, EdliPercent = 0;
            try
            {
                TotEarn = 0;
                TotDeduct = 0;
                TotLopAmt = 0;

                DaysInMonth = DateTime.DaysInMonth(Year, Month);
                if (LpDays <= 0)
                    return;

                Con_Oracle = new DBConnection();

                if (LpDays > 0 && Month > 0 && Year > 0)
                {
                    sql = " select * from salarym ";
                    sql += " where SAL_EMP_ID='" + EMP_ID + "'";
                    sql += " and SAL_MONTH=" + Month;
                    sql += " and SAL_YEAR=" + Year;
                    Dt_SalMonth = new DataTable();
                    Dt_SalMonth = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_SalMonth.Rows.Count > 0)
                    {
                        PF_Emplr_Pension_Per = Lib.Convert2Decimal(Dt_SalMonth.Rows[0]["SAL_PF_EMPLR_PENSION_PER"].ToString());
                        PF_Special_LimitAmt = Lib.Convert2Decimal(Dt_SalMonth.Rows[0]["SAL_PF_LIMIT"].ToString());
                        PF_percent = Lib.Convert2Decimal(Dt_SalMonth.Rows[0]["SAL_PF_PER"].ToString());
                        PF_CeilLimit = Lib.Convert2Decimal(Dt_SalMonth.Rows[0]["SAL_PF_CEL_LIMIT"].ToString());
                        ESI_percent = Lib.Convert2Decimal(Dt_SalMonth.Rows[0]["SAL_ESI_EMPLY_PER"].ToString());
                        ESILimit = Lib.Convert2Decimal(Dt_SalMonth.Rows[0]["SAL_ESI_LIMIT"].ToString());
                        if (Dt_SalMonth.Rows[0]["SAL_IS_ESI"].ToString() == "Y")
                            ESIAllowed = true;
                        else
                            ESIAllowed = false;

                        if (ESILimit == 0 && ESIAllowed)
                            return;

                        sql = " select * from salarym ";
                        sql += " where SAL_EMP_ID = '" + EMP_ID + "'";
                        sql += " and SAL_MONTH = 0";
                        sql += " and SAL_YEAR = 0";
                        Dt_SalMaster = new DataTable();
                        Dt_SalMaster = Con_Oracle.ExecuteQuery(sql);
                        if (Dt_SalMaster.Rows.Count > 0)
                        {
                            Dr = Dt_SalMaster.Rows[0];
                            bESIGovShare = IsESIGovShare(Lib.Convert2Decimal(Dr["SAL_GROSS_EARN"].ToString()), Lib.Convert2Decimal(Dr["D02"].ToString()), Month, Year);
                            if (DaysWorked == 0)
                                LpDays = DaysInMonth;

                            TotEarn = 0;
                            TotDeduct = 0;
                            TotLopAmt = 0;
                            pf_excluded_Amt = 0;

                            sql = " select ps_pf_col_excluded ";
                            sql += " from payroll_setting a ";
                            sql += " where a.rec_company_code = '" + Dr["REC_COMPANY_CODE"].ToString() + "'";
                            sql += " and a.rec_branch_code = '" + Dr["REC_BRANCH_CODE"].ToString() + "'";
                            DataTable Dt_PS = new DataTable();
                            Dt_PS = Con_Oracle.ExecuteQuery(sql);
                            if (Dt_PS.Rows.Count > 0)
                                pf_excluded_Cols = Dt_PS.Rows[0]["ps_pf_col_excluded"].ToString();

                            sql = " Update Salarym set ";
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A01"].ToString()), LpDays, DaysInMonth);
                            sql += " A01=" + TempAmt; if (pf_excluded_Cols.Contains("A01")) pf_excluded_Amt += TempAmt;

                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A02"].ToString()), LpDays, DaysInMonth);
                            sql += ",A02=" + TempAmt; if (pf_excluded_Cols.Contains("A02")) pf_excluded_Amt += TempAmt;

                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A03"].ToString()), LpDays, DaysInMonth);
                            sql += ",A03=" + TempAmt; if (pf_excluded_Cols.Contains("A03")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A04"].ToString()), LpDays, DaysInMonth);
                            sql += ",A04=" + TempAmt; if (pf_excluded_Cols.Contains("A04")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A05"].ToString()), LpDays, DaysInMonth);
                            sql += ",A05=" + TempAmt; if (pf_excluded_Cols.Contains("A05")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A06"].ToString()), LpDays, DaysInMonth);
                            sql += ",A06=" + TempAmt; if (pf_excluded_Cols.Contains("A06")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A07"].ToString()), LpDays, DaysInMonth);
                            sql += ",A07=" + TempAmt; if (pf_excluded_Cols.Contains("A07")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A08"].ToString()), LpDays, DaysInMonth);
                            sql += ",A08=" + TempAmt; if (pf_excluded_Cols.Contains("A08")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A09"].ToString()), LpDays, DaysInMonth);
                            sql += ",A09=" + TempAmt; if (pf_excluded_Cols.Contains("A09")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A10"].ToString()), LpDays, DaysInMonth);
                            sql += ",A10=" + TempAmt; if (pf_excluded_Cols.Contains("A10")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A11"].ToString()), LpDays, DaysInMonth);
                            sql += ",A11=" + TempAmt; if (pf_excluded_Cols.Contains("A11")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A12"].ToString()), LpDays, DaysInMonth);
                            sql += ",A12=" + TempAmt; if (pf_excluded_Cols.Contains("A12")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A13"].ToString()), LpDays, DaysInMonth);
                            sql += ",A13=" + TempAmt; if (pf_excluded_Cols.Contains("A13")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A14"].ToString()), LpDays, DaysInMonth);
                            sql += ",A14=" + TempAmt; if (pf_excluded_Cols.Contains("A14")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A15"].ToString()), LpDays, DaysInMonth);
                            sql += ",A15=" + TempAmt; if (pf_excluded_Cols.Contains("A15")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A16"].ToString()), LpDays, DaysInMonth);
                            sql += ",A16=" + TempAmt; if (pf_excluded_Cols.Contains("A16")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A17"].ToString()), LpDays, DaysInMonth);
                            sql += ",A17=" + TempAmt; if (pf_excluded_Cols.Contains("A17")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A18"].ToString()), LpDays, DaysInMonth);
                            sql += ",A18=" + TempAmt; if (pf_excluded_Cols.Contains("A18")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A19"].ToString()), LpDays, DaysInMonth);
                            sql += ",A19=" + TempAmt; if (pf_excluded_Cols.Contains("A19")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A20"].ToString()), LpDays, DaysInMonth);
                            sql += ",A20=" + TempAmt; if (pf_excluded_Cols.Contains("A20")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A21"].ToString()), LpDays, DaysInMonth);
                            sql += ",A21=" + TempAmt; if (pf_excluded_Cols.Contains("A21")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A22"].ToString()), LpDays, DaysInMonth);
                            sql += ",A22=" + TempAmt; if (pf_excluded_Cols.Contains("A22")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A23"].ToString()), LpDays, DaysInMonth);
                            sql += ",A23=" + TempAmt; if (pf_excluded_Cols.Contains("A23")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A24"].ToString()), LpDays, DaysInMonth);
                            sql += ",A24=" + TempAmt; if (pf_excluded_Cols.Contains("A24")) pf_excluded_Amt += TempAmt;
                            TempAmt = FindCalcAmt(Lib.Convert2Decimal(Dr["A25"].ToString()), LpDays, DaysInMonth);
                            sql += ",A25=" + TempAmt; if (pf_excluded_Cols.Contains("A25")) pf_excluded_Amt += TempAmt;

                            decimal SalMas_pf_excluded_Amt = 0;
                            if (pf_excluded_Cols.Trim().Length > 0)
                            {
                                string[] sdata = pf_excluded_Cols.Split(',');
                                foreach (string scol in sdata)
                                    if (scol.Trim().Length > 0)
                                    {
                                        SalMas_pf_excluded_Amt += Lib.Conv2Decimal(Dr[scol.Trim()].ToString());
                                    }
                            }

                            if (PF_Special_LimitAmt > 0)
                                nPF_BaseAmt = FindCalcPFAmt(PF_Special_LimitAmt, LpDays, DaysInMonth);
                            else if ((Lib.Convert2Decimal(Dr["SAL_GROSS_EARN"].ToString()) - SalMas_pf_excluded_Amt) > PF_CeilLimit) //HRA (A04)
                                nPF_BaseAmt = FindCalcPFAmt(PF_CeilLimit, LpDays, DaysInMonth);
                            else
                                nPF_BaseAmt = (TotEarn - pf_excluded_Amt) > PF_CeilLimit ? PF_CeilLimit : (TotEarn - pf_excluded_Amt);

                            nPF_Amt = nPF_BaseAmt * (PF_percent / 100);
                            nPF_Amt = Math.Round(nPF_Amt);

                            sql += ",D01=" + Lib.Convert2Decimal(Lib.NumericFormat(nPF_Amt.ToString(), 2));
                            TotDeduct = nPF_Amt;

                            TotEarn_master = Lib.Convert2Decimal(Dr["SAL_GROSS_EARN"].ToString());//gross earn from master
                            //during LOP totearn may < 21000 whose masterearn>21000 so sal_gross taken from master
                            if (Lib.Convert2Decimal(Dr["SAL_GROSS_EARN"].ToString()) <= ESILimit || ESIAllowed)//SalaryMaster Tot Eranings 
                                TempAmt = Math.Ceiling((TotEarn * (ESI_percent / 100)));
                            else
                                TempAmt = 0;
                            if (bESIGovShare && TempAmt > 0)//if ESI Exist
                            {
                                sql += ",sal_esi_gov_share=" + Lib.Convert2Decimal(Lib.NumericFormat(TempAmt.ToString(), 2));
                                TempAmt = 0;
                            }
                            else
                                sql += ",sal_esi_gov_share = 0";

                            sql += ",D02=" + Lib.Convert2Decimal(Lib.NumericFormat(TempAmt.ToString(), 2));
                            TotDeduct += TempAmt;

                            Dr = Dt_SalMonth.Rows[0];//deduction from actual payroll
                            TotDeduct += Lib.Convert2Decimal(Dr["D03"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D04"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D05"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D06"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D07"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D08"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D09"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D10"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D11"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D12"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D13"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D14"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D15"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D16"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D17"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D18"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D19"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D20"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D21"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D22"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D23"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D24"].ToString());
                            TotDeduct += Lib.Convert2Decimal(Dr["D25"].ToString());

                            TotLopAmt = Math.Round(TotLopAmt);
                            TotEarn = Math.Round(TotEarn);
                            TotDeduct = Math.Round(TotDeduct);
                            PensionAmt = nPF_BaseAmt * (PF_Emplr_Pension_Per / 100);
                            PensionAmt = Math.Round(PensionAmt);

                            sql += ",SAL_PF_BASE =" + nPF_BaseAmt;
                            sql += ",SAL_PF_EPS_AMT =" + nPF_BaseAmt;
                            sql += ",SAL_PF_EMPLR =" + nPF_Amt;//Same as Employee
                            sql += ",SAL_PF_EMPLR_PENSION =" + PensionAmt;
                            sql += ",SAL_PF_EMPLR_SHARE =" + (nPF_Amt - PensionAmt);
                            sql += ",SAL_DAYS_WORKED =" + DaysWorked;
                            sql += ",SAL_LOP_AMT =" + TotLopAmt;
                            sql += ",SAL_GROSS_EARN =" + TotEarn;
                            sql += ",SAL_GROSS_DEDUCT =" + TotDeduct;
                            sql += ",SAL_NET =" + (TotEarn - TotDeduct);

                            AdminPercent = Lib.Convert2Decimal(Dr["SAL_ADMIN_PER"].ToString());
                            nTot = nPF_BaseAmt * (AdminPercent / 100);
                            sql += " ,SAL_ADMIN_EMPLY = " + Lib.NumericFormat(nTot.ToString(), 2);
                            if (Dr["SAL_ADMIN_BASED_ON"].ToString() == "E")
                                sql += " ,SAL_ADMIN_AMT = " + Lib.NumericFormat(nTot.ToString(), 2);

                            EdliPercent = Lib.Convert2Decimal(Dr["SAL_EDLI_PER"].ToString());
                            nTot = nPF_BaseAmt * (EdliPercent / 100);
                            sql += " ,SAL_EDLI_EMPLY = " + Lib.NumericFormat(nTot.ToString(), 2);
                            if (Dr["SAL_EDLI_BASED_ON"].ToString() == "E")
                                sql += " ,SAL_EDLI_AMT = " + Lib.NumericFormat(nTot.ToString(), 2);

                            sql += " where SAL_EMP_ID='" + EMP_ID + "'";
                            sql += " and SAL_MONTH=" + Month;
                            sql += " and SAL_YEAR=" + Year;

                            Con_Oracle.BeginTransaction();
                            bTrans = true;
                            Con_Oracle.ExecuteNonQuery(sql);
                            Con_Oracle.CommitTransaction();
                            UpdateTolerance(EMP_ID, Year, Month, TotEarn_master, TotEarn, TotDeduct, LpDays, DaysInMonth);
                        }
                    }

                }
                Con_Oracle.CloseConnection();
            }
            catch (Exception e)
            {
                if (Con_Oracle != null)
                {
                    if (bTrans)
                        Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                Con_Oracle.CreateErrorLog("EmployeeLOPUpdate" + e.Message.ToString());
                //throw (e);
            }
        }
        private static void UpdateTolerance(string EMP_ID, int Year, int Month, decimal sal_gross_master, decimal sal_gross_payroll, decimal sal_Deduct_payroll, decimal Lpdays, int DaysInMonth)
        {
            string sql = "";
            DBConnection Con_Oracle = new DBConnection();
            try
            {
                if (sal_gross_payroll > 0)
                {
                    if (sal_gross_master != sal_gross_payroll)
                    {
                        decimal NewGross = 0;
                        NewGross = Math.Round(sal_gross_master - (sal_gross_master * Lpdays / DaysInMonth));

                        if (NewGross != sal_gross_payroll)
                        {
                            decimal sal_tolerance = NewGross - sal_gross_payroll;
                            string ColFldName = "", ColFldVal = "";
                            sql = "select A25,A24,A23,A22,A21,A20,A19,A18,A17,A16,A15,A14,A13,A12,A11,A10";
                            sql += ",A09,A08,A07,A06,A05,A04,A03,A02,A01 from salarym ";
                            sql += " where sal_emp_id='" + EMP_ID + "'";
                            sql += " and sal_month=" + Month;
                            sql += " and sal_year=" + Year;
                            DataTable Dt_temp = new DataTable();
                            Dt_temp = Con_Oracle.ExecuteQuery(sql);
                            foreach (DataRow dr in Dt_temp.Rows)
                            {
                                foreach (DataColumn dCol in Dt_temp.Columns)
                                {
                                    if (Lib.Convert2Decimal(dr[dCol.ColumnName].ToString()) > 0)
                                    {
                                        ColFldName = dCol.ColumnName;
                                        ColFldVal = dr[dCol.ColumnName].ToString();
                                        break;
                                    }
                                }
                            }

                            sal_gross_payroll += sal_tolerance;
                            decimal sal_gross_net = sal_gross_payroll - sal_Deduct_payroll;
                            decimal sal_gross_lop = sal_gross_master - sal_gross_payroll;

                            sql = " Update Salarym set sal_gross_earn = " + sal_gross_payroll.ToString();
                            sql += " ,sal_net =" + sal_gross_net.ToString();
                            sql += " ,sal_lop_amt = " + sal_gross_lop.ToString();
                            sql += " ," + ColFldName + " = " + (Lib.Convert2Decimal(ColFldVal) + sal_tolerance).ToString();
                            sql += " where sal_emp_id='" + EMP_ID + "'";
                            sql += " and sal_month=" + Month;
                            sql += " and sal_year=" + Year;
                            Con_Oracle.BeginTransaction();
                            Con_Oracle.ExecuteNonQuery(sql);
                            Con_Oracle.CommitTransaction();
                        }
                    }
                }
                Con_Oracle.CloseConnection();
            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                Con_Oracle.CreateErrorLog("EmployeeToleranceUpdate" + ex.Message.ToString());
            }
        }

        public static string getProcessNumber(string company_code, string tablename, string prefix)
        {
            string sRet = "";
            int iCount = 0;
            DBConnection Con_Oracle = new DBConnection();
            try
            {
                string sql = "";
                Con_Oracle.BeginTransaction();
                sql = "update allnum set table_prefix ='" + prefix + "', table_value = 100 ";
                sql += " where table_company_code ='" + company_code + "' and table_name = '" + tablename + "' and table_prefix <> '" + prefix + "'";
                iCount = Con_Oracle.ExecuteNonQuery(sql);
                if (iCount <= 0)
                {
                    sql = "update allnum set table_value = nvl(table_value,0) + 1 ";
                    sql += " where table_company_code ='" + company_code + "' and table_name = '" + tablename + "' and table_prefix = '" + prefix + "'";
                    iCount = Con_Oracle.ExecuteNonQuery(sql);
                    if (iCount <= 0)
                    {
                        sql = "insert into allnum(table_pkid, table_company_code, table_name, table_prefix, table_value)";
                        sql += " values('{PKID}', '{COMPANY}', '{TABLENAME}', '{PREFIX}', 100) ";
                        sql = sql.Replace("{PKID}", System.Guid.NewGuid().ToString().ToUpper());
                        sql = sql.Replace("{COMPANY}", company_code);
                        sql = sql.Replace("{TABLENAME}", tablename);
                        sql = sql.Replace("{PREFIX}", prefix);
                        iCount = Con_Oracle.ExecuteNonQuery(sql);
                    }
                }
                sql = "select table_value from allnum where table_company_code ='" + company_code + "' and table_name = '" + tablename + "' and table_prefix = '" + prefix + "'";
                Object sVal = Con_Oracle.ExecuteScalar(sql);
                sRet = sVal.ToString();
                Con_Oracle.CommitTransaction();
            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                Con_Oracle.CreateErrorLog("Process Id Error" + ex.Message.ToString());
            }
            Con_Oracle.CloseConnection();
            return sRet;
        }

        private static decimal FindCalcAmt(decimal Amount, decimal LPDays, int DaysInMonth)// = 30
        {
            decimal LopAmt = 0;
            LopAmt = Amount * LPDays / DaysInMonth;
            TotLopAmt += Math.Round(LopAmt);
            Amount = Amount - LopAmt;
            Amount = Math.Round(Amount);
            TotEarn += Amount;
            return Amount;
        }
        private static decimal FindCalcPFAmt(decimal Amount, decimal LPDays, int DaysInMonth)
        {
            decimal LopAmt = 0;
            LopAmt = Amount * LPDays / DaysInMonth;
            Amount = Amount - LopAmt;
            Amount = Math.Round(Amount);
            return Amount;

        }
        private static bool IsESIGovShare(decimal grossearnings, decimal ESIAmt, int mm, int yyyy)
        {
            bool bOk = false;
            if (ESIAmt > 0)
            {
                int DaysInMonth = DateTime.DaysInMonth(yyyy, mm);
                decimal PerDaySalary = grossearnings / DaysInMonth;
                if (PerDaySalary < 140)
                    bOk = true;
            }
            return bOk;
        }


        public static string uploadFiles2Ftp(string server, string user, string pwd, string fileNames, string remotefolder = "", string ftp_agent = "")
        {
            string RetValue = "";
            try
            {
                string[] str1 = fileNames.Split('|');
                //Ftp client = new Ftp("ftp://54.39.204.134", "cargomar@transferedi.com", "cargomar2019");
                Ftp client = new Ftp(server, user, pwd);
                if (ftp_agent == "MOTHERLINES-US")
                    client.FTP_UPLOAD_PASSIVE = true;
                foreach (string str in str1)
                {
                    //client.upload(str, "/Export/" + System.IO.Path.GetFileName(str));
                    client.upload(str, remotefolder + System.IO.Path.GetFileName(str));
                }
                RetValue = "Complete";
            }
            catch (Exception Ex)
            {
                RetValue = Ex.Message.ToString();
                Lib.CreateErrorLog("FTPERROR-" + Ex.Message.ToString());
            }
            return RetValue;
        }

        public static string DownloadFilesFromFtp(string server, string user, string pwd, string downloadfilePath, string remotefolder, ref string downloadcount)
        {
            string RetValue = "";
            string RemoteFile = "";
            string LocalFile = "";
            string sRemarks = "";
            string logFileName = "";
            int filesDwnldCount = 0;
            try
            {
                //downloadfilePath = downloadfilePath.Replace("\\", "//");

                if (!Directory.Exists(downloadfilePath))
                    Directory.CreateDirectory(downloadfilePath);

                DateTime Dt = DateTime.Now;
                sRemarks = Dt.ToString("yyyy-MM-dd:HH:mm:ss tt");
                logFileName = downloadfilePath + "\\PROCESSED\\log.txt";

                string[] str1 = GetFtpDirFiles(server, user, pwd, remotefolder);

                Ftp client = new Ftp(server, user, pwd);

                foreach (string str in str1)
                {
                    if (str.ToString().ToUpper().EndsWith(".CSV"))
                    {
                        RemoteFile = remotefolder + str;
                        LocalFile = downloadfilePath + "\\" + str;
                        CreateErrorLog(sRemarks + "," + RemoteFile + "," + LocalFile, logFileName);
                        if (client.download(RemoteFile, LocalFile) == "")
                        {
                            filesDwnldCount++;
                            if (System.IO.File.Exists(LocalFile))
                            {
                                client.delete(RemoteFile);
                            }
                        }
                    }
                }

                downloadcount = filesDwnldCount.ToString();
                RetValue = "Complete";
            }
            catch (Exception Ex)
            {
                RetValue = Ex.Message.ToString();
                CreateErrorLog(sRemarks + "," + RetValue, logFileName);
            }
            return RetValue;
        }

        private static string[] GetFtpDirFiles(string server, string user, string pwd, string remotefolder)
        {
            Ftp client = new Ftp(server, user, pwd);
            string[] str1 = client.directoryListSimple(remotefolder);

            //str1 = client.directoryListDetailed(remotefolder);

            return str1;
        }

        public static string GetCntrno(string sCntr)
        {
            sCntr = sCntr.Replace(" ", "");
            if (sCntr.Length > 12)
                sCntr = sCntr.Substring(0, 12);
            return sCntr;
        }

        /*
        private static void CreateFileLog(string sRemarks,string ProcessedPath)
        {
            try
            {
                DateTime Dt = DateTime.Now;
                string FileName = ProcessedPath + "log.txt";
                string sData = Dt.ToString("yyyy-MM-dd:HH:mm:ss tt") ;
                StreamWriter sw = new StreamWriter(FileName, true);

                if (sRemarks.Contains(","))
                {
                    string[] sArry = sRemarks.Split(',');
                    for (int i = 0; i < sArry.Length; i++)
                    {
                        sRemarks = sData + "," + sArry[i];
                        sw.WriteLine(sRemarks);
                    }
                }
                else
                {
                    sData += "," + sRemarks  ;
                    sw.WriteLine(sData);
                }
                sw.Flush();
                sw.Close();
            }
            catch (Exception) { }
        }
        */
        public static Boolean IsCC_CategoryExist(string CC_ID, ref string REFNO)
        {
            Boolean bOk = false;
            string sql = "";
            sql = "select jvh_pkid,jvh_docno from ledgerh where jvh_cc_id = '" + CC_ID + "'";

            DBConnection Con_Oracle = new DBConnection();
            Con_Oracle = new DBConnection();
            DataTable Dt_CC = new DataTable();
            Dt_CC = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            if (Dt_CC.Rows.Count > 0)
            {
                bOk = true;
                REFNO = Dt_CC.Rows[0]["jvh_docno"].ToString();
            }
            Dt_CC.Rows.Clear();
            
            return bOk;
        }


        public static Boolean IsValidGst(string brid, string GSTNO )
        {
            Boolean bOk = true;
            string sql = "";

            if ( brid !="" && GSTNO != "")
            {
                bOk = false;
                sql = "select add_gstin from addressm where add_pkid ='" + brid + "' and add_gstin = '" + GSTNO + "'";
                DBConnection Con_Oracle = new DBConnection();
                Con_Oracle = new DBConnection();
                DataTable Dt_CC = new DataTable();
                Dt_CC = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                if (Dt_CC.Rows.Count > 0)
                {
                    bOk = true;
                }
                Dt_CC.Rows.Clear();
                Con_Oracle.CloseConnection();
            }

            return bOk;
        }



        public static void GetEmployeeBranch(int Emp_Br_Grp, ref string BrName, ref string Addr1, ref string Addr2, ref string Addr3)
        {
            if (Emp_Br_Grp == 2)
            {
                BrName = "CARGOMAR PRIVATE LTD";
                Addr1 = "DOOR NO.13,NEAR VELAN UTHARA HOTEL,BUNGALOW STREET";
                Addr2 = "TIRUPUR - 641 602";
                Addr3 = "";
            }
        }


        public static bool IsValidGST(bool IsGst, string GstNo, string StateCode, Boolean Igst_Exception = false)
        {
            bool bOk = true;
            try
            {
                if (GstNo == null)
                    GstNo = "";
                if (StateCode == null)
                    StateCode = "";

                if (IsGst && GstNo.Trim().Length > 0)
                {
                    if (GstNo.Trim().Length != 15)
                        bOk = false;
                    if (bOk)
                    {
                        if (GstNo.Trim().Substring(0, 2) != StateCode.Trim())
                            bOk = false;
                        if (Igst_Exception && StateCode.Trim() != "97")
                            bOk = false;
                        if (Igst_Exception && StateCode.Trim() == "97")
                            bOk = true;

                    }
                }
            }
            catch (Exception)
            {
                bOk = true;
            }
            return bOk;
        }


        public static void InsertMappingData_EDI_ORDER(string id)
        {
            string sql = "";
            sql += " select distinct 'CUSTOMER' as catg, 'POL-AGENT' as subcatg, 'ID' as link_mode,  rec_company_code, ord_sender as sender, ord_pol_agent as source_name from edi_order where ord_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'CUSTOMER' as catg, 'SHIPPER' as subcatg, 'ID' as link_mode,  rec_company_code, ord_sender as sender, ord_exp_name as source_name from edi_order where ord_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'CUSTOMER' as catg, 'CONSIGNEE' as subcatg, 'ID' as link_mode, rec_company_code, ord_sender as sender,  ord_imp_name as source_name  from edi_order where ord_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'PORT' as catg, 'PORT' as subcatg,'ID' as link_mode,  rec_company_code, ord_sender as sender,  ord_pol as source_name from edi_order where ord_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'PORT' as catg, 'PORT' as subcatg,'ID' as link_mode,  rec_company_code, ord_sender as sender,  ord_pod as source_name  from edi_order where ord_headerid = '" + id + "' ";

            DataTable Dt_Temp = new DataTable();
            DBConnection Con_Oracle = new DBConnection();
            Dt_Temp = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            InsertInwardEdiLink("INWARD", Dt_Temp);
        }


        public static void InsertMappingData_EDI_BL(string id)
        {
            string sql = "";
            sql += " select distinct 'CUSTOMER' as catg, 'POL-AGENT' as subcatg, 'ID' as link_mode,  rec_company_code,hbl_sender as sender, hbl_pol_agent as source_name from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'CUSTOMER' as catg, 'SHIPPER' as subcatg, 'ID' as link_mode,  rec_company_code,hbl_sender as sender, hbl_shipper_name as source_name from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'CUSTOMER' as catg, 'CONSIGNEE' as subcatg, 'ID' as link_mode, rec_company_code,hbl_sender as sender,  hbl_consignee_name as source_name  from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'CUSTOMER' as catg, 'NOTIFY' as subcatg, 'ID' as link_mode, rec_company_code,hbl_sender as sender,  hbl_notify_name as source_name  from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'SEA CARRIER' as catg, 'CARRIER' as subcatg, 'ID' as link_mode, rec_company_code,hbl_sender as sender,  hbl_carrier_name as source_name  from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'PORT' as catg, 'PORT' as subcatg,'ID' as link_mode,  rec_company_code,hbl_sender as sender,  hbl_pol_name as source_name from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'PORT' as catg, 'PORT' as subcatg,'ID' as link_mode,  rec_company_code,hbl_sender as sender,  hbl_pod_name as source_name from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'PORT' as catg, 'PORT' as subcatg,'ID' as link_mode,  rec_company_code,hbl_sender as sender,  hbl_pofd_name as source_name from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'VESSEL' as catg, 'VESSEL' as subcatg,'ID' as link_mode,  a.rec_company_code,hbl_sender as sender,  vsl_name as source_name from edi_house a inner join edi_house_vessel b on a.hbl_pkid = b.vsl_hbl_id where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            //sql += " select distinct 'PORT' as catg, 'PORT' as subcatg,'ID' as link_mode,  a.rec_company_code,hbl_sender as sender,  vsl_pol_name as source_name from edi_house a inner join edi_house_vessel b on a.hbl_pkid = b.vsl_hbl_id where hbl_headerid = '" + id + "' ";
            //sql += " union all ";
            //sql += " select distinct 'PORT' as catg, 'PORT' as subcatg,'ID' as link_mode,  a.rec_company_code,hbl_sender as sender,  vsl_pod_name as source_name from edi_house a inner join edi_house_vessel b on a.hbl_pkid = b.vsl_hbl_id where hbl_headerid = '" + id + "' ";
            //sql += " union all ";
            sql += " select distinct 'CONTAINER TYPE' as catg, 'CONTAINER TYPE' as subcatg,'ID' as link_mode,  a.rec_company_code,hbl_sender as sender,  cntr_size as source_name from edi_house a inner join edi_house_container b on a.hbl_pkid = b.cntr_hbl_id where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'UNIT' as catg, 'UNIT' as subcatg,'ID' as link_mode,  a.rec_company_code,hbl_sender as sender,  cntr_pkgs_unit as source_name from edi_house a inner join edi_house_container b on a.hbl_pkid = b.cntr_hbl_id where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'BRANCH' as catg, 'BRANCH' as subcatg,'ID' as link_mode,  rec_company_code,hbl_sender as sender,  hbl_pod_name as source_name from edi_house where hbl_headerid = '" + id + "' ";
            sql += " union all ";
            sql += " select distinct 'UNIT' as catg, 'UNIT' as subcatg,'ID' as link_mode,  rec_company_code,hbl_sender as sender,  hbl_pkg_unit as source_name from edi_house where hbl_headerid = '" + id + "' ";

            DataTable Dt_Temp = new DataTable();
            DBConnection Con_Oracle = new DBConnection();
            Dt_Temp = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            InsertInwardEdiLink("INWARD", Dt_Temp);
        }

        public static void InsertInwardEdiLink(string link_type, DataTable Dt_Temp)
        {
            string id = "";
            string sql = "";

            foreach (DataRow Dr in Dt_Temp.Rows)
            {

                id = Guid.NewGuid().ToString().ToUpper();
                DBRecord mRec = new DBRecord();
                mRec.CreateRow("edi_link", "ADD", "link_pkid", id);
                mRec.InsertString("link_type", link_type);
                mRec.InsertString("rec_company_code", Dr["rec_company_code"].ToString());
                mRec.InsertString("link_messagesender", Dr["sender"].ToString());
                mRec.InsertString("link_category", Dr["catg"].ToString());
                mRec.InsertString("link_subcategory", Dr["subcatg"].ToString());
                mRec.InsertString("link_source_name", Dr["source_name"].ToString());
                mRec.InsertString("link_status", "N");
                mRec.InsertString("link_mode", Dr["link_mode"].ToString());
                try
                {
                    DBConnection Con_Oracle = new DBConnection();
                    Con_Oracle.BeginTransaction();
                    sql = mRec.UpdateRow();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    Con_Oracle.CloseConnection();
                }
                catch (Exception) { }
            }

        }

        public static string getSettings(string parentid, string caption, string fldname)
        {
            DBConnection Con_Oracle = new DBConnection();
            string str = "";
            try
            {
                string sql = "";
                sql += " select id, code, name from settings where ";
                sql += " parentid = '" + parentid + "'";
                sql += " and caption = '" + caption + "'";

                DataTable Dt_test = new DataTable();

                Dt_test = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                if (Dt_test.Rows.Count > 0)
                    return Dt_test.Rows[0][fldname].ToString();
                else
                    return "";
            }
            catch (Exception)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                str = "";
            }
            return str;
        }

        public static Boolean createThumbnailFromPdf(string source, string target, string fname)
        {
            Boolean bRet = false;
            try
            {
                

                bRet = true;
            }
            catch ( Exception Ex)
            {
                bRet = false;
            }
            return bRet;
        }


        public static string getPath(string ServerPath, string comp_code, string table_name, string slno ,Boolean createFolder)
        {
            string Folder = Path.Combine(ServerPath, comp_code, table_name, slno.ToString());
            if ( createFolder)
                Lib.CreateFolder(Folder);
            return Folder;
        }


        public static string GetSeverImageURLWithCompany(string comp_code)
        {
            string serverUrl = Lib.getSettings(comp_code, "SERVER-URL", "NAME");
            string ServerImageFolder = Lib.getSettings(comp_code, "IMAGE-FOLDER", "NAME");
            return Path.Combine(serverUrl, ServerImageFolder, comp_code);
        }

        public static string GetImagePathWithCompany(string comp_code)
        {
            string serverUrl = Lib.getSettings(comp_code, "SERVER-URL", "NAME");
            string ServerImageFolder = Lib.getSettings(comp_code, "IMAGE-FOLDER", "NAME");
            return System.Web.HttpContext.Current.Server.MapPath("~/" + ServerImageFolder + "/" + comp_code);
        }



        public static string GetSeverImageURL(string comp_code)
        {
            string serverUrl = Lib.getSettings(comp_code, "SERVER-URL", "NAME");
            string ServerImageFolder = Lib.getSettings(comp_code, "IMAGE-FOLDER", "NAME");
            return Path.Combine(serverUrl, ServerImageFolder);
        }

        public  static string GetImagePath(string comp_code)
        {
            string serverUrl = Lib.getSettings(comp_code, "SERVER-URL", "NAME");
            string ServerImageFolder = Lib.getSettings(comp_code, "IMAGE-FOLDER", "NAME");
            return System.Web.HttpContext.Current.Server.MapPath("~/" + ServerImageFolder);
        }

        public static string GetReportPath(string comp_code)
        {
            string serverUrl = Lib.getSettings(comp_code, "SERVER-URL", "NAME");
            string ServerReportFolder = Lib.getSettings(comp_code, "REPORTS-FOLDER", "NAME");
            return System.Web.HttpContext.Current.Server.MapPath("~/" + ServerReportFolder);
        }

        public static string getUploadedPath(string comp_code, string mainid = "", string subid = "", Boolean isUrl = true, string table_name = "DOCS")
        {
            string ServerImageUrl = "";
            if ( isUrl )
                ServerImageUrl = Lib.GetSeverImageURL(comp_code);
            else 
                ServerImageUrl = Lib.GetImagePath(comp_code);
            string Folder = Path.Combine(ServerImageUrl, comp_code, table_name, mainid, subid);
            return Folder;
        }


        public static Boolean RemoveFileFolder(string comp_code, string mainid = "", string subid = "", string table_name = "DOCS")
        {
            string ServerImagePath = GetImagePath(comp_code);
            string Folder = Path.Combine(ServerImagePath, comp_code, "DOCS", mainid);
            Lib.RemoveFolder(Folder);
            return true;
        }

        public static string UploadFile(System.Web.HttpPostedFile hpf, string comp_code,string mainid ="", string subid ="", string table_name = "DOCS")
        {
            string retval = "";
            string FileName = "";

            string ServerImagePath = GetImagePath(comp_code);
            string Folder = Path.Combine(ServerImagePath, comp_code, "DOCS", mainid, subid);

            try
            {
                Lib.CreateFolder(Folder);
                FileName = Path.Combine(Folder, Path.GetFileName(hpf.FileName));
                hpf.SaveAs(FileName);
            }
            catch (Exception Ex)
            {
                retval = Ex.Message.ToString();
            }
            return retval;
        }


        public static string getEmail(string e1, string e2)
        {
            string str = e1;
            if (e2.ToString().Length > 0)
            {
                if (str.Length > 0)
                    str += ",";
                str += e2;
            }
            return str;
        }





    }
}
