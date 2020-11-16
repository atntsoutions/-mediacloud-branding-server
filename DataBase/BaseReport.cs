using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing.Printing;

namespace DataBase
{
    public class BaseReport
    {
        public Boolean CanWrite = true;
        private string sChar = "\b";
        public List<string> ExportList = null;

        public Single ROW_HT = 15;

        public Single HCOL1 = 10;
        public Single HCOL2 = 400;
        public Single HCOL3 = 600;
        public Single HCOL4 = 800;


        public Single HCOL5 = 0;
        public Single HCOL6 = 0;
        public Single HCOL7 = 0;
        public Single HCOL8 = 0;
        public Single HCOL9 = 0;
        public Single HCOL10 = 0;

        public Single HCOL_START = 0;

        public Single HCOL_MAX_WIDTH = 595;
        public Single HCOL_MAX_HEIGHT = 842;

        // A4 - W 595 H 842 
        public object RootFrame = null;

        public Single Row = 10;
        public int ifontSize = 10;
        public int Page_Height = 1024;
        public int Page_Width = 800;
        public int Total_Pages = 0;

        public string ifontName = "Times New Roman";


        public int CurrentPageNumber = 0;

        public BaseReport()
        {
        }

        public void AddPage(int iHeight, int iWidth)
        {
            addList("{ADDPAGE}", iHeight.ToString(), iWidth.ToString());
            Total_Pages++;
        }

        public void BeginReport(int iHt, int iWd)
        {
            Page_Height = iHt;
            Page_Width = iWd;
        }

        public void EndReport()
        {
            addList("{END}");
        }
        public void AddTextFile()
        {
            addList("{TEXTFILE}");
        }
        public void addExportList(string str)
        {
            ExportList.Add(str);
        }

        public void DrawHLine(object rootframe, Single Left, Single Top, Single Width, string ColorCode = "0")
        {
            addList("{HLINE1}", Left.ToString(), Top.ToString(), Width.ToString(), ColorCode);
        }

        public void DrawHLine(Single Left, Single Top, Single Width, string ColorCode = "0")
        {
            addList("{HLINE1}", Left.ToString(), Top.ToString(), Width.ToString(), ColorCode);
        }

        public void DrawXLHBorder(int iRow, int StartCol, int EndCol, int iHt, string style)
        {
            addList("{HBORDER}", iRow.ToString(), StartCol.ToString(), EndCol.ToString(), iHt.ToString(), style);
        }
        public void DrawXLVBorder(int StartRow, int EndRow, int iCol, string style)
        {
            addList("{VBORDER}", StartRow.ToString(), EndRow.ToString(), iCol.ToString(), style);
        }

        public void DrawVLine(object rootframe, Single Left, Single Top, Single Height, string ColorCode = "0")
        {
            addList("{VLINE1}", Left.ToString(), Top.ToString(), Height.ToString(), ColorCode);
        }
        public void DrawVLine(Single Left, Single Top, Single Height, string ColorCode = "0")
        {
            addList("{VLINE1}", Left.ToString(), Top.ToString(), Height.ToString(), ColorCode);
        }

        //public void AddXYLabel(Single x, Single y, Single RowHeight, Single RowWidth, string Text1, string iFontName, int iFontSize, string sBorder, string sStyle, int XLRow = 0, int XLCol = 0, int ColSpan = 0, int RowHT = 16, int RowSpan = 0, Single Pdfy = 0, Single Xtolerance = 0, int PdfFontSize = 0)
        //{
           // AddXYLabel(x, y, RowHeight, RowWidth, Text1, iFontName, iFontSize, sBorder, sStyle, XLRow, XLCol, ColSpan, RowHT, RowSpan, Pdfy, Xtolerance, PdfFontSize);
        //}

        public void AddXYLabel(Single x, Single y, Single RowHeight, Single RowWidth, string Text1, string iFontName, int iFontSize, string sBorder, string sStyle, string fillColor = "")
        {

            addList("{XYLABEL}", x.ToString(), y.ToString(), RowHeight.ToString(), RowWidth.ToString(), Text1, iFontName.ToString(), iFontSize.ToString(), sBorder, sStyle, fillColor );

    /*
            string ClrCode = "0";
            if (sBorder.IndexOf("0") >= 0)
                ClrCode = "0";
            if (sBorder.IndexOf("1") >= 0)
                ClrCode = "1";
            if (sBorder.IndexOf("2") >= 0)
                ClrCode = "2";

            if (sBorder.IndexOf("L") >= 0)
                DrawVLine(x, y, RowHeight, ClrCode);
            if (sBorder.IndexOf("T") >= 0)
                DrawHLine(x, y, RowWidth, ClrCode);
            if (sBorder.IndexOf("B") >= 0)
                DrawHLine(x, y + RowHeight, RowWidth, ClrCode);
            if (sBorder.IndexOf("R") >= 0)
                DrawVLine(x + RowWidth, y, RowHeight, ClrCode);
            if (sBorder.IndexOf("l") >= 0)
                DrawVLine(x, y, RowHeight, RowWidth, -5, ClrCode);
            if (sBorder.IndexOf("r") >= 0)
                DrawVLine(x, y, RowHeight, RowWidth, 5, ClrCode);
                */
        }

        public void addList(params string[] slist)
        {
            string str = "";

            if (CanWrite == false)
                return;
            for (int i = 0; i < slist.Length; i++)
            {
                str += (str != "") ? sChar : "";
                str += slist[i];
            }

            // Adding PageNumbers to All Rows
            str += (str != "") ? sChar : "";
            str += (CurrentPageNumber == 0) ? 1 : CurrentPageNumber;

            if (ExportList == null)
            {
                ExportList = new List<string>();
            }
            ExportList.Add(str);
        }


        public void DrawPolygon(string fillcolor, int totalpoints, string points)
        {
            addList("{POLYGON}",fillcolor, totalpoints.ToString(), points);
        }


        public void DrawVLine(Single Left, Single Top, Single Height, Single Width, Single Radius, string ColorCode = "0")
        {
            addList("{VLINE2}", Left.ToString(), Top.ToString(), Height.ToString(), Width.ToString(), Radius.ToString(), ColorCode);
        }

        public void SetFillRectangle(float left, float top, float height, float width, int BorderThickness = 1, string BorderColor = "BLACK", string FillColor = "LIGHTGRAY")
        {
            addList("{BOX}", left.ToString(), top.ToString(), height.ToString(), width.ToString(), BorderThickness.ToString(), BorderColor, FillColor);
        }

        public void DrawDashLine(Single Left, Single Top, Single Width, string ColorCode = "0")
        {
            addList("{HLINE1}", Left.ToString(), Top.ToString(), Width.ToString(), ColorCode, "DOT");
        }
        public void DrawHBoldLine(Single Left, Single Top, Single Width, string ColorCode = "0")
        {
            addList("{HLINE1}", Left.ToString(), Top.ToString(), Width.ToString(), ColorCode);
        }

        public void LoadImage(object mCanvas, float Left, float Top, string ImageName, float Height = 0, float Width = 0, int PdfY = 100)
        {
            try
            {
                addList("{IMAGE1}", ImageName, Left.ToString(), Top.ToString(), Height.ToString(), Width.ToString(), PdfY.ToString());
            }
            catch (Exception)
            {
            }
        }

        public void LoadLogoctpatsm(object mCanvas, float Left, float Top)
        {
            try
            {
                LoadImage("", Left, Top, "logoctpatsm.jpg", 57, 150);
            }
            catch (Exception)
            {
            }
        }


        // END OF REPORT
    }
}
