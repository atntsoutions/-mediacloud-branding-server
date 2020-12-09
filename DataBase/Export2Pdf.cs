using System;
using System.Collections.Generic;
using System.Net;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;

using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using System.IO;

namespace DataBase
{
    public class Export2Pdf
    {
        private char sChar = '\b';

        public List<string> ExportList = null;

        public Boolean ShowFile = false;


        private PdfDocument _Doc = null;
        private PdfPage _Page = null;

        XGraphics gfx = null;

        public Boolean CanWrite = true;

        public string FolderID = "";
        public string FileName = "";

        public string Default_FileName = "";


        public int Page_Height = 1024;
        public int Page_Width = 800;

        public Boolean IsError = false;
        public string sError = "";

        public int TotPages = 1;

        public int LEFT_MARGIN = 0;


        public void Process()
        {
            try
            {
                IsError = false;
                sError = "";
                _Doc = new PdfDocument();

                foreach (string str in ExportList)
                {
                    if (str.IndexOf("{ADDPAGE}") >= 0)
                        AddPage(str);
                    else if (str.IndexOf("{XYLABEL}") >= 0)
                        AddXYLabel(str);
                    else if (str.IndexOf("{HLINE1}") >= 0)
                        AddHLine1(str);
                    else if (str.IndexOf("{VLINE1}") >= 0)
                        AddVLine1(str);
                    else if (str.IndexOf("{VLINE2}") >= 0)
                        AddVLine2(str);
                    else if (str.IndexOf("{BOX}") == 0)
                        AddBox(str);
                    else if (str.IndexOf("{POLYGON}") == 0)
                        AddPolyGon (str);
                    else if (str.IndexOf("{IMAGE1}") >= 0)
                        AddImage(str);
                    else if (str.IndexOf("{IMAGE2}") >= 0)
                        AddImage2(str);
                }

                SavePdf();
                CloseAll();
            }
            catch (Exception Ex)
            {
                IsError = true;
                sError = Ex.Message.ToString();
                // MessageBox.Show(Ex.Message.ToString(), "Export Pdf", MessageBoxButtons.OK);
            }
        }

        private void CloseAll()
        {
            if (gfx != null)
                gfx.Dispose();
            if (_Doc != null)
                _Doc.Dispose();
        }

        private void AddPage(string str)
        {

            _Page = _Doc.AddPage();

            _Page.Size = PdfSharp.PageSize.A4;

            if (Page_Height <= 0 || Page_Width <= 0)
            {
                _Page.Size = PdfSharp.PageSize.A4;
                _Page.Orientation = PdfSharp.PageOrientation.Landscape;
            }
            else
            {
                _Page.Width = XUnit.FromPoint(Page_Width);
                _Page.Height = XUnit.FromPoint(Page_Height);
            }


            gfx = XGraphics.FromPdfPage(_Page);
        }

       



        private void AddXYLabel(string str)
        {
            string[] sData = str.Split(sChar);
            int x = LEFT_MARGIN + int.Parse(sData[1].ToString());
            int y = int.Parse(sData[2].ToString());
            int RowHeight = int.Parse(sData[3].ToString());
            int RowWidth = int.Parse(sData[4].ToString());
            string Text1 = sData[5].ToString();
            string FontName = sData[6].ToString();
            int iFontSize = int.Parse(sData[7].ToString());
            string sBorder = sData[8].ToString();
            string sStyle = sData[9].ToString();
            string sFillColor = sData[10].ToString();

            string borderColor = "";

            string sColorCode = "0";//BLACK

            XFontStyle mStyle = XFontStyle.Regular;
            if (sStyle.IndexOf("B") >= 0)
                mStyle = mStyle ^ XFontStyle.Bold;
            if (sStyle.IndexOf("I") >= 0)
                mStyle = mStyle ^ XFontStyle.Italic;
            if (sStyle.IndexOf("U") >= 0)
                mStyle = mStyle ^ XFontStyle.Underline;
            if (sStyle.IndexOf("S") >= 0)
                mStyle = mStyle ^ XFontStyle.Strikeout;
            if (sStyle.IndexOf("0") >= 0)
                sColorCode = "0";
            if (sStyle.IndexOf("1") >= 0)
                sColorCode = "1";
            if (sStyle.IndexOf("2") >= 0)
                sColorCode = "2";
            if (sStyle.IndexOf("3") >= 0)
                sColorCode = "3";


            XFont mFont = new XFont(FontName, iFontSize, mStyle);

            XStringFormat mAlign = new XStringFormat();
            mAlign.LineAlignment = XLineAlignment.Center;
            mAlign.Alignment = XStringAlignment.Near;
            if (sStyle.IndexOf("C") >= 0)
                mAlign.Alignment = XStringAlignment.Center;
            if (sStyle.IndexOf("R") >= 0)
                mAlign.Alignment = XStringAlignment.Far;
            if (sStyle.IndexOf("J") >= 0)
            {
                DrawJustify(x, y, RowWidth, RowHeight, Text1, mFont, Lib.GetColorName(sColorCode));
            }
            else if (sStyle.IndexOf("W") >= 0)
            {
                DrawWaterMark(x, y, Text1, mFont, Lib.GetColor(sColorCode));
            }
            else if (sStyle.IndexOf("D") >= 0)
            {
                DrawWaterMarkDiagonal(x, y, Text1, mFont, Lib.GetColor(sColorCode));
            }

            else
            {
                if (mAlign.Alignment == XStringAlignment.Near)
                    x = x + 1;
                XRect rect = new XRect(x, y, RowWidth, RowHeight);

                if (sBorder.IndexOf("A") >= 0)
                    borderColor = "BLACK";
                if (borderColor != "" || sFillColor != "")
                {
                    DrawBox(x, y, RowHeight, RowWidth, 1, borderColor, sFillColor);
                    //DrawBorder(x, y, RowHeight, RowWidth, 1, "BLACK");
                    if ( mAlign.Alignment == XStringAlignment.Near)
                        rect = new XRect(x +1, y, RowWidth, RowHeight);
                }

                // gfx.DrawString(Text1, mFont, XBrush("BLACK"), rect, mAlign);
                gfx.DrawString(Text1, mFont, XBrush(Lib.GetColorName(sColorCode)), rect, mAlign);
            }
        }



        private void DrawWaterMark(int wx, int wy, string wText, XFont wFont, Color wColor)
        {
            XSize sf = gfx.MeasureString(wText, wFont);
            gfx.TranslateTransform(wx / 2, wy / 2);

            //double Agle = -Math.Atan(wy / wx) * 180 / Math.PI;
            //gfx.RotateTransform((float)Agle);
            //gfx.TranslateTransform(-wx / 2, -wy / 2);

            XStringFormat wformat = new XStringFormat();
            wformat.Alignment = XStringAlignment.Near;
            wformat.LineAlignment = XLineAlignment.Near;

            // XBrush wbrush = new XSolidBrush(Color.FromArgb(20, Color.Black));
            XBrush wbrush = new XSolidBrush(Color.FromArgb(20, wColor));

            wx = (int)(wx - sf.Width) / 2;
            wy = (int)(wy - sf.Height) / 2;

            XPoint Point_xy = new XPoint(wx, wy);
            //gfx.DrawString(wText, wFont, wbrush, Point_xy, wformat);


            // OUTLINED METHOD
            // System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            // StringFormat wformat2 = new StringFormat();
            // wformat2.Alignment = StringAlignment.Near;
            // wformat2.LineAlignment = StringAlignment.Near;
            Point Point_xy2 = new Point(wx, wy);
            // path.AddString(wText, new FontFamily("Arial"), (int)FontStyle.Italic, 120, Point_xy2, wformat2);

            // Pen pen = new Pen(Color.FromArgb(64,Color.Black), 3);               
            //gfx.Graphics.DrawPath(pen, path);


            XGraphicsPath path = new XGraphicsPath();
            path.AddString(wText, wFont.FontFamily, XFontStyle.Italic, 103, Point_xy2, wformat);
            Pen pen = new Pen(Color.FromArgb(64, wColor), 3);
            //e.Graphics.DrawPath(pen, path);
            gfx.DrawPath(pen, path);
        }
        private void DrawJustify(int jx, int jy, int jRowWidth, int jRowHeight, string jText, XFont jFont, string jBrushColor)
        {
            string[] sdata = null;
            int NoOfSpace = 0;
            int BalWidth = 0;

            XStringFormat SF = new XStringFormat();
            SF.LineAlignment = XLineAlignment.Center;

            sdata = jText.Trim().Split(' ');
            if (sdata.Length <= 1)
            {
                XRect rect = new XRect(jx + 1, jy, jRowWidth, jRowHeight);
                gfx.DrawString(jText, jFont, XBrush(jBrushColor), rect, SF);
            }
            else
            {
                NoOfSpace = sdata.Length - 1;
                XSize sf;
                int TotWordLength = 0;
                for (int i = 0; i < sdata.Length; i++)
                {
                    sf = gfx.MeasureString(sdata[i], jFont, SF);
                    TotWordLength += (int)sf.Width;
                }

                BalWidth = (int)jRowWidth - TotWordLength;

                int SpaceWidth = BalWidth / NoOfSpace;
                int RemSpaceWidth = BalWidth % NoOfSpace;

                int xCol = jx;
                int xCol2 = 0;

                for (int i = 0; i < sdata.Length; i++)
                {
                    sf = gfx.MeasureString(sdata[i], jFont, SF);

                    xCol2 = xCol + (int)sf.Width;
                    if (RemSpaceWidth > 0)
                    {
                        xCol2 += 1;
                        RemSpaceWidth--;
                    }

                    jx = xCol;
                    if (i == sdata.Length - 1)
                    {
                        SF.Alignment = XStringAlignment.Far;
                        jx = jx - 2;
                    }
                    else
                    {
                        xCol2 += SpaceWidth;
                        SF.Alignment = XStringAlignment.Near;
                        jx += 1;
                    }

                    RectangleF rectf_xy = new RectangleF(jx, jy, xCol2 - xCol + 4, jRowHeight);
                    gfx.DrawString(sdata[i], jFont, XBrush(jBrushColor), rectf_xy, SF);

                    xCol = xCol2;
                }
            }
        }

        private void DrawHLine(int x, int y, int Width, string lColorCode = "0", string sLineType = "")
        {
            // XPen mPen = new XPen(XColors.Black);
            XPen mPen = new XPen(XColor(Lib.GetColorName(lColorCode)));
            if (sLineType == "DOT")
                mPen.DashStyle = XDashStyle.Dot;
            gfx.DrawLine(mPen, x, y, x + Width, y);
        }
        private void DrawVLine(int Left, int Top, int Height, string lColorCode = "0")
        {
            // XPen mPen = new XPen(XColors.Black);
            XPen mPen = new XPen(XColor(Lib.GetColorName(lColorCode)));
            gfx.DrawLine(mPen, Left, Top, Left, (Top + Height));

            //gfx.DrawLine(mPen, Left, iHeight - Top, Left, iHeight - (Top + Height));
        }
        private void DrawVLine(int Left, int Top, int Height, int Width, int Radius, string lColorCode = "0")
        {
            int R = 0;
            //XPen mPen = new XPen(XColors.Black);
            XPen mPen = new XPen(XColor(Lib.GetColorName(lColorCode)));
            if (Radius < 0)
            {
                R = Math.Abs(Radius);
                gfx.DrawLine(mPen, Left - R, Top, Left, (Top + Height));
                //gfx.DrawLine(mPen, Left - R, iHeight - Top, Left, iHeight - (Top + Height));
            }
            else if (Radius > 0)
            {
                R = Radius;
                gfx.DrawLine(mPen, Left + Width + R, Top, Left + Width, (Top + Height));
                //gfx.DrawLine(mPen, Left + Width + R, iHeight - Top, Left + Width, iHeight - (Top + Height));
            }
        }

        private void DrawBox(int x, int y, int Height, int Width, int BorderThickness, string BorderColor, string FillColor)
        {
            XPen mPen = null;
            XBrush xBrush = null;
            if (BorderColor != "")
                mPen = new XPen(XColor(BorderColor));
            if (FillColor != "")
                xBrush = XBrush(FillColor);
            if (BorderColor != "" && FillColor != "")
                gfx.DrawRectangle(mPen,xBrush, x, y, Width, Height);
            else if (BorderColor != "")
                gfx.DrawRectangle(mPen, x, y, Width, Height);
            if (FillColor != "")
                gfx.DrawRectangle(xBrush, x, y, Width, Height);
        }


        private void DrawBorder(int x, int y, int Height, int Width, int BorderThickness, string BorderColor)
        {
            XPen mPen = new XPen(XColor(BorderColor));
            gfx.DrawRectangle(mPen, x, y, Width, Height);
        }


        private void AddPolyGon(string str)
        {
            XPoint[] points;
            string[] sData = str.Split(sChar);
            int iCtr = 0;

            string fillColor = sData[1].ToString();
            int totpoints = int.Parse(sData[2].ToString());
            string [] spoints = sData[3].ToString().Split(',');
            points = new XPoint[totpoints];
            for (int i = 0 ; i < totpoints; i++) {
                points[i] = new XPoint( double.Parse(spoints[iCtr]) ,double.Parse(spoints[iCtr + 1]));
                iCtr += 2;
            }
            DrawPolygon(fillColor, points);
        }


        private void DrawPolygon(string fillColor, XPoint [] points )
        {
            //XPen pen = new XPen(XColors.DarkBlue, 2.5);
            
            XBrush xBrush = XBrush(fillColor);
            gfx.DrawPolygon(xBrush, points, XFillMode.Winding );
        }


        private void AddHLine1(string str)
        {
            string sLineType = "";
            string[] sData = str.Split(sChar);
            int Left = int.Parse(sData[1].ToString());
            int Top = int.Parse(sData[2].ToString());
            int Width = int.Parse(sData[3].ToString());
            string sColorCode = sData[4].ToString();

            if (sColorCode.Trim() == "")
                sColorCode = "0";
            if (sData.Length >= 6)
                sLineType = sData[5].ToString();

            DrawHLine(Left, Top, Width, sColorCode, sLineType);
        }

        private void AddVLine1(string str)
        {
            string[] sData = str.Split(sChar);
            int Left = int.Parse(sData[1].ToString());
            int Top = int.Parse(sData[2].ToString());
            int Height = int.Parse(sData[3].ToString());
            string sColorCode = sData[4].ToString();

            if (sColorCode.Trim() == "")
                sColorCode = "0";

            DrawVLine(Left, Top, Height, sColorCode);
        }

        private void AddVLine2(string str)
        {
            string[] sData = str.Split(sChar);
            int Left = int.Parse(sData[1].ToString());
            int Top = int.Parse(sData[2].ToString());
            int Height = int.Parse(sData[3].ToString());
            int Width = int.Parse(sData[4].ToString());
            int Radius = int.Parse(sData[5].ToString());
            string sColorCode = sData[6].ToString();

            if (sColorCode.Trim() == "")
                sColorCode = "0";

            DrawVLine(Left, Top, Height, Width, Radius, sColorCode);
        }

        private void AddBox(string str)
        {
            string[] sData = str.Split(sChar);
            int Left = int.Parse(sData[1].ToString());
            int Top = int.Parse(sData[2].ToString());
            int Height = int.Parse(sData[3].ToString());
            int Width = int.Parse(sData[4].ToString());

            int BorderThicknes = int.Parse(sData[5].ToString());

            string BorderColor = sData[6].ToString();
            string FillColor = sData[7].ToString();
            DrawBox(Left, Top, Height, Width, BorderThicknes, BorderColor, FillColor);
        }

        private XColor XColor(string str)
        {
            XColor clr = XColors.Black;
            if (str.ToUpper() == "GREEN")
                clr = XColors.Green;
            if (str.ToUpper() == "RED")
                clr = XColors.Red;
            if (str.ToUpper() == "GRAY")
                clr = XColors.Gray;
            if (str.ToUpper() == "DARKGRAY")
                clr = XColors.DarkGray;
            if (str.ToUpper() == "BLUE")
                clr = XColors.Blue;
            if (str.ToUpper() == "WHITE")
                clr = XColors.White;
            if (str.ToUpper() == "LIGHTGRAY")
                clr = XColors.LightGray;
            if (str.ToUpper() == "DEEPSKYBLUE")
                clr = XColors.DeepSkyBlue;
            if (str.ToUpper() == "YELLOW")
                clr = XColors.Yellow;
            if (str.ToUpper() == "ORANGE")
                clr = XColors.Orange;

            return clr;
        }

        private XBrush XBrush(string str)
        {
            XBrush clr = XBrushes.Black;
            if (str.ToUpper() == "GREEN")
                clr = XBrushes.Green;
            if (str.ToUpper() == "RED")
                clr = XBrushes.Red;
            if (str.ToUpper() == "GRAY")
                clr = XBrushes.Gray;
            if (str.ToUpper() == "DARKGRAY")
                clr = XBrushes.DarkGray;

            if (str.ToUpper() == "BLUE")
                clr = XBrushes.Blue;
            if (str.ToUpper() == "WHITE")
                clr = XBrushes.White;
            if (str.ToUpper() == "LIGHTGRAY")
                clr = XBrushes.LightGray;
            if (str.ToUpper() == "DEEPSKYBLUE")
                clr = XBrushes.DeepSkyBlue;
            if (str.ToUpper() == "YELLOW")
                clr = XBrushes.Yellow;
            if (str.ToUpper() == "ORANGE")
                clr = XBrushes.Orange;
            

            return clr;
        }

        private void SavePdf()
        {

            try
            {
                _Doc.Save(FileName);

            }
            catch (Exception Ex)
            {
                IsError = true;
                sError = Ex.Message.ToString();
                //MessageBox.Show(Ex.Message.ToString(), "Export Pdf", MessageBoxButtons.OK);
            }
        }


        private void AddImage(string str)
        {
            string[] sData = str.Split(sChar);

            string ImageName = sData[1].ToString();
            int Left = int.Parse(sData[2].ToString());
            int Top = int.Parse(sData[3].ToString());
            int Height = int.Parse(sData[4].ToString());
            int Width = int.Parse(sData[5].ToString());
            int stretch = int.Parse(sData[6].ToString());
            int PdfY = int.Parse(sData[7].ToString());

            DrawImage(ImageName, Left, Top, Height, Width,stretch, PdfY);

        }


        private static Image reesizeImage(Image imgToResize, Size size)
        {

            int sourceWidth = imgToResize.Width;
            int sourceHeight = imgToResize.Height;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)size.Width / (float)sourceWidth);
            nPercentH = ((float)size.Height / (float)sourceHeight);

            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            

            g.Dispose();

            return (Image)b;
        }




        private void DrawImage(string ImgName, int x, int y, int height, int width, int stretch, int Pdy)
        {
            try
            {
                if (System.IO.File.Exists(ImgName))
                {

                    Image image = Image.FromFile(ImgName);
                    image = reesizeImage(image, new Size(width, height));
                    XImage newImage = XImage.FromGdiPlusImage(image);
                    if (height > 0 && width > 0)
                        gfx.DrawImage(newImage,x,y, image.Width, image.Height);
                    else
                        gfx.DrawImage(newImage, x, y);
                }
            }
            catch (Exception)
            {

            }
        }



        private void AddImage2(string str)
        {
            string[] sData = str.Split(sChar);
            int Left = int.Parse(sData[1].ToString());
            int Top = int.Parse(sData[2].ToString());
            string ImgName = sData[3].ToString();
            int pdfy = int.Parse(sData[4].ToString());
            LoadImage2(Left, Top, ImgName, pdfy);
        }

        private void LoadImage2(int x, int y, string ImgName, int pdfy = 100)
        {
            //try
            //{
            //    BitmapImage bi = new BitmapImage();
            //    bi.UriSource = new Uri(GLOBALCONTANTS.WWW_ROOT + "/Images/" + ImgName);
            //    Image img = new Image();
            //    img.Source = bi;

            //    Canvas ImageCanvas = new Canvas();
            //    ImageCanvas.Width = img.Width;
            //    ImageCanvas.Height = img.Height;
            //    ImageCanvas.Children.Add(img);

            //    int imgh = (int)img.Height;
            //    _Page.addImage(ImageCanvas, null, x, iHeight - (y + pdfy));

            //}
            //catch (Exception)
            //{

            //}
        }



        private void DrawWaterMarkDiagonal(int wx, int wy, string wText, XFont wFont, Color wColor)
        {
            XSize sf = gfx.MeasureString(wText, wFont);
            gfx.TranslateTransform(wx / 2, wy / 2);

            double Agle = -Math.Atan(wy / wx) * 90 / Math.PI;
            gfx.RotateTransform((float)Agle);
            gfx.TranslateTransform(-wx / 2, -wy / 2);

            XStringFormat wformat = new XStringFormat();
            wformat.Alignment = XStringAlignment.Near;
            wformat.LineAlignment = XLineAlignment.Near;

            XBrush wbrush = new XSolidBrush(Color.FromArgb(20, wColor));

            wx = (int)(wx - sf.Width) / 2;
            wy = (int)(wy - sf.Height) / 2;

            XPoint Point_xy = new XPoint(wx, wy);
            gfx.DrawString(wText, wFont, wbrush, Point_xy, wformat);
        }


    }
}
