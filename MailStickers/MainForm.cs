using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using Excel = Microsoft.Office.Interop.Excel;

namespace MailStickers
{
    public partial class MainForm : Form
    {
        protected XColor ShadowColor;
        protected double BorderWidth;
        protected XPen BorderPen;
        protected XColor BackColor;
        protected XColor BackColor2;

        XGraphicsState state;

        public MainForm()
        {
            InitializeComponent();

            this.ShadowColor = XColors.Gainsboro;
            this.BorderWidth = 4.5;
            this.BorderPen = new XPen(XColor.FromArgb(94, 118, 151), this.BorderWidth);

            this.BackColor = XColors.Ivory;
            this.BackColor2 = XColors.WhiteSmoke;

            this.BackColor = XColor.FromArgb(212, 224, 240);
            this.BackColor2 = XColor.FromArgb(253, 254, 254);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public void BeginBox(XGraphics gfx, int number, string title)
        {
            const int dEllipse = 15;
            XRect rect = new XRect(0, 20, 300, 200);
            if (number % 2 == 0)
                rect.X = 300 - 5;
            rect.Y = 40 + ((number - 1) / 2) * (200 - 5);
            rect.Inflate(-10, -10);
            XRect rect2 = rect;
            rect2.Offset(this.BorderWidth, this.BorderWidth);
            gfx.DrawRoundedRectangle(new XSolidBrush(this.ShadowColor), rect2, new XSize(dEllipse + 8, dEllipse + 8));
            XLinearGradientBrush brush = new XLinearGradientBrush(rect, this.BackColor, this.BackColor2, XLinearGradientMode.Vertical);
            gfx.DrawRoundedRectangle(this.BorderPen, brush, rect, new XSize(dEllipse, dEllipse));
            rect.Inflate(-5, -5);

            XFont font = new XFont("Verdana", 12, XFontStyle.Regular);
            gfx.DrawString(title, font, XBrushes.Navy, rect, XStringFormats.TopCenter);

            rect.Inflate(-10, -5);
            rect.Y += 20;
            rect.Height -= 20;
            //gfx.DrawRectangle(XPens.Red, rect);

            this.state = gfx.Save();
            gfx.TranslateTransform(rect.X, rect.Y);
        }

        public void EndBox(XGraphics gfx)
        {
            gfx.Restore(this.state);
        }

        void MeasureText(XGraphics gfx, int number, string xtext)
        {
            const XFontStyle style = XFontStyle.Regular;
            XFont font = new XFont("Times New Roman", 20, style);

            string text = xtext;
            const double x = 20, y = 100;
            XSize size = gfx.MeasureString(text, font);

            double lineSpace = font.GetHeight(gfx);
            int cellSpace = font.FontFamily.GetLineSpacing(style);
            int cellAscent = font.FontFamily.GetCellAscent(style);
            int cellDescent = font.FontFamily.GetCellDescent(style);
            int cellLeading = cellSpace - cellAscent - cellDescent;

            double ascent = lineSpace * cellAscent / cellSpace;
            gfx.DrawRectangle(XBrushes.Bisque, x, y - ascent, size.Width, ascent);

            double descent = lineSpace * cellDescent / cellSpace;
            gfx.DrawRectangle(XBrushes.LightGreen, x, y, size.Width, descent);

            double leading = lineSpace * cellLeading / cellSpace;
            gfx.DrawRectangle(XBrushes.Yellow, x, y + descent, size.Width, leading);

            XColor color = XColors.DarkSlateBlue;
            color.A = 0.6;
            gfx.DrawString(text, font, new XSolidBrush(color), x, y);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int startCell = (int)numericUpDown1.Value;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();

            xlWorkBook = xlApp.Workbooks.Open(textBoxFN.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];

            int countr = 0;

            int cnt = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row - startCell + 1;

            MessageBox.Show(cnt.ToString());

            PdfDocument s_document;

            string filename = String.Format("{0}_tempfile.pdf", Guid.NewGuid().ToString("D").ToUpper());
            s_document = new PdfDocument();
            s_document.Info.Title = "";
            s_document.Info.Author = "";
            s_document.Info.Subject = "";
            s_document.Info.Keywords = "";

            XGraphics gfx = XGraphics.FromPdfPage(s_document.AddPage());

            XSize page = gfx.PageSize;

            XRect rect;
            XRect rectTxt;

            XTextFormatter tf = new XTextFormatter(gfx);
            const XFontStyle style = XFontStyle.Regular;

            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            XFont font = new XFont("Verdana", 8, style, options);

            XFont fontBig = new XFont("Verdana", 15, style, options);

            XBrush xb = new XSolidBrush(XColor.FromArgb(Color.Black));

            string clmn;

            int stickrow = 0;
            int stickcol = 0;

            int stickwidth = (int)(page.Width / (int)numericUpDownStickCols.Value);
            int stickhegth = (int)(page.Height / (int)numericUpDownStickRows.Value);

            try
            {
                do
                {
                    var cell1 = xlWorkSheet.Cells[countr + startCell, 1];  //Организация

                    string clmn1 = cell1.Value != null ? (cell1.Value2.ToString()).Trim() : string.Empty;
                    clmn = cell1.Value != null ? (cell1.Value2.ToString()).Trim() : string.Empty;

                    if (clmn == string.Empty)
                    {
                        break;
                    }

                    if (stickcol == (int)numericUpDownStickRows.Value)
                    {
                        stickcol = 0;
                        stickrow++;
                    }

                    if (stickrow == (int)numericUpDownStickCols.Value)
                    {
                        stickrow = 0;
                        stickcol = 0;

                        gfx = XGraphics.FromPdfPage(s_document.AddPage());
                        page = gfx.PageSize;

                        tf = new XTextFormatter(gfx);

                    }

                    var cell2 = xlWorkSheet.Cells[countr + startCell, 2];  //Кому
                    string clmn2 = cell2.Value != null ? (cell2.Value2.ToString()).Trim() : string.Empty;

                    var cell3 = xlWorkSheet.Cells[countr + startCell, 3]; //Индекс
                    string clmn3 = cell3.Value != null ? (cell3.Value2.ToString()).Trim() : string.Empty;

                    var cell4 = xlWorkSheet.Cells[countr + startCell, 4]; //Улица
                    string clmn4 = cell4.Value != null ? (cell4.Value2.ToString()).Trim() : string.Empty;

                    var cell5 = xlWorkSheet.Cells[countr + startCell, 5]; //Дом
                    string clmn5 = cell5.Value != null ? (cell5.Value2.ToString()).Trim() : string.Empty;

                    var cell6 = xlWorkSheet.Cells[countr + startCell, 6];  //Город
                    string clmn6 = cell6.Value != null ? (cell6.Value2.ToString()).Trim() : string.Empty;

                    var cell7 = xlWorkSheet.Cells[countr + startCell, 7];  //Область
                    string clmn7 = cell7.Value != null ? (cell7.Value2.ToString()).Trim() : string.Empty;

                    var cell8 = xlWorkSheet.Cells[countr + startCell, 8];  //Страна
                    string clmn8 = cell8.Value != null ? (cell8.Value2.ToString()).Trim() : string.Empty;

                    rect = new XRect(stickwidth * stickrow, stickhegth * stickcol, stickwidth, stickhegth);

                    rectTxt = new XRect(stickwidth * stickrow + 15, stickhegth * stickcol + 10, stickwidth - 20, stickhegth - 20);

                    int toSixLines = (int)((stickhegth - 10) / 6);

                    XSize size = gfx.MeasureString("5", fontBig);

                    size.Width = size.Width + 3;

                    XRect rectTxt1 = new XRect(stickwidth * stickrow + 15, stickhegth * stickcol + 10 + toSixLines * 0, stickwidth - 30, toSixLines); //stickhegth - 20 - toSixLines * 0
                    XRect rectTxt2 = new XRect(stickwidth * stickrow + 15, stickhegth * stickcol + 10 + toSixLines * 1, stickwidth - 30, toSixLines);
                    XRect rectTxt3 = new XRect(stickwidth * stickrow + 15, stickhegth * stickcol + 10 + toSixLines * 2, stickwidth - 30, toSixLines);
                    XRect rectTxt4 = new XRect(stickwidth * stickrow + 15, stickhegth * stickcol + 10 + toSixLines * 3, stickwidth - 30, toSixLines);
                    XRect rectTxt5 = new XRect(stickwidth * stickrow + 15, stickhegth * stickcol + 10 + toSixLines * 4, stickwidth - 30, toSixLines);

                    XRect rectTxt5Ind1 = new XRect(stickwidth * stickrow + 15 + size.Width * 0, stickhegth * stickcol + 5 + toSixLines * 4,
                        size.Width, toSixLines);
                    XRect rectTxt5Ind2 = new XRect(stickwidth * stickrow + 15 + size.Width * 1, stickhegth * stickcol + 5 + toSixLines * 4,
                        size.Width, toSixLines);
                    XRect rectTxt5Ind3 = new XRect(stickwidth * stickrow + 15 + size.Width * 2, stickhegth * stickcol + 5 + toSixLines * 4,
                        size.Width, toSixLines);
                    XRect rectTxt5Ind4 = new XRect(stickwidth * stickrow + 15 + size.Width * 3, stickhegth * stickcol + 5 + toSixLines * 4,
                        size.Width, toSixLines);
                    XRect rectTxt5Ind5 = new XRect(stickwidth * stickrow + 15 + size.Width * 4, stickhegth * stickcol + 5 + toSixLines * 4,
                        size.Width, toSixLines);
                    XRect rectTxt5Ind6 = new XRect(stickwidth * stickrow + 15 + size.Width * 5, stickhegth * stickcol + 5 + toSixLines * 4,
                        size.Width, toSixLines);

                    XRect rectTxt6 = new XRect(stickwidth * stickrow + 15, stickhegth * stickcol + 10 + toSixLines * 5, stickwidth - 30, toSixLines);

                    listBox1.Items.Add(rect);

                    stickcol++;

                    string toOut = string.Format("{0} \n{1} \n{2} \n{3}", clmn1, clmn2, clmn3, clmn4);

                    XPen pen = new XPen(XColors.Gray, 0.01);
                    XPen penBlack = new XPen(XColors.Black, 0.01);

                    gfx.DrawRectangle(pen, rect);
                    tf.Alignment = XParagraphAlignment.Left;

                    gfx.DrawLine(penBlack, new Point((int)rectTxt1.X, (int)(rectTxt1.Y + rectTxt1.Height - 5)), new Point((int)(rectTxt1.X + rectTxt1.Width), (int)(rectTxt1.Y + rectTxt1.Height - 5)));
                    gfx.DrawLine(penBlack, new Point((int)rectTxt2.X, (int)(rectTxt2.Y + rectTxt2.Height - 5)), new Point((int)(rectTxt2.X + rectTxt2.Width), (int)(rectTxt2.Y + rectTxt2.Height - 5)));
                    gfx.DrawLine(penBlack, new Point((int)rectTxt3.X, (int)(rectTxt3.Y + rectTxt3.Height - 5)), new Point((int)(rectTxt3.X + rectTxt3.Width), (int)(rectTxt3.Y + rectTxt3.Height - 5)));
                    gfx.DrawLine(penBlack, new Point((int)rectTxt4.X, (int)(rectTxt4.Y + rectTxt4.Height - 5)), new Point((int)(rectTxt4.X + rectTxt4.Width), (int)(rectTxt4.Y + rectTxt4.Height - 5)));
                    gfx.DrawLine(penBlack, new Point((int)rectTxt5.X, (int)(rectTxt5.Y + rectTxt5.Height - 5)), new Point((int)(rectTxt5.X + rectTxt5.Width), (int)(rectTxt5.Y + rectTxt5.Height - 5)));
                    gfx.DrawLine(penBlack, new Point((int)rectTxt6.X, (int)(rectTxt6.Y + rectTxt6.Height - 5)), new Point((int)(rectTxt6.X + rectTxt6.Width), (int)(rectTxt6.Y + rectTxt6.Height - 5)));

                    tf.DrawString(clmn1, font, xb, rectTxt1, XStringFormats.TopLeft);
                    tf.DrawString(clmn2, font, xb, rectTxt2, XStringFormats.TopLeft);
                    tf.DrawString(clmn4, font, xb, rectTxt3, XStringFormats.TopLeft);

                    tf.Alignment = XParagraphAlignment.Right;
                    tf.DrawString(string.Format("{0} {1}", "д.", clmn5.ToUpper()), font, xb, rectTxt4, XStringFormats.TopLeft);
                    tf.DrawString(clmn7 != string.Empty ? clmn7 : clmn8, font, xb, rectTxt6, XStringFormats.TopLeft);
                    tf.DrawString(clmn6, font, xb, rectTxt5, XStringFormats.TopLeft);

                    tf.Alignment = XParagraphAlignment.Center;

                    if (clmn3.Length >= 6)
                    {
                        tf.DrawString(clmn3[0].ToString(CultureInfo.InvariantCulture), fontBig, xb, rectTxt5Ind1,
                            XStringFormats.TopLeft);
                        tf.DrawString(clmn3[1].ToString(CultureInfo.InvariantCulture), fontBig, xb, rectTxt5Ind2,
                            XStringFormats.TopLeft);
                        tf.DrawString(clmn3[2].ToString(CultureInfo.InvariantCulture), fontBig, xb, rectTxt5Ind3,
                            XStringFormats.TopLeft);
                        tf.DrawString(clmn3[3].ToString(CultureInfo.InvariantCulture), fontBig, xb, rectTxt5Ind4,
                            XStringFormats.TopLeft);
                        tf.DrawString(clmn3[4].ToString(CultureInfo.InvariantCulture), fontBig, xb, rectTxt5Ind5,
                            XStringFormats.TopLeft);
                        tf.DrawString(clmn3[5].ToString(CultureInfo.InvariantCulture), fontBig, xb, rectTxt5Ind6,
                            XStringFormats.TopLeft);

                        int hY = (int)rectTxt5.Y - 5;
                        int lY = (int)(rectTxt5.Y + rectTxt5.Height) - 5;

                        gfx.DrawLine(penBlack, new Point((int)rectTxt5Ind1.X, hY), new Point((int)rectTxt5Ind1.X, lY));
                        gfx.DrawLine(penBlack, new Point((int)rectTxt5Ind2.X, hY), new Point((int)rectTxt5Ind2.X, lY));
                        gfx.DrawLine(penBlack, new Point((int)rectTxt5Ind3.X, hY), new Point((int)rectTxt5Ind3.X, lY));
                        gfx.DrawLine(penBlack, new Point((int)rectTxt5Ind4.X, hY), new Point((int)rectTxt5Ind4.X, lY));
                        gfx.DrawLine(penBlack, new Point((int)rectTxt5Ind5.X, hY), new Point((int)rectTxt5Ind5.X, lY));
                        gfx.DrawLine(penBlack, new Point((int)rectTxt5Ind6.X, hY), new Point((int)rectTxt5Ind6.X, lY));
                        gfx.DrawLine(penBlack, new Point((int)(rectTxt5Ind6.X + rectTxt5Ind6.Width), hY),
                            new Point((int)(rectTxt5Ind6.X + rectTxt5Ind6.Width), lY));
                    }

                    countr++;
                } while (clmn != string.Empty);
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка :(");
            }

            s_document.Save(filename);

            Process.Start(filename);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "MS Excell files (*.xls*)|*.xls*|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBoxFN.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
    }
}