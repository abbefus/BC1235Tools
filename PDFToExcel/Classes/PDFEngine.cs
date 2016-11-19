using java.awt.geom;
using OfficeOpenXml;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.pdmodel.common;
using org.apache.pdfbox.pdmodel.graphics;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using static ABUtils.FileUtils;

namespace PDFToExcel
{
    public static class PDFEngine
    {
        
        //public static IEnumerable<TextLine> ParseToLines(StrippedPDFPage page)
        //{
        //    int[] yvals = textchars.Keys.OrderBy(x => x).ToArray();

        //    for (int i = 0; i < yvals.Length - 1; i++)
        //    {
        //        int diff = Math.Abs(yvals[i + 1] - yvals[i]);
        //        if (diff < PDFMetrics.LINELIMIT)
        //        {
        //            textchars[yvals[i + 1]].AddRange(textchars[yvals[i]]);
        //            continue;
        //        }
        //        yield return new TextLine(textchars[yvals[i]], yvals[i], diff);
        //    }
        //}

        public static void TabifyPDF(string pdfpath, int numcolumns=1, int? startpage=null, int? endpage=null, string headerstring=null)
        {
            PDDocument doc = PDDocument.load(pdfpath);
            DecryptPDF(ref doc);

            // turn user-intuitive pages (1-inclusive) into PDFBox intuitive pages (0-exclusive)
            startpage = startpage > 0 ? startpage - 1 : null;
            endpage = endpage >= startpage ? endpage : null;

            StrippedPDF pdf = StripPDF(ref doc, startpage, endpage);

            doc.close();
            

            List<ClassifiedPDFRow> lines = new List<ClassifiedPDFRow>();
            int count = 0;
            foreach (TextLine line in new TextLine[3])//ParseToLines(doc, startpage, endpage))
            {
                TextSet[] txtsets = line.GenerateTextSets().ToArray();

                PDFRowClass tableclass;
                tableclass = (txtsets.Length > 1 && txtsets.Length <= numcolumns) ? PDFRowClass.data : PDFRowClass.unknown;
                lines.Add(new ClassifiedPDFRow
                {
                    LineType = tableclass,
                    TextSets = txtsets,
                    PageNumber = line.PageNumber,
                    YIndex = Convert.ToInt32(line.Index),
                    Index = count++
                });
            }

            // sort out table metrics
            PDFColumn[] table = new PDFColumn[numcolumns];
            float margin = lines.Where(x=> x.LineType == PDFRowClass.data).Min(x => x.X);

            for (int i = 0; i < numcolumns; i++)
            {

            }
                
                
        }
        public static bool DecryptPDF(ref PDDocument doc)
        {
            if (doc.isEncrypted())
            {
                try
                {
                    doc.decrypt("");
                    doc.setAllSecurityToBeRemoved(true);
                    return true;
                }
                catch (Exception e)
                {
                    MessageBox.Show(string.Format("The document is encrypted, and we can't decrypt it.\r\n{0}", e.Message));
                    return false;
                }
            }
            else
            {
                return true;
            }
        }
        private static StrippedPDF StripPDF(ref PDDocument doc, int? startpage = null, int? endpage = null, Rectangle2D parseregion = null)
        {
            int start = startpage ?? 0;
            int end = endpage ?? doc.getNumberOfPages();
            java.util.ArrayList pages = GetPageRange(ref doc, start, end);  //potentially same as cat.getPageNodes().getKids()
            MaskedStripper stripper = new MaskedStripper();
            return stripper.stripPDF(pages);
        }

        private static java.util.ArrayList GetPageRange(ref PDDocument doc, int start, int end)
        {
            java.util.ArrayList pages = new java.util.ArrayList();

            try
            {
                PDDocumentCatalog cat = doc.getDocumentCatalog();
                java.util.List catpages = cat.getAllPages();

                for (int i = start; i < end; i++)
                {
                    pages.add(catpages.get(i));
                }

                return pages;
            }
            catch
            {
                throw new ArgumentOutOfRangeException("Page range not found in document");
            }
            
        }


    }
    public class PDFColumn
    {
        public float X { get; set; }
        public float Width { get; set; }
    }
    public enum PDFRowClass
    {
        header,
        data,
        delete,
        unknown
    }


    public struct TextChar
    {
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float Width { get; set; }
        public float SpaceAfter { get; set; }
        public float SpaceWidth { get; set; }
        public char Char { get; set; }
        public string Font { get; set; }
        public float FontSize { get; set; }
        //public int PageNumber { get; set; }
    }
    public class TextLine
    {
        public TextChar[] TextChars { get; set; }
        public int PageNumber { get; set; }
        public float Index { get; set; }
        public int LinespaceAfter { get; set; }

        public IOrderedEnumerable<TextSet> GenerateTextSets(double toleranceAdjustment=0)
        {
            double tol = PDFMetrics.CHARSPACING_TOLERANCE - toleranceAdjustment;
            try
            {
                List<TextSet> seqlist = new List<TextSet>();

                bool startofsequence = true;
                TextSet seq = new TextSet();
                StringBuilder sb = new StringBuilder();
                int count = 0;

                for (int i = 0; i < TextChars.Length; i++)
                {
                    count++;
                    TextChar tc = TextChars[i];
                    sb.Append(tc.Char);
                    if (startofsequence)
                    {
                        seq = new TextSet { StartX = tc.X };
                        startofsequence = false;
                    }
                    if (tc.SpaceAfter > (tc.SpaceWidth + tol))
                    {
                        seq.EndX = tc.X + tc.Width;
                        seq.Count = count;
                        seq.Text = sb.ToString();
                        seq.SpaceAfter = tc.SpaceAfter;
                        seq.SpaceWidth = tc.SpaceWidth;
                        seqlist.Add(seq); //yield return seq

                        sb.Clear();
                        startofsequence = true;
                        count = 0;
                    }
                }
                TextChar lasttc = TextChars.LastOrDefault();
                seq.EndX = lasttc.X + lasttc.Width;
                seq.Count = count;
                seq.Text = sb.ToString();
                seq.SpaceAfter = lasttc.SpaceAfter;
                seq.SpaceWidth = lasttc.SpaceWidth;
                seqlist.Add(seq);
                return seqlist.OrderBy(x => x.StartX);
            }
            catch
            {
                throw new Exception("TextChars is empty or not populated correctly.");
            }
        }
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (TextChar tc in TextChars)
            {
                sb.Append(tc.Char);
                if (tc.SpaceAfter > (tc.SpaceWidth + PDFMetrics.CHARSPACING_TOLERANCE))
                {
                    int numspaces = (int)Math.Floor(tc.SpaceAfter / tc.SpaceWidth);
                    sb.Append(' ', numspaces);
                }                
            }
            return sb.ToString();
        }
    }



    public class TextSet
    {
        public float StartX { get; set; }
        public float EndX { get; set; }
        public float Width { get { return EndX - StartX; } }
        public int Count { get; set; }
        public string Text { get; set; }
        public float SpaceAfter { get; set; }
        public float SpaceWidth { get; set; }
        public int TrailingSpaces
        {
            get { return (int)Math.Round(SpaceAfter / SpaceWidth); }
        }
    }
    public class ClassifiedPDFRow : INotifyPropertyChanged
    {
        private PDFRowClass _lineType;
        public PDFRowClass LineType
        {
            get { return _lineType; }
            set
            {
                _lineType = value;
                NotifyPropertyChanged();
            }
        }
        private TextSet[] _textsets;
        public TextSet[] TextSets
        {
            get { return _textsets; }
            set
            {
                _textsets = value;
                TotalChars = _textsets.Sum(x => x.Count);
                float sumw = _textsets.Sum(x => x.Width);
                AverageCharWidth = sumw / TotalChars;
                X = _textsets.Min(x => x.StartX);
            }
        }
        public int TotalChars { get; private set; }
        public float AverageCharWidth { get; private set; }
        public int PageNumber { get; set; }
        public int YIndex { get; set; }
        public int Index { get; set; }
        public float X { get; set; }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (TextSet ts in _textsets)
            {
                sb.Append(ts.Text).Append(' ', ts.TrailingSpaces);
            }
            return sb.ToString();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }

    public class StrippedPDF
    {
        public StrippedPDFPage[] Pages { get; set; }
    }
    public class StrippedPDFPage
    {
        public StrippedPDFPage(int pagenumber)
        {
            PageNumber = pagenumber;
        }
        public int PageNumber { get; set; }
        public Size PageSize { get; set; }

        public TextChar[][] TextCharLines { get; set; }


    }




    public class MaskedStripper : PDFTextStripper
    {
        public bool UseMask { get; set; }
        private Rectangle2D Mask { get; set; }

        //private List<TextChar> textChars = new List<TextChar>();
        private List<TextChar>[] tcPages;
        private TextPosition lasttp;
        private int lastPageNo = -1;


        public MaskedStripper(Rectangle2D mask=null) : base()
        {
            UseMask = mask != null;
            Mask = mask;
        }

        public StrippedPDF stripPDF(java.util.List pages)
        {
            int pagecount = pages.size();
            tcPages = new List<TextChar>[pagecount];
            base.processPages(pages);
            pages.clear();
            pages = null;
            AddEODTextPosition();

            StrippedPDF pdf = new StrippedPDF();
            List<StrippedPDFPage> pdfpages = new List<StrippedPDFPage>();
            for (int i = 0; i < pagecount; i++)
            {
                StrippedPDFPage pdfpage = new StrippedPDFPage(i);
                pdfpage.PageSize = UseMask ?
                    new Size
                    {
                        Width = Mask.getWidth(),
                        Height = Mask.getHeight()
                    } :
                    new Size
                    {
                        Width = (int)((PDPage)pages.get(i)).getMediaBox().getWidth(),
                        Height = (int)((PDPage)pages.get(i)).getMediaBox().getHeight()
                    };
                pdfpage.TextCharLines = tcPages[pdfpage.PageNumber].
                    GroupBy(tc => tc.Y).        // group into lines by y position
                    Select(grp => grp.OrderBy(x => x.X).ToArray()). // order group by x position
                    OrderBy(x => x.FirstOrDefault().Y).ToArray();   // order lines by y position
                pdfpages.Add(pdfpage);
            }
            pdf.Pages = pdfpages.ToArray();

            return pdf;
        }
        private void AddEODTextPosition()
        {
            if (lasttp == null) return;
            float Y = lasttp.getYDirAdj();
            TextChar lasttc = TextPositionToTextChar(lasttp);
            lasttc.Y = (float)Math.Round(Y, 0, MidpointRounding.ToEven);
            tcPages[lastPageNo].Add(lasttc);

            lasttp = null;
            lastPageNo = -1;
        }


        protected override void processTextPosition(TextPosition tp)
        {
            //y position and fontsize are dependent on the characters within the string of text
            //for example, a 'pqyg' increases the 'adjusted' y position relative to an 'oweruaszxcvnm' because of the tails
            //and vice versa for 'tidfhklb'
            //note 'j' goes in both directions
            //caps are all the same 'ABCDEFGHIJKLMNOPRSTUVWXYZ' except 'Q'
            //punctuation creates another situation with ',;' and '"'

            if (lasttp == null) { lasttp = tp; return; }

            float X = lasttp.getXDirAdj();
            float Y = lasttp.getYDirAdj();

            if (UseMask && isOOB(X, Y)) { lasttp = tp; return; }


            int newPgNumber = getCurrentPageNo();

            if (lastPageNo != -1 && lastPageNo != newPgNumber) //if new page
            {
                TextChar lasttc = TextPositionToTextChar(lasttp);
                lasttc.Y = (float)Math.Round(Y, 0, MidpointRounding.ToEven);
                tcPages[lastPageNo].Add(lasttc);
                lastPageNo = newPgNumber;
            }
            else // if same page or start
            {
                TextChar tc = TextPositionToTextChar(lasttp, tp);
                lastPageNo = newPgNumber;
                tc.Y = (float)Math.Round(Y, 0, MidpointRounding.ToEven);
                tcPages[newPgNumber].Add(tc);
            }
            lasttp = tp;
        }


        private bool isOOB(float x, float y)
        {
            return x < Mask.getMinX() || x > Mask.getMaxX() ||
                    y < Mask.getMinY() || y > Mask.getMaxY();
        }
        private TextChar TextPositionToTextChar(TextPosition tp, TextPosition nexttp = null)
        {
            if (tp.getCharacter().Length != 1) throw new IndexOutOfRangeException("Textposition does not contain 1 character.");

            TextChar tc = new TextChar();
            tc.Char = tp.getCharacter()[0];
            tc.X = tp.getXDirAdj();
            tc.Width = tp.getWidthDirAdj();
            tc.SpaceAfter = nexttp != null ? Math.Max(0, nexttp.getXDirAdj() - (tp.getXDirAdj() + tp.getWidthDirAdj())) : 0;
            tc.SpaceWidth = tp.getWidthOfSpace();

            tc.Font = tp.getFont().getBaseFont();
            tc.FontSize = tp.getFontSizeInPt();

            try
            {
                int[] flags = GetBits(tp.getFont().getFontDescriptor().getFlags());
                tc.IsBold = findBold(tp, flags, getGraphicsState());
                tc.IsItalic = findItalics(tp, flags);
            }
            catch { }

            return tc;
        }

        private static java.util.ArrayList GetPageRange(ref PDDocument doc, int start, int end)
        {
            java.util.ArrayList pages = new java.util.ArrayList();

            try
            {
                PDDocumentCatalog cat = doc.getDocumentCatalog();
                java.util.List catpages = cat.getAllPages();

                for (int i = start; i < end; i++)
                {
                    pages.add(catpages.get(i));
                }

                return pages;
            }
            catch
            {
                throw new ArgumentOutOfRangeException("Page range not found in document");
            }

        }
        private static int[] GetBits(int flags)
        {
            string bitstring = Convert.ToString(flags, 2) + "0";
            return bitstring.PadLeft(32, '0').Select(x => int.Parse(x.ToString())).ToArray().Reverse().ToArray();
        }
        private static bool findBold(TextPosition tp, int[] flags, PDGraphicsState graphicsstate)
        {
            float fontweight = tp.getFont().getFontDescriptor().getFontWeight();
            int renderingmode = graphicsstate.getTextState().getRenderingMode();
            

            bool isbold1 = fontweight >= 700;
            bool isbold2 = flags[19] == 1;
            bool isbold3 = renderingmode == 2;
            bool isbold4 = tp.getFont().getBaseFont().ToLower().Contains("bold");
            return isbold1 || isbold2 || isbold3 || isbold4;
        }
        private static bool findItalics(TextPosition tp, int[] flags)
        {
            bool isitalic1 = tp.getFont().getFontDescriptor().getItalicAngle() != 0;
            bool isitalic2 = flags[7] == 1;
            return isitalic1 || isitalic2;
        }

        protected override void writePage() { return; } //prevents exception
    }

    public static class PDFMetrics
    {
        public static readonly float CHARSPACING_TOLERANCE = 0.05F;
        public static readonly short LINELIMIT = 5;
        public static readonly Size LETTER_SIZE = new Size(612, 792);
    }

}



//double linewidth = graphicsstate.getLineWidth();
//Matrix matrix = graphicsstate.getCurrentTransformationMatrix();
//Matrix matrix2 = getTextMatrix();
//PDMatrix matrix3 = getGraphicsState().getTextState().getFont().getFontMatrix();
//PDRectangle bbox = tp.getFont().getFontDescriptor().getFontBoundingBox();

//protected override void processTextPosition(TextPosition tp)
//{
//    //y position and fontsize are dependent on the characters within the string of text
//    //for example, a 'pqyg' increases the 'adjusted' y position relative to an 'oweruaszxcvnm' because of the tails
//    //and vice versa for 'tidfhklb'
//    //note 'j' goes in both directions
//    //caps are all the same 'ABCDEFGHIJKLMNOPRSTUVWXYZ' except 'Q'
//    //punctuation creates another situation with ',;' and '"'

//    if (lasttp == null) { lasttp = tp; return; }

//    float X = lasttp.getXDirAdj();
//    float Y = lasttp.getYDirAdj();

//    if (UseMask && isOOB(X, Y)) { lasttp = tp; return; }


//    int newPgNumber = getCurrentPageNo();

//    if (lastPageNo != -1 && lastPageNo != newPgNumber) //if new page
//    {
//        TextChar lasttc = TextPositionToTextChar(lasttp);
//        lasttc.PageNumber = lastPageNo;
//        lasttc.Y = (float)Math.Round(Y, 0, MidpointRounding.ToEven);
//        lastPageNo = newPgNumber;
//        textChars.Add(lasttc);
//    }
//    else // if same page or start
//    {
//        TextChar tc = TextPositionToTextChar(lasttp, tp);
//        lastPageNo = newPgNumber;
//        tc.PageNumber = newPgNumber;
//        tc.Y = (float)Math.Round(Y, 0, MidpointRounding.ToEven);
//        textChars.Add(tc);
//    }
//    lasttp = tp;
}   // main processor

//public class TurbidityDataTypeI
//{
//    public DateTime Time { get; set; }
//    public double? Turbidity { get; set; }
//}
//public static class MathFunctions
//{
//    public static float RoundToNearest(double x, double xbase)
//    {
//        return (float)(xbase * Math.Round(x / xbase, MidpointRounding.AwayFromZero));
//    }
//}
//ToString function takes space adjustment argument
//public struct Block
//{
//    public TextLine[] textlines { get; set; }
//    public int index { get; set; }
//    public int pagenumber { get; set; }
//    public int classification { get; set; }
//    public Rectangle2D bbox { get; set; }

//    public string ToString(float adjustment)
//    {
//        StringBuilder sb = new StringBuilder();

//        foreach (TextLine textline in textlines)
//        {
//            for (int i = 0; i < textline.TextChars.Length - 1; i++)
//            {
//                sb.Append(textline.TextChars[i].CHAR);
//                int numspacesafter = (int)Math.Round(textline.TextChars[i].spaceafter / (textline.TextChars[i].spacewidth + adjustment), MidpointRounding.AwayFromZero);
//                sb.Append(' ', numspacesafter);
//            }
//            sb.Append(textline.TextChars[textline.TextChars.Length - 1].CHAR);
//        }
//        return sb.ToString();
//    }
//}

//private static Dictionary<int, List<TextChar>> _textchardict;
//private static TextLine[] _textlines;
//parses to excel
//public void ParseToExcel(string pdf)
//{
//    Console.WriteLine("Reading PDF...");
//    PDDocument doc = PDDocument.load(pdf);
//    if (doc.isEncrypted())
//    {
//        try
//        {
//            doc.decrypt("");
//            doc.setAllSecurityToBeRemoved(true);
//        }
//        catch (Exception e)
//        {
//            MessageBox.Show(string.Format("The document is encrypted, and we can't decrypt it.\r\n{0}", e.Message));
//            return;
//        }
//    }

//    _textchardict = RunStripper(ref doc);
//    OldAssembleTextLines();

//    WriteDataToExcel(ExtractTurbidityDataTypeI());
//}
////pdfbox static classes
//private static void OldAssembleTextLines()
//{
//    List<TextLine> textlines = new List<TextLine>();
//    List<int> yvals = _textchardict.Keys.ToList();
//    yvals.Sort();

//    for (int i = 0; i < yvals.Count - 1; i++)
//    {
//        int diff = Math.Abs(yvals[i + 1] - yvals[i]);
//        if (diff < LINELIMIT)
//        {
//            _textchardict[yvals[i + 1]].AddRange(_textchardict[yvals[i]]);
//            continue;
//        }
//        textlines.Add(new TextLine(_textchardict[yvals[i]], yvals[i], diff));
//    }
//    _textlines = textlines.ToArray();
//}





//private void WriteDataToExcel(TurbidityDataTypeI[] turbidity)
//{
//    ExcelPackage xpkg = new ExcelPackage();
//    ExcelWorksheet ws = xpkg.Workbook.Worksheets.Add("PDF Result");

//    ws.Cells[1, 1].Value = header[0];
//    ws.Cells[1, 2].Value = header[1];
//    ws.Cells[1, 3].Value = header[2];

//    for (int i = 0; i < turbidity.Length; i++)
//    {
//        ws.Cells[i + 2, 1].Value = turbidity[i].Time.ToShortDateString();
//        ws.Cells[i + 2, 2].Value = turbidity[i].Time.ToShortTimeString();
//        ws.Cells[i + 2, 3].Value = turbidity[i].Turbidity;
//    }

//    ws.Cells.AutoFitColumns(0);
//    string folder = @"C:\Users\abefus\Documents\Visual Studio 2015\Projects\Compliance Document Parser\Compliance Document Parser\Results";
//    string name = "Test";
//    string path = Increment(Path.Combine(folder, name + ".xlsx"));
//    FileInfo fi = new FileInfo(path);
//    xpkg.SaveAs(fi);
//}
//private TurbidityDataTypeI[] ExtractTurbidityDataTypeI()
//{
//    short numcolumns = 3;
//    bool datamode = false;
//    char[] separator = "\t".ToCharArray();
//    int year = 2013;
//    string format = "dd-MMM-yyyy HH:mm";

//    List<TurbidityDataTypeI> data = new List<TurbidityDataTypeI>();

//    for (int i = 0; i < _textlines.Length; i++)
//    {
//        if (datamode)
//        {
//            string line = _textlines[i].ToString();
//            string[] linearray = line.Split(separator);
//            if (linearray.Length != numcolumns)
//            {
//                datamode = false;//falls through
//            }
//            else
//            {
//                linearray[0] = linearray[0] + "-" + year.ToString();
//                try
//                {
//                    string stringdate = string.Join(" ", linearray.Take(2));
//                    DateTime dt = DateTime.ParseExact(stringdate, format, CultureInfo.InvariantCulture);
//                    double turb = double.Parse(linearray[2]);
//                    data.Add(new TurbidityDataTypeI
//                    {
//                        Time = dt,
//                        Turbidity = turb
//                    });
//                }
//                catch
//                {
//                    data.Add(new TurbidityDataTypeI
//                    {
//                        Time = DateTime.Parse("09-Sep-9999 09:00"),
//                        Turbidity = 999.999
//                    });
//                }
//                continue;//will not fall through
//            }
//        }
//        //look for header (assumes table on each page will always start with this header)
//        if (_textlines[i].ToString().Equals(headersig["Type I"]))
//        {
//            datamode = true;
//            if (header == null) header = _textlines[i].ToString().Split(separator);
//        }
//    }
//    return data.ToArray();
//}





//public static int endPage
//{
//    get { return _endpage; }
//    set { _endpage = value; }
//}
//public static int startPage
//{
//    get { return _startpage; }
//    set { _startpage = value; }
//}
//public static Rectangle2D parseRegion
//{
//    get { return _parseregion; }
//    set { _parseregion = value; }
//}
