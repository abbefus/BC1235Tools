using java.awt.geom;
using OfficeOpenXml;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.pdmodel.common;
using org.apache.pdfbox.pdmodel.graphics;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using static ABUtils.FileUtils;
using Size = System.Drawing.Size;

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
                    YIndex = Convert.ToInt32(line.Box.Bottom),
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
            int end = endpage + 1 ?? doc.getNumberOfPages();
            java.util.ArrayList pages = GetPageRange(ref doc, start, end);  //potentially same as cat.getPageNodes().getKids()
            PDFStripper stripper = new PDFStripper();
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


    // tabifying
    public enum PDFRowClass
    {
        header,
        data,
        delete,
        unknown
    }
    public class ClassifiedPDFRow : PDFRow, INotifyPropertyChanged
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
                float sumw = _textsets.Sum(x => x.Box.Width);
                AverageCharWidth = sumw / TotalChars;
                X = _textsets.Min(x => x.Box.X);
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
    public class PDFColumn
    {
        public float X { get; set; }
        public float Width { get; set; }
    }
    public class PDFRow
    {
        public float Y { get; set; }
        public float Height { get; set; }
    }
    public class PDFTable
    {
        public PDFColumn[] Columns { get; set; }
        public PDFRow[] Rows { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
    }


    // stripping
    public struct TextChar
    {
        public float Direction { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public RectangleF Box { get; set; }
        public float SpaceAfter { get; set; }
        public float SpaceWidth { get; set; }
        public char Char { get; set; }
        public string Font { get; set; }
        public float FontSize { get; set; }
    }   // contains a char and character metrics
    public class TextLine
    {
        public TextChar[] TextChars { get; set; }
        public float LineSpace { get; set; }
        public RectangleF Box { get; set; }

        public IOrderedEnumerable<TextSet> GenerateTextSets(double toleranceAdjustment = 0)
        {
            double tol = PDFMetrics.CHARSPACING_TOLERANCE - toleranceAdjustment;
            try
            {
                List<TextSet> setlist = new List<TextSet>();

                bool startofset = true;
                TextSet set = new TextSet();
                StringBuilder sb = new StringBuilder();
                RectangleF bbox = new RectangleF();
                int count = 0;

                for (int i = 0; i < TextChars.Length; i++)
                {
                    count++;
                    TextChar tc = TextChars[i];
                    sb.Append(tc.Char);
                    if (startofset)
                    {
                        set = new TextSet();
                        bbox.X = tc.Box.Left;
                        startofset = false;
                    }
                    if (tc.SpaceAfter > (tc.SpaceWidth + tol))
                    {
                        int test = i - count;
                        IEnumerable<TextChar> setchars = TextChars.Skip(i - count).Take(count);

                        bbox.Width = tc.Box.Right - bbox.Left;
                        bbox.Height = setchars.Max(x => x.Box.Height);
                        bbox.Y = setchars.Min(x => x.Box.Y);
                        set.Box = bbox;
                        set.Count = count;
                        set.Text = sb.ToString();
                        set.SpaceAfter = tc.SpaceAfter;
                        set.SpaceWidth = tc.SpaceWidth;
                        setlist.Add(set); //yield return seq

                        sb.Clear();
                        startofset = true;
                        count = 0;
                    }
                }
                TextChar lasttc = TextChars.LastOrDefault();
                bbox.Width = (lasttc.Box.X + lasttc.Box.Width) - bbox.X;
                bbox.Height = TextChars.Skip(TextChars.Length - count).Take(count).Max(c => c.Box.Height);
                bbox.Y = TextChars.Skip(TextChars.Length - count).Take(count).Min(c => c.Box.Y);
                set.Box = bbox;
                set.Count = count;
                set.Text = sb.ToString();
                set.SpaceAfter = lasttc.SpaceAfter;
                set.SpaceWidth = lasttc.SpaceWidth;
                setlist.Add(set);
                return setlist.OrderBy(x => x.Box.X);
            }
            catch
            {
                throw new Exception("TextChars is empty or not populated.");
            }
        }
        public IOrderedEnumerable<TextLine> Split(double toleranceAdjustment = 0)
        {
            double tol = PDFMetrics.CHARSPACING_TOLERANCE - toleranceAdjustment;
            try
            {
                List<TextLine> setlist = new List<TextLine>();

                bool startofset = true;
                TextLine line = new TextLine();
                List<TextChar> chars = new List<TextChar>();
                RectangleF bbox = new RectangleF();
                int count = 0;

                for (int i = 0; i < TextChars.Length; i++)
                {
                    count++;
                    TextChar tc = TextChars[i];
                    chars.Add(tc);
                    if (startofset)
                    {
                        line = new TextLine();
                        bbox.X = tc.Box.Left;
                        startofset = false;
                    }
                    if (tc.SpaceAfter > (tc.SpaceWidth + tol))
                    {
                        int test = i - count;
                        IEnumerable<TextChar> setchars = TextChars.Skip(i - count).Take(count);

                        bbox.Width = tc.Box.Right - bbox.Left;
                        bbox.Height = setchars.Max(x => x.Box.Height);
                        bbox.Y = setchars.Min(x => x.Box.Y);
                        line.Box = bbox;
                        line.TextChars = chars.ToArray();
                        setlist.Add(line);

                        chars.Clear();
                        startofset = true;
                        count = 0;
                    }
                }
                TextChar lasttc = TextChars.LastOrDefault();
                bbox.Width = (lasttc.Box.X + lasttc.Box.Width) - bbox.X;
                bbox.Height = TextChars.Skip(TextChars.Length - count).Take(count).Max(c => c.Box.Height);
                bbox.Y = TextChars.Skip(TextChars.Length - count).Take(count).Min(c => c.Box.Y);
                line.Box = bbox;
                line.TextChars = chars.ToArray();
                setlist.Add(line);
                return setlist.OrderBy(x => x.Box.X);
            }
            catch
            {
                throw new Exception("TextChars is empty or not populated.");
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
    }    // contains array of TextChar and line metrics
    public class TextSet
    {
        public RectangleF Box { get; set; }
        public int Count { get; set; }
        public string Text { get; set; }
        public float SpaceAfter { get; set; }
        public float SpaceWidth { get; set; }
        public int TrailingSpaces
        {
            get { return (int)Math.Round(SpaceAfter / SpaceWidth); }
        }
    }     // a contiguous set of characters and associated metrics

    public class StrippedPDF
    {
        public StrippedPDFPage[] Pages { get; set; }
    }       // contains an array of StippedPDFPages and pdf properties
    public class StrippedPDFPage
    {
        public StrippedPDFPage(int pagenumber)
        {
            PageNumber = pagenumber;
        }
        public int PageNumber { get; set; }
        public Size PageSize { get; set; }
        public TextLine[] TextLines { get; set; }
    }   // contains an array of TextLines and page metrics



    //TODO: MaskedStripper: Extend PDFTextStripperByArea instead of trying to make this do both

    //NOTES:
        //y position and fontsize are dependent on the characters within the string of text
        //for example, a 'pqyg' increases the 'adjusted' y position relative to an 'oweruaszxcvnm' because of the tails
        //and vice versa for 'tidfhklb'
        //note 'j' goes in both directions
        //caps are all the same 'ABCDEFGHIJKLMNOPRSTUVWXYZ' except 'Q'
        //punctuation creates another situation with ',;' and '"'
    public class PDFStripper : PDFTextStripper
    {
        private List<TextChar>[] tcPages;


        public PDFStripper() : base()
        {
        }

        public StrippedPDF stripPDF(java.util.List pages)
        {
            int pagecount = pages.size();
            tcPages = new List<TextChar>[pagecount+1];
            base.processPages(pages);
            
            StrippedPDF pdf = new StrippedPDF();
            List<StrippedPDFPage> pdfpages = new List<StrippedPDFPage>();
            for (int i = 1; i < pagecount; i++)
            {
                StrippedPDFPage pdfpage = new StrippedPDFPage(i);

                // sets page dimensions
                pdfpage.PageSize = new Size
                    {
                        Width = (int)((PDPage)pages.get(i-1)).getMediaBox().getWidth(),
                        Height = (int)((PDPage)pages.get(i-1)).getMediaBox().getHeight()
                    };
                
                // groups TextChars into their lines and converts them to TextLines
                IOrderedEnumerable<IGrouping<float,TextChar>> tclines = tcPages[pdfpage.PageNumber]
                    .GroupBy(tc => tc.Box.Bottom).OrderBy(grp => grp.Key);

                pdfpage.TextLines = BuildTextLines(tclines);
                pdfpages.Add(pdfpage);
            }
            pdf.Pages = pdfpages.ToArray();
            pages.clear();
            pages = null;

            return pdf;
        }

        protected override void processTextPosition(TextPosition tp)
        {
            PDGraphicsState gs = getGraphicsState();
            TextChar tc = BuildTextChar(tp, gs);
            tcPages[getCurrentPageNo()].Add(tc);
        }

        private static TextChar BuildTextChar(TextPosition tp, PDGraphicsState gstate)
        {
            if (tp.getCharacter().Length != 1) throw new Exception("Textposition does not contain 1 character.");

            TextChar tc = new TextChar();
            tc.Char = tp.getCharacter()[0];

            float h = (float)Math.Floor(tp.getHeightDir());
            tc.Box = new RectangleF
            (
                tp.getXDirAdj(),
                (float)Math.Round(tp.getYDirAdj(), 0, MidpointRounding.ToEven) - h, // adjusted Y to top
                tp.getWidthDirAdj(),
                h
            );
            
            tc.Direction = tp.getDir();
            tc.SpaceWidth = tp.getWidthOfSpace();

            tc.Font = tp.getFont().getBaseFont();
            tc.FontSize = tp.getFontSizeInPt();

            try
            {
                int[] flags = GetBits(tp.getFont().getFontDescriptor().getFlags());
                tc.IsBold = findBold(tp, flags, gstate);
                tc.IsItalic = findItalics(tp, flags);
            }
            catch { }

            return tc;
        }
        private TextLine[] BuildTextLines(IOrderedEnumerable<IGrouping<float,TextChar>> lines)
        {
            List<TextLine> TLList = new List<TextLine>();
            TextLine lineA = null;
            foreach (IGrouping<float,TextChar> line in lines)
            {
                // create new textline
                TextLine lineB = new TextLine
                {
                    TextChars = line.OrderBy(tc => tc.Box.Left).ToArray()
                };
                float w = lineB.TextChars.Last().Box.Right - lineB.TextChars.First().Box.Left;
                float h = lineB.TextChars.Max(x => x.Box.Height);
                lineB.Box = new RectangleF
                (
                    lineB.TextChars.First().Box.X,
                    line.Key - h,
                    w,
                    h
                );

                // set the SpaceAfter property (now that chars are ordered)
                for (int i = 0; i < lineB.TextChars.Length-1; i++)
                {
                    float x1 = lineB.TextChars[i].Box.Right;
                    float x2 = lineB.TextChars[i + 1].Box.Left;
                    lineB.TextChars[i].SpaceAfter = (float)Math.Round(x2 - x1,1);
                }

                // if there's a line above...and it overlaps...
                // merge them and move to next else add the above line as is to TLList
                if (lineA != null)
                {
                    lineA.LineSpace = lineB.Box.Bottom - lineA.Box.Bottom;
                    if (lineA.LineSpace < lineB.Box.Height)
                    {
                        lineA = MergeTextLines(lineA, lineB);
                        continue;
                    }
                    else
                    {
                        TLList.Add(lineA);
                    }
                }
                lineA = lineB;
            }

            //add the last remaining line to TLList
            TLList.Add(lineA);

            return TLList.ToArray();
        }
        private static TextLine MergeTextLines(TextLine tl1, TextLine tl2)
        {
            // first, create a spacechar with the same metrics as tl1 to separate tl1 and tl2
            TextChar sc = tl1.TextChars.Where(x => x.Char == ' ').FirstOrDefault();
            if (sc.Char != '|')
            {
                sc = tl1.TextChars.FirstOrDefault();
                sc.Char = '|';
                sc.Box = new RectangleF(sc.Box.X, sc.Box.Y, (int)Math.Round(sc.SpaceWidth), sc.Box.Height);
            }
            sc.SpaceAfter = 0;

            List<TextLine> tlList = tl1.Split().Concat(tl2.Split()).OrderBy(x => x.Box.Left).ToList();

            List<TextLine> finished = new List<TextLine>();

            int count = 0;

            while (true)
            {
                TextLine tl = tlList[count++];
                if (count == tlList.Count) break;

                List<TextChar> tc = new List<TextChar>();
                for (int i = count; i < tlList.Count; i++)
                {
                    switch (tl.Box.HorizontalIntersect(tlList[i].Box))
                    {
                        case LineIntersectType.DoesNotIntersect:
                            finished.Add(new TextLine())
                        case LineIntersectType.Contains:
                            tc.AddRange(tl.TextChars.Concat(Enumerable.Repeat(sc, 1)).Concat(tlList[i].TextChars));

                    }
                }
            }


            


            tl1.Box = RectangleF.Union(tl1.Box, tl2.Box);

            if (tl1.Box.Bottom == tl2.Box.Bottom) tl1.LineSpace = tl2.LineSpace;


            return tl1;
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
        public static readonly short LINELIMIT = 6;
        public static readonly Size LETTER_SIZE = new Size(612, 792);

        public static bool TestWithinRange(float number, float target, int limit)
        {
            return number + limit > target && number - limit < target;
        }
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
//}   // main processor

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
