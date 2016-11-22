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
        public static PDFRowLite[] TabifyPDF(string pdfpath, int numcolumns=1, int? startpage=null, int? endpage=null, string headerstring=null)
        {
            PDDocument doc = PDDocument.load(pdfpath);
            DecryptPDF(ref doc);

            // turn user-intuitive pages (1-inclusive) into PDFBox intuitive pages (0-exclusive)
            startpage = startpage > 0 ? startpage - 1 : null;
            endpage = endpage >= startpage ? endpage : null;

            StrippedPDF pdf = StripPDF(ref doc, startpage, endpage);

            doc.close();

            PDFTable table = new PDFTable(numcolumns, pdf.Pages.FirstOrDefault().PageSize.Width);

            TextLine[] alltextlines = pdf.Pages.Select(x => x.TextLines).Aggregate(new List<TextLine>(), (a, b) => a.Concat(b).ToList()).ToArray();

            IEnumerable<IGrouping<int, TextLine>> grps = alltextlines.GroupBy(x => x.NumColumns());
            TextLine[] data = grps.Where(grp => grp.Key == numcolumns).SelectMany(x => x).ToArray();

            table.X = data.Min(x => x.Box.Left);
            table.Width = data.Max(x => x.Box.Right) - table.X;

            PDFRowLite[] rows = new PDFRowLite[data.Length];
            
            if (data != null)
            {
                string headerstr = data.First().ToString();
                for (int i = 0; i < rows.Length; i++)
                {
                    rows[i] = new PDFRowLite();
                    rows[i].TextLines = data[i].Split().ToArray();
                    rows[i].Index = i + 1;
                    if (i == 0 || rows[i].ToString() == headerstr)
                    {
                        rows[i].LineType = PDFRowClass.header;
                    }
                    else
                    {
                        rows[i].LineType = PDFRowClass.data;
                    }
                }
            }

            return rows;
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
    public class PDFRowLite : PDFRow, INotifyPropertyChanged
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
        public TextLine[] TextLines { get; set; }
        public int Index { get; set; }


        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (TextLine ts in TextLines)
            {
                sb.Append(ts.ToString()).Append(' ', (int)Math.Round(ts.TextChars.Last().SpaceAfter / ts.TextChars.Last().SpaceWidth));
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
                Box = Box.SetX(_textsets.Min(x => x.Box.X));
            }
        }
        public int TotalChars { get; private set; }
        public float AverageCharWidth { get; private set; }
        public int PageNumber { get; set; }
        public int YIndex { get; set; }
        public int Index { get; set; }

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
        public RectangleF Box { get; set; }
    }
    public class PDFRow
    {
        public RectangleF Box { get; set; }
    }
    public class PDFTable
    {
        public PDFTable(int numcolumns, int width)
        {
            Columns = new PDFColumn[numcolumns];
            _width = width;
            for (int i=0;i<numcolumns;i++)
            {
                Columns[i] = new PDFColumn();
            }
        }
        public PDFColumn[] Columns { get; private set; }
        public PDFRow[] Rows { get; set; }
        private float _x;
        public float X
        {
            get { return _x; }
            set
            {
                float offset = value - _x;
                _x = value;
                Array.ForEach(Columns, x => x.Box = x.Box.SetX(x.Box.X + offset));
            }
        }
        public float Y { get; set; }
        private float _width;
        public float Width
        {
            get { return _width; }
            set
            {
                double fraction = value / _width;
                _width = value;
                Array.ForEach(Columns, x => x.Box = x.Box.SetWidth((float)(x.Box.Width * fraction)));
            }
        }
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

        public int NumColumns(double toleranceAdjustment = 0)
        {
            double tol = PDFMetrics.CHARSPACING_TOLERANCE - toleranceAdjustment;
            return TextChars.Where(tc => (tc.SpaceAfter > (tc.SpaceWidth + tol))).Count() + 1;
        }
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
                    TextChar tc = TextChars[i];
                    chars.Add(tc);
                    if (startofset)
                    {
                        line = new TextLine();
                        bbox.X = tc.Box.Left;
                        startofset = false;
                    }
                    if (tc.SpaceAfter > (tc.SpaceWidth + tol) && count > 0)
                    {
                        IEnumerable<TextChar> setchars = TextChars.Skip(i - count).Take(count);
                        bbox.Width = tc.Box.Right - bbox.Left;
                        bbox.Height = setchars.Max(x => x.Box.Height); //setchars is empty
                        bbox.Y = setchars.Min(x => x.Box.Y);
                        line.Box = bbox;
                        line.TextChars = chars.ToArray();
                        setlist.Add(line);

                        chars.Clear();
                        startofset = true;
                        count = 0;
                        continue;
                    }
                    count++;
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
            for (int i = 1; i < pagecount+1; i++)
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
            int currentPageNo = getCurrentPageNo();
            if (tcPages.ElementAtOrDefault(currentPageNo) == null)
            {
                tcPages[currentPageNo] = new List<TextChar>();
            }
            tcPages[currentPageNo].Add(tc);
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
            List<TextLine> MergeQueue = new List<TextLine>();
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

                // if there's a line above...and it overlaps vertically...
                // add them to merge queue
                if (lineA != null)
                {
                    lineA.LineSpace = lineB.Box.Bottom - lineA.Box.Bottom;
                    if (lineA.LineSpace <= lineB.Box.Height + 1) // +1 is a threshold value - distance between lines to consider merging
                    {
                        if (MergeQueue.Count == 0) MergeQueue.AddRange(lineA.Split());
                        MergeQueue.AddRange(lineB.Split());
                    }
                    else if (MergeQueue.Count > 0) // no overlap but have to write and clear merge queue before moving on
                    {
                        TLList.AddRange(MergeLines(MergeQueue));
                        MergeQueue.Clear();
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

        private static IEnumerable<TextLine> MergeLines(IEnumerable<TextLine> mergelist)
        {
            TextLine[] ordered = mergelist.OrderBy(x => x.Box.Left).ToArray();
            List<TextLine> overlap = new List<TextLine>();
            List<TextLine> nooverlap = new List<TextLine>(); 

            for (int i = 0; i < ordered.Length; i++)
            {
                if (i == ordered.Length - 1)
                {
                    nooverlap.Add(ordered[i]);
                    break;
                }
                if (!ordered[i].Box.IntersectsHorizontallyWith(ordered[i+1].Box))
                {
                    nooverlap.Add(ordered[i]);
                    continue;
                }
                else
                {
                    IEnumerable<TextLine> yOrder = new TextLine[] { ordered[i], ordered[i + 1] }.OrderBy(x => x.Box.Bottom);
                    overlap.Add(yOrder.First());
                    switch (yOrder.First().Box.HorizontalIntersect(yOrder.Last().Box))
                    {
                        case LineIntersectType.Contains:    //then tmp.First() is ordered[i]
                            nooverlap.Add(yOrder.Last());
                            i++;                        //skip ordered[i+1]
                            continue;
                        case LineIntersectType.Within:      //then tmp.First() is ordered[i+1]
                            ordered[i] = yOrder.First();
                            ordered[i + 1] = yOrder.Last();     //need to evaluate this one again
                            continue;
                        case LineIntersectType.ContainsStart:   //then tmp.First() is ordered[i]
                            continue;
                        case LineIntersectType.ContainsEnd:     //then tmp.First() is ordered[i+1]
                            nooverlap.Add(yOrder.Last());
                            i++;                        //skip ordered[i+1]
                            continue;
                    }
                }
            }
            foreach (IGrouping<float,TextLine> grps in overlap.GroupBy(x => x.Box.Bottom).OrderBy(grp => grp.Key))
            {
                TextLine groupedTL = new TextLine();
                groupedTL.TextChars = grps.Select(x => x.TextChars).
                    Aggregate(new List<TextChar>(), (a, b) => a.Concat(b).ToList()).OrderBy(x => x.Box.Left).ToArray();
                groupedTL.LineSpace = grps.Max(x => x.LineSpace);
                groupedTL.Box = grps.Select(x => x.Box).Aggregate(new RectangleF(), (a, b) => RectangleF.Union(a, b));
                yield return groupedTL;
            }

            TextLine[] lines = nooverlap.OrderBy(x => x.Box.Left).ToArray();
            for (int i = 0; i < lines.Length - 1; i++)
            {
                lines[i].TextChars[lines[i].TextChars.Length - 1].SpaceAfter =
                    lines[i + 1].TextChars.First().Box.Left - lines[i].TextChars.Last().Box.Right;
            }
            lines[lines.Length - 1].TextChars[lines[lines.Length - 1].TextChars.Length - 1].SpaceAfter = 0;

            TextLine mergedTL = new TextLine();
            mergedTL.TextChars = lines.Select(x => x.TextChars).
                Aggregate(new List<TextChar>(), (a, b) => a.Concat(b).ToList()).OrderBy(x => x.Box.Left).ToArray();
            mergedTL.LineSpace = lines.Max(x => x.LineSpace);
            mergedTL.Box = lines.Select(x => x.Box).Aggregate(new RectangleF(), (a, b) => RectangleF.Union(a, b));
            yield return mergedTL;
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
