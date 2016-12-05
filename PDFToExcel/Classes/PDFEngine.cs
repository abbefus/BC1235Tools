using java.awt.geom;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.pdmodel.graphics;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using Size = System.Drawing.Size;

namespace PDFToExcel
{
    public class PDFEngine : IDisposable
    {

        //BUGS:
        //  1. Noticed some data still being classified as unknown in Logger_Unnamed - find out which page and what is causing the mistake
        //  2. Need to keep track of order and put the data back in that order when it is exported

        //NICE-TO-HAVES:
        //  1. also user specify if multi-page table
        //  2. prompt user to verify extracted header info instead of assuming
        //  3. handle cases where no data matching the specified number of columns is found
        //          by running through, gathering data and then prompting user with missing column (greatest VDivBuffer or first/last)
        //          by lowering the number of columns to the max found and proceeding
        //          by adjusting the tolerance so that a column separated by less than the default tolerance is found
        //  4. set PageNumber to actual in the pdf when done using it
        //  5. have table.Box.Height automatically adjust to column.Box.Height
        //  6. allow user to view completed table before export

        //NOTES:
        //  1. table.VDivBuffers are two values representing the last known max extent of the current column and the min extent of the next column
        //          VDivBuffer for a given column always represents the fuzzy boundary on the right side of the column
        //          this can be used to determine a fuzzy boundary within which non-columnized values can be assigned a column

        private PDDocument _doc;
        public PDDocument Doc
        {
            get { return _doc; }
            set
            {
                _doc = value;
                if (_doc.isEncrypted())
                {
                    if (!DecryptPDF())
                    {
                        Dispose();
                    }
                    else
                    {
                        Pages = new PageRange(Doc.getNumberOfPages());
                    }
                }
                else
                {
                    Pages = new PageRange(Doc.getNumberOfPages());
                }
            }
        }
        public PageRange Pages { get; private set; }

        public void OpenPDF(string pdfpath)
        {
            if (Doc != null) Dispose();
            Doc = PDDocument.load(pdfpath);
        }
        public bool CheckDoc()
        {
            if (Doc != null && Pages != null) return true;
            MessageBox.Show("There is no PDF document loaded.", "No PDF", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        
        public PDFTable TabifyPDF(int numcolumns=1)
        {
            if (!CheckDoc()) return null;

            //turn user-intuitive pages (1 - inclusive) into pdfbox page range (0 - exclusive)
            StrippedPDF pdf = StripPDF(Pages.StartPage - 1, Pages.EndPage);
            pdf.ConcatenateAllPages();

            // group data by the number of columns found (NumColumns found using spaceafter > spacewidth)
            IEnumerable<IGrouping<int, TextLine>> grps = pdf.Pages[0].TextLines.GroupBy(x => x.NumColumns());

            // these keys represent the difference between the expected number of columns and the number found for the row
            // note if more than the expected number of columns is present, these will be added to the 0 group
            int[] keys = grps.Select(grp => numcolumns - Math.Min(numcolumns,grp.Key)).OrderBy(x => x).ToArray();


            // build and classify rows
            List<PDFRow> rows = new List<PDFRow>();
            string headerstr = "";
            RectangleF box = new RectangleF();
            for (int i = 0; i < keys.Length; i++)
            {
                IGrouping<int, TextLine> group = grps.Where(grp => grp != null && numcolumns - Math.Min(numcolumns, grp.Key) == keys[i]).FirstOrDefault();

                foreach (IEnumerable<TextLine> line in group.Select(x => x.Split().Take(numcolumns))) // split by SpaceAfter and truncate extra columns
                {
                    PDFRow row = new PDFRow();
                    foreach (TextLine tl in line) tl.Trim();
                    row.TextLines = line.OrderBy(x => x.Box.Left).ToArray(); // make sure ordered left to right

                    #region Assign Row Class and Header, set table.Box
                    if (i == 0)
                    {
                        row.ColumnValues = row.TextLines;
                        if (keys[i] == 0 && string.IsNullOrWhiteSpace(headerstr))
                        {
                            // This might be where you prompt the user to verify whether this row is the header******
                            headerstr = row.ToString();
                            row.RowClass = PDFRowClass.header;
                            box = row.Box;
                        }
                        else
                        {
                            row.RowClass = row.ToString() == headerstr ? PDFRowClass.delete : PDFRowClass.data;
                            float xmin = Math.Min(box.X, row.Box.X);
                            float ymin = Math.Min(box.Y, row.Box.Y);
                            box.X = xmin;
                            box.Y = ymin;
                            box.Width = Math.Max(box.Right - xmin, row.Box.Right - xmin);
                            box.Height = Math.Max(box.Bottom - ymin, row.Box.Bottom - ymin);
                        }
                    }
                    else
                    {
                        row.ColumnValues = new TextLine[numcolumns];
                        row.RowClass = PDFRowClass.unknown;
                    }
                    #endregion

                    rows.Add(row);
                }
            }

            PDFTable table = new PDFTable(numcolumns);
            
            table.Box = box;
            table.Rows = rows.ToArray();

            ProfileRows(ref table);

            // assign rows to table, sort them and assign an index
            table.Rows = table.Rows.OrderBy(row => row.Box.Y).ToArray();
            for (int i = 0; i < table.Rows.Length; i++)
            {
                table.Rows[i].Index = i;
            }

            return table;
        }

        private static void ProfileRows(ref PDFTable table)
        {
            bool headerAdjusted = false;
            PDFRow headerRow = table.Rows.Where(x => x.RowClass == PDFRowClass.header).FirstOrDefault();
            if (headerRow != null) headerAdjusted = 
                    AdjustColumnsUsingHeader(ref table, headerRow.TextLines.Select(x => x.Box).ToArray());

            ColumnizeRows(ref table, headerAdjusted);
        }

        private static bool AdjustColumnsUsingHeader(ref PDFTable table, RectangleF[] headerBox)
        {
            try
            {
                for (int col = 0; col < table.Columns.Length; col++)
                {
                    float minx = headerBox[col].Left;
                    float miny = table.Box.Top;
                    table.Columns[col].Box = new RectangleF
                        (
                            minx,
                            miny,
                            headerBox[col].Right - minx,
                            table.Box.Height
                        );
                    if (col != table.Columns.Length - 1)
                    {
                        table.VDivBuffers[col, 0] = table.Columns[col].Box.Right;
                        table.VDivBuffers[col, 1] = headerBox[col + 1].Left;
                    }
                }
            }
            catch
            {
                return false;
            }
            return true;           
        }
        private static void ColumnizeRows(ref PDFTable table, bool headerAdjusted)
        {
            int numcolumns = table.Columns.Length;
            PDFRow[] datarows = table.Rows.Where(x => x.RowClass == PDFRowClass.data).ToArray();

            RectangleF[] headerBoxes = headerAdjusted ?
                table.Rows.Where(x => x.RowClass == PDFRowClass.header).FirstOrDefault().TextLines.Select(x => x.Box).ToArray() : null;

            if (datarows.Length > 0) // can assume every column represented
            {

                #region At Least One Full Row Exists

                for (int col = 0; col < numcolumns; col++)
                {
                    IEnumerable<RectangleF> cbox = datarows.Select(row => row.TextLines[col].Box);

                    float minx = headerAdjusted ?
                        Math.Min(table.Columns[col].Box.Left, cbox.Min(x => x.Left)) :  // take into account header could be farther left than data
                        cbox.Min(x => x.Left);

                    table.Columns[col].Box = new RectangleF
                        (
                            minx,
                            table.Box.Y,
                            Math.Max(table.Columns[col].Box.Right, cbox.Max(x => x.Right)) - minx,
                            table.Box.Height
                        );

                    if (headerAdjusted &&
                        cbox.All(x => Math.Round(x.X + x.Width / 2) == Math.Round(headerBoxes[col].X + headerBoxes[col].Width / 2)))
                    {
                        table.Columns[col].HorizontalAlignment = HorizontalAlignment.Center;
                    }
                    else if (cbox.All(x => Math.Round(x.X + x.Width / 2) == Math.Round(cbox.First().X + cbox.First().Width / 2)))
                    {
                        table.Columns[col].HorizontalAlignment = HorizontalAlignment.Center;
                    }
                    else if (cbox.All(x => Math.Round(x.X) == Math.Round(cbox.First().X)))
                    {
                        table.Columns[col].HorizontalAlignment = HorizontalAlignment.Left;
                    }
                    else if (cbox.All(x => Math.Round(x.Right) == Math.Round(cbox.First().Right)))
                    {
                        table.Columns[col].HorizontalAlignment = HorizontalAlignment.Right;
                    }
                    
                    else
                    {
                        table.Columns[col].HorizontalAlignment = HorizontalAlignment.Stretch;
                    }
                }

                HorizontalAlignment[] haligns = table.Columns.Select(x => x.HorizontalAlignment).ToArray();

                // refine VDivBuffers with datarows
                for (int col = 1; col < numcolumns - 1; col++)
                {
                    IEnumerable<RectangleF> col1box = datarows.Select(row => row.TextLines[col].Box);
                    IEnumerable<RectangleF> col2box = datarows.Select(row => row.TextLines[col + 1].Box);
                    if (haligns[col] == HorizontalAlignment.Left)
                    {
                        if (headerAdjusted)
                        {
                            table.VDivBuffers[col, 0] = Math.Max(table.VDivBuffers[col, 0], col1box.Max(x => x.Right));
                            table.VDivBuffers[col, 1] = Math.Min(table.VDivBuffers[col, 1], col2box.Min(x => x.Left));
                        }
                        else
                        {
                            table.VDivBuffers[col, 0] = col1box.Max(x => x.Right);
                            table.VDivBuffers[col, 1] = col2box.Min(x => x.Left);
                        }
                        
                    }
                }

                RectangleF firstrow = datarows.First().Box;
                // assign columns to unknown values
                // catch: unknown value might not be assigned if first column is aligned center or stretch and extends farther than both header and all known data
                foreach (PDFRow row in table.Rows.Where(x => x.RowClass == PDFRowClass.unknown))
                {
                    for (int col = 0; col < numcolumns; col++)
                    {
                        float xlimitRight = col < numcolumns - 1 ? table.VDivBuffers[col, 1] : float.MaxValue;    // farthest possible limit to the right for this column
                        float xlimitLeft = col > 0 ? table.VDivBuffers[col - 1, 0] : firstrow.Left; // farthest possible limit to the left for this column
                        if (haligns[col] == HorizontalAlignment.Left)
                        {
                            row.ColumnValues[col] = row.TextLines
                                .Where(x => Math.Round(x.Box.Left) == Math.Round(firstrow.Left) &&
                                            x.Box.Right < xlimitRight)
                                .FirstOrDefault();
                        }
                        else if (haligns[col] == HorizontalAlignment.Right)
                        {
                            row.ColumnValues[col] = row.TextLines
                                .Where(x => Math.Round(x.Box.Left) == Math.Round(firstrow.Left) &&
                                            x.Box.Right < xlimitRight)
                                .FirstOrDefault();
                        }
                        else if (haligns[col] == HorizontalAlignment.Center)
                        {
                            row.ColumnValues[col] = row.TextLines
                                .Where(x => Math.Round(x.Box.Left + x.Box.Width / 2) == Math.Round(firstrow.Left + firstrow.Width / 2) &&
                                            x.Box.Right < xlimitRight && x.Box.Left > xlimitLeft) // catch
                                .FirstOrDefault();
                        }
                        else
                        {
                            row.ColumnValues[col] = row.TextLines
                                .Where(x => x.Box.Right < xlimitRight && x.Box.Left > xlimitLeft).FirstOrDefault(); // catch
                        }
                    }
                }
                #endregion

            }
            else
            {
                ColumnizeUnknownRows(ref table, headerAdjusted);
                AlignFromRows(ref table, headerBoxes);
            }        
        }
        public static void ColumnizeUnknownRows(ref PDFTable table, bool headerAdjusted)
        {
            IEnumerable<PDFRow> unknownrows = table.Rows.Where(x => x.RowClass == PDFRowClass.unknown);
            if (headerAdjusted) // no data but header adjusted
            {
                foreach (PDFRow row in unknownrows)
                {
                    if (row.Box.Top < table.Box.Top) continue;
                    for (int i = 0; i < row.TextLines.Length; i++)
                    {
                        PDFColumn[] cols = table.Columns
                            .Where(c => c.Box.IntersectsHorizontallyWith(row.TextLines[i].Box)).ToArray(); // if textline intesects column
                        if (cols.Length != 1)
                        {
                            row.ColumnValues = new TextLine[table.Columns.Length];
                            row.RowClass = PDFRowClass.delete;
                            break;
                        }// each TextLine must intersect 1 column else delete row and move on

                        PDFColumn col = cols.First();
                        if (col.Index == 0)
                        {
                            if (row.TextLines[i].Box.Right < table.VDivBuffers[col.Index, 1])
                            {
                                row.ColumnValues[col.Index] = row.TextLines[i];
                            }
                        }
                        else if (col.Index != table.Columns.Length - 1)
                        {
                            if (row.TextLines[i].Box.Right < table.VDivBuffers[col.Index, 1] && 
                                row.TextLines[i].Box.Left > table.VDivBuffers[col.Index - 1,0])
                            {
                                row.RowClass = PDFRowClass.data;
                                row.ColumnValues[col.Index] = row.TextLines[i];
                                table.VDivBuffers[col.Index, 0] = Math.Max(table.VDivBuffers[col.Index, 0], row.TextLines[i].Box.Right); //adjust inside buffer to match new data
                                table.VDivBuffers[col.Index - 1, 1] = Math.Min(table.VDivBuffers[col.Index - 1, 1], row.TextLines[i].Box.Left); //adjust outside buffer to match new data
                            }
                        }
                        else
                        {
                            if (row.TextLines[i].Box.Left > table.VDivBuffers[col.Index - 1, 0])
                            {
                                row.RowClass = PDFRowClass.data;
                                row.ColumnValues[col.Index] = row.TextLines[i];
                                table.VDivBuffers[col.Index - 1, 1] = Math.Min(table.VDivBuffers[col.Index - 1, 1], row.TextLines[i].Box.Left); //adjust outside buffer to match new data
                            }
                        }  
                    }
                }

                RectangleF[] rowBoxes = table.Rows.Where(x => x.RowClass == PDFRowClass.data).Select(row => row.Box).ToArray();
                float xmin = Math.Min(table.Box.Left, rowBoxes.Min(x => x.Left));
                float ymin = Math.Min(table.Box.Top, rowBoxes.Min(x => x.Top));
                table.Box = new RectangleF
                    (
                        xmin,
                        ymin,
                        Math.Max(table.Box.Right - xmin, rowBoxes.Max(x => x.Right) - xmin),
                        Math.Max(table.Box.Bottom - ymin, rowBoxes.Max(x => x.Bottom) - ymin)
                    );

                foreach (PDFColumn column in table.Columns)
                {
                    xmin = column.Index == 0 ? table.Box.Left : table.VDivBuffers[column.Index - 1,1];
                    ymin = column.Box.Top;

                    column.Box = new RectangleF
                        (
                            xmin,
                            ymin,
                            column.Index == table.Columns.Length - 1 ? table.Box.Right - xmin : table.VDivBuffers[column.Index,0] - xmin,
                            table.Box.Height
                        );
                }
            }
            else // no header and no data -- need to assume columns line up along an axis
            {

                #region No Header and No Data
                MessageBox.Show(string.Format("Cannot find any data or headers, unable to\r\npiece together an incomplete table"),
                "No Data Found",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
                return;
                // try to create table around unknown values if no data
                //if (box.IsEmpty)
                //{
                //    IEnumerable<SizeF> unknownsizes = rows.Where(row => row.RowClass == PDFRowClass.unknown).Select(row => row.Box.Size);
                //    if (box.Width == 0) box.Width = unknownsizes.Max(size => size.Width);
                //    if (box.Height == 0) box.Height = unknownsizes.Max(size => size.Height);
                //}
                //if (box.Location.IsEmpty)
                //{
                //    IEnumerable<PointF> unknownlocations = rows.Where(row => row.RowClass == PDFRowClass.unknown).Select(loc => loc.Box.Location);
                //    box.X = unknownlocations.Max(loc => loc.X);
                //    box.Y = unknownlocations.Max(loc => loc.Y);
                //}
                #endregion

            }
        }
        // Note: Will set unaligned rows back to unknown if they don't match alignment
        public static void AlignFromRows(ref PDFTable table, RectangleF[] headerBoxes)
        {
            IEnumerable<PDFRow> rows = table.Rows.Where(x => x.RowClass == PDFRowClass.data);
            //IEnumerable<PDFRow> headerrow = table.Rows.Where(x => x.RowClass == PDFRowClass.header);
            for (int col = 0; col < table.Columns.Length; col++)
            {
                table.Columns[col].HorizontalAlignment = HorizontalAlignment.Stretch;
                IEnumerable<PDFRow> memberrows = rows.Where(x => x.ColumnValues[col] != null);
                if (!memberrows.Any()) continue;
                IEnumerable<IGrouping<int, PDFRow>> centeraligned = memberrows
                    .GroupBy(x => (int)Math.Round(x.ColumnValues[col].Box.Left + (x.ColumnValues[col].Box.Width / 2)));
                int centermax = centeraligned.Max(x => x.Count());
                if (centeraligned.Count() == 1)
                {
                    table.Columns[col].HorizontalAlignment = HorizontalAlignment.Center;
                    continue;
                }
                else // if it matches a header
                {
                    
                    int centerline = centeraligned.Where(x => x.Count() == centermax).First().Key;
                    if ((int)Math.Round(headerBoxes[col].Left + (headerBoxes[col].Width / 2)) == centerline)
                    {
                        if (centermax > 1)
                        {
                            IEnumerable<IGrouping<int,PDFRow>> others = centeraligned.Where(x => x.Key != centerline);
                            foreach (IGrouping<int,PDFRow> other in others)
                            {
                                foreach (PDFRow row in other) row.RowClass = PDFRowClass.unknown;
                            }
                        }
                        
                        table.Columns[col].HorizontalAlignment = HorizontalAlignment.Center;
                        continue;
                    }
                }

                IEnumerable<IGrouping<int, PDFRow>> leftaligned = memberrows
                    .GroupBy(x => (int)Math.Round(x.ColumnValues[col].Box.Left));
                int leftmax = leftaligned.Max(x => x.Count());
                if (leftaligned.Count() == 1)
                {
                    table.Columns[col].HorizontalAlignment = HorizontalAlignment.Left;
                    continue;
                }
                else // if it matches a header
                {
                    
                    int leftline = leftaligned.Where(x => x.Count() == leftmax).First().Key;
                    if ((int)Math.Round(headerBoxes[col].Left) == leftline)
                    {
                        if (leftmax > 1)
                        {
                            IEnumerable<IGrouping<int, PDFRow>> others = centeraligned.Where(x => x.Key != leftline);
                            foreach (IGrouping<int, PDFRow> other in others)
                            {
                                foreach (PDFRow row in other) row.RowClass = PDFRowClass.unknown;
                            }
                        }
                        table.Columns[col].HorizontalAlignment = HorizontalAlignment.Left;
                        continue;
                    }
                }

                IEnumerable<IGrouping<int, PDFRow>> rightaligned = memberrows
                    .GroupBy(x => (int)Math.Round(x.ColumnValues[col].Box.Right));
                int rightmax = rightaligned.Max(x => x.Count());
                if (rightaligned.Count() == 1)
                {
                    table.Columns[col].HorizontalAlignment = HorizontalAlignment.Right;
                    continue;
                }

                if (centermax > leftmax)
                {
                    if (centermax > rightmax)
                    {
                        table.Columns[col].HorizontalAlignment = HorizontalAlignment.Center;
                    }
                }


            }
        }
        

        public bool DecryptPDF()
        {
            if (Doc.isEncrypted())
            {
                try
                {
                    Doc.decrypt("");
                    Doc.setAllSecurityToBeRemoved(true);
                    return true;
                }
                catch (Exception e)
                {
                    MessageBox.Show(
                        string.Format("Unable to decrypt this document.\r\n{0}", e.Message),
                        "Unable to Decrypt",
                        MessageBoxButton.OK,
                        MessageBoxImage.Exclamation);
                    return false;
                }
            }
            else
            {
                return true;
            }
        }
        private StrippedPDF StripPDF(int? startpage = null, int? endpage = null, Rectangle2D parseregion = null)
        {
            int start = startpage ?? 0;
            int end = endpage ?? Doc.getNumberOfPages();
            java.util.ArrayList pages = GetPageRange(start, end);  //potentially same as cat.getPageNodes().getKids()
            PDFStripper stripper = new PDFStripper();
            return stripper.stripPDF(pages);
        }
        private java.util.ArrayList GetPageRange(int start, int end)
        {
            java.util.ArrayList pages = new java.util.ArrayList();

            try
            {
                PDDocumentCatalog cat = Doc.getDocumentCatalog();
                java.util.List catpages = cat.getAllPages();

                for (int i = start; i < end; i++)
                {
                    pages.add(catpages.get(i));
                }

                return pages;
            }
            catch
            {
                throw new ArgumentOutOfRangeException("Page range not found in document.");
            }
            
        }

        public void Dispose()
        {
            Doc.close();
            Pages = null;
        }
    }


    // tabifying
    public class PDFRow : INotifyPropertyChanged
    {
        public RectangleF Box { get; private set; }
        private PDFRowClass _rowClass;
        public PDFRowClass RowClass
        {
            get { return _rowClass; }
            set
            {
                _rowClass = value;
                NotifyPropertyChanged();
            }
        }
        private TextLine[] _textlines;
        public TextLine[] TextLines
        {
            get { return _textlines; }
            set
            {
                _textlines = value;
                float minx = TextLines.First().Box.Left;
                float miny = TextLines.Min(x => x.Box.Top);
                Box = new RectangleF
                (
                    minx,
                    miny,
                    TextLines.Last().Box.Right - minx,
                    TextLines.Max(x => x.Box.Bottom) - miny
                );
            }
        }
        public int Index { get; set; }
        public TextLine[] ColumnValues { get; set; }

        public override string ToString()
        {
            if (ColumnValues.Any(x => x != null))
            {
                return string.Join("\t", ColumnValues.Select(x => x != null ? x.ToString() : ""));
            }
            else
            {
                IEnumerable<string> tlstrings = TextLines.Select(x => x.ToString());
                return string.Join("\t", tlstrings);
            }
            
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
        header = 1,
        data = 2,
        delete = 3,
        unknown = 4
    }
    public class PDFColumn
    {
        public PDFColumn(int index)
        {
            Index = index;
        }
        public RectangleF Box { get; set; }
        public int Index { get; set; }
        public HorizontalAlignment HorizontalAlignment { get; set; }
    }
    public class PDFTable
    {
        public PDFTable(int numcolumns)
        {
            Columns = new PDFColumn[numcolumns];
            VDivisions = new float[numcolumns + 1];
            VDivBuffers = new float[numcolumns - 1,2];
            for (int i=0;i<numcolumns;i++)
            {
                Columns[i] = new PDFColumn(i);
            }
        }
        public RectangleF Box { get; set; }
        public float[] VDivisions { get; set; }
        public float[,] VDivBuffers { get; set; }
        public PDFColumn[] Columns { get; private set; }
        public PDFRow[] Rows { get; set; }

        private float GetBuffer(int col)
        {
            return VDivBuffers[col, 1] - VDivBuffers[col, 0];
        }
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

        public void ShiftChar(float offsetX=0, float offsetY=0)
        {
            Box = new RectangleF
                (
                    Box.X + offsetX,
                    Box.Y + offsetY,
                    Box.Width,
                    Box.Height
                );
        }// shifts the position of the character
    }   // contains a char and character metrics
    public class TextLine
    {
        private TextChar[] _textChars;
        public TextChar[] TextChars
        {
            get { return _textChars; }
            set
            {
                _textChars = value;
                float minx = TextChars.First().Box.Left;
                float miny = TextChars.Min(tc => tc.Box.Top);
                Box = new RectangleF
                (
                    minx,
                    miny,
                    TextChars.Last().Box.Right - minx,
                    TextChars.Max(x => x.Box.Bottom) - miny
                );
            }
        }
        public float LineSpace { get; set; }
        public RectangleF Box { get; private set; }

        public void Trim()
        {
            int take = TextChars.Length;
            int skip = 0;
            while (take > 0 && TextChars[take - 1].Char == ' ') take--;
            while (take > 1 && skip < (take - 1) && TextChars[skip].Char == ' ') { skip++; take--; }
            TextChars = TextChars.Skip(skip).Take(take).ToArray();
            TextChars[TextChars.Length - 1].SpaceAfter = 0;
        }
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
                List<TextLine> textLineList = new List<TextLine>();

                TextLine line = new TextLine { LineSpace = LineSpace };
                List<TextChar> textCharList = new List<TextChar>();

                for (int i = 0; i < TextChars.Length; i++)
                {
                    TextChar tc = TextChars[i];
                    textCharList.Add(tc);
                    if (tc.SpaceAfter > (tc.SpaceWidth + tol))
                    {
                        line.TextChars = textCharList.ToArray();
                        textLineList.Add(line);

                        textCharList.Clear();
                        line = new TextLine { LineSpace = LineSpace };
                    }
                }
                if (textCharList.Count > 0)
                {
                    TextChar lasttc = TextChars.LastOrDefault();
                    line.TextChars = textCharList.ToArray();
                    textLineList.Add(line);
                }
                return textLineList.OrderBy(x => x.Box.X);
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
        public string ToString(bool append=true)
        {
            StringBuilder sb = new StringBuilder();
            foreach (TextChar tc in TextChars)
            {
                sb.Append(tc.Char);
                if (append && tc.SpaceAfter > (tc.SpaceWidth + PDFMetrics.CHARSPACING_TOLERANCE))
                {
                    int numspaces = (int)Math.Floor(tc.SpaceAfter / tc.SpaceWidth);
                    sb.Append(' ', numspaces);
                }
            }
            return sb.ToString();
        }

        public void ShiftLine(float offsetX=0, float offsetY=0)
        {
            for (int i = 0; i < _textChars.Length; i++)
            {
                _textChars[i].ShiftChar(offsetX, offsetY);
            }
            Box = new RectangleF
            (
                TextChars.First().Box.Left,
                TextChars.Min(tc => tc.Box.Top),
                Box.Width,
                Box.Height
            );
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

    public class PageRange
    {
        private int upperLimit = int.MaxValue;
        private int lowerLimit = 1;
        public int StartPage { get; private set; }
        public int EndPage { get; private set; }

        public PageRange(int numberOfPages, int startPage=1)
        {
            StartPage = startPage;
            EndPage = numberOfPages;
            upperLimit = numberOfPages;
            lowerLimit = startPage;
        }
        
        public bool SetStartPage(int start)
        {
            bool success = false;
            if (ContainsPage(start) && EndPage >= start)
            {
                StartPage = start;
                success = true;
            }
            return success;
        }
        public bool SetEndPage(int end)
        {
            bool success = false;
            if (ContainsPage(end) && end >= StartPage)
            {
                EndPage = end;
                success = true;
            }
            return success;
        }
        public bool SetPageRange(int start, int end)
        {
            bool success = false;
            if (ContainsPage(start) && ContainsPage(end) && end >= start)
            {
                StartPage = start;
                EndPage = end;
                success = true;
            }
            return success;
        }

        public bool ContainsPage(int value)
        {
            return lowerLimit <= value && upperLimit >= value;
        }
        public override string ToString()
        {
            if (StartPage == EndPage)
            {
                return string.Format("page {0}", StartPage);
            }
            else
            {
                return string.Format("pages {0}-{1}", StartPage, EndPage);
            }
            
        }
    }
    public class StrippedPDF
    {
        public StrippedPDFPage[] Pages { get; set; }

        public void ConcatenateAllPages()
        {
            if (Pages.Length < 2) return;
            Size pagesize = new Size
            {
                Width = Pages.Max(x => x.PageSize.Width),
                Height = 0
            };

            for (int pg = 0; pg < Pages.Length; pg++)
            {
                if (pg != 0)
                {
                    foreach (TextLine tl in Pages[pg].TextLines)
                    {
                        tl.ShiftLine(0, pagesize.Height);
                    }
                }
                pagesize.Height += Pages[pg].PageSize.Height;
            }

            StrippedPDFPage pageone = new StrippedPDFPage(1);
            pageone.TextLines = Pages.Select(x => x.TextLines)
                                    .Aggregate(new List<TextLine>(), (a, b) => a.Concat(b).ToList())
                                    .ToArray();
            pageone.PageSize = pagesize;
            Pages = new StrippedPDFPage[1] { pageone };
        }
    }       // contains an array of StrippedPDFPages and pdf properties
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



    //TODO: PDFStripperByArea: Extend PDFTextStripperByArea instead

    //NOTES:
        //y position and fontsize are dependent on the characters within the string of text
            //doesn't seem to do this for non-ocr documents
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
            if (tp.getCharacter().Length != 1)
            {
                throw new Exception("Textposition does not contain 1 character.");
            }

            TextChar tc = new TextChar();
            tc.Char = tp.getCharacter()[0];

            float h = tp.getHeightDir();

            tc.Box = new RectangleF
            (
                tp.getXDirAdj(),
                tp.getYDirAdj() - h, // adjusted Y to top
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

        //TODO:
        //  1. Create recursion to merge additional lines - at moment, only bottom line is merged, other lines are left unmerged (1 merge only)
        private TextLine[] BuildTextLines(IOrderedEnumerable<IGrouping<float,TextChar>> lines)
        {
            List<TextLine> TLList = new List<TextLine>();
            List<TextLine> MergeQueue = new List<TextLine>();
            TextLine lineAbove = null;

            foreach (IGrouping<float,TextChar> line in lines)
            {
                // create new textline
                TextLine lineBelow = new TextLine
                {
                    TextChars = line.OrderBy(tc => tc.Box.Left).ToArray()
                };
                

                // set the SpaceAfter property (now that chars are ordered)
                for (int i = 0; i < lineBelow.TextChars.Length-1; i++)
                {
                    float x1 = lineBelow.TextChars[i].Box.Right;
                    float x2 = lineBelow.TextChars[i + 1].Box.Left;
                    lineBelow.TextChars[i].SpaceAfter = (float)Math.Round(x2 - x1,1);
                }

                if (lineAbove != null) // if the lineAbove has been added...
                {
                    lineAbove.LineSpace = lineBelow.Box.Bottom - lineAbove.Box.Bottom;
                    if (MergeQueue.Count == 0) // if there are no queued merges
                    {
                        if (lineAbove.LineSpace <= lineBelow.Box.Height + 1) // if overlaps add lineAbove to queue
                        {
                            MergeQueue.AddRange(lineAbove.Split());
                        }
                        else // add non-overlapping lineAbove
                        {
                            TLList.Add(lineAbove);
                        }
                    }
                    else // if there are queued merges, add line above to queue if it overlaps
                    {
                        if (MergeQueue.Last().LineSpace <= lineAbove.Box.Height + 1)
                        {
                            MergeQueue.AddRange(lineAbove.Split());
                        }
                        else // else merge and then add the non-overlapping lineAbove
                        {
                            TLList.AddRange(MergeLines(MergeQueue));
                            MergeQueue.Clear();
                            if (lineAbove.LineSpace <= lineBelow.Box.Height + 1) // if overlaps add lineAbove to queue
                            {
                                MergeQueue.AddRange(lineAbove.Split());
                            }
                            else // add non-overlapping lineAbove
                            {
                                TLList.Add(lineAbove);
                            }
                        }
                    }
                }
                lineAbove = lineBelow;
            }
            //add the last remaining line to TLList
            if (MergeQueue.Count == 0)
            {
                TLList.Add(lineAbove);
            }
            else if (MergeQueue.Last().LineSpace <= lineAbove.Box.Height + 1)
            {
                MergeQueue.AddRange(lineAbove.Split());
                TLList.AddRange(MergeLines(MergeQueue));
                MergeQueue.Clear();
            }
            else
            {
                TLList.AddRange(MergeLines(MergeQueue));
                MergeQueue.Clear();
                TLList.Add(lineAbove);
            }
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

            // overlapping text is split into separate textlines by bottom value
            foreach (IGrouping<float,TextLine> grps in overlap.GroupBy(x => x.Box.Bottom).OrderBy(grp => grp.Key))
            {
                TextLine groupedTL = new TextLine();
                groupedTL.TextChars = grps.Select(x => x.TextChars).
                    Aggregate(new List<TextChar>(), (a, b) => a.Concat(b).ToList()).OrderBy(x => x.Box.Left).ToArray();
                groupedTL.LineSpace = grps.Max(x => x.LineSpace);
                yield return groupedTL;
            }

            // non-overlapping text is merged together on one line, SpaceAfter and LineSpace is modified, TextChar.Box is untouched
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
