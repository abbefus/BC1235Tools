using ABUtils;
using Microsoft.Win32;
using Microsoft.Windows.Controls.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using RectangleF = System.Drawing.RectangleF;
using PointF = System.Drawing.PointF;

namespace PDFToExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class PDF2DataMain : SecuredWindow
    {
        ObservableCollection<PDFRow> PDFTextLines { get; set; }


        public PDF2DataMain()
            : base
            (
#if DEBUG
                    @"C:\Users\abefus\Documents\Visual Studio 2015\Projects\BC1235Tools\PDFToExcel\bin\Debug\PDFToData.exe", //fixed folder location
                    "31-DEC-2016"  //expiry date
#elif FINAL
                  @"\\CD1002-F03\GEOMATICS\Utilities\GIS\PDFToData.exe", //fixed folder location
                    "31-DEC-2016"  //expiry date
#elif RELEASE
                  @"C:\Users\abefus\Documents\Visual Studio 2015\Projects\BC1235Tools\PDFToExcel\bin\Release\PDFToData.exe", //fixed folder location
                    "31-DEC-2016"  //expiry date

#endif
            )
        {
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            grid.DataContext = this;
            PDFTextLines = new ObservableCollection<PDFRow>();
            datagrid.DataContext = PDFTextLines;
            InitializeStaticComboBoxes();
        }

        private void InitializeStaticComboBoxes()
        {
            for (int i=2; i<=20;i++)
            {
                numcolumns_rgc.Items.Add(new RibbonGalleryItem { Content = i.ToString() });
            }
            numcolumns_rg.SelectedValue = ((RibbonGalleryItem)numcolumns_rgc.Items[0]).Content;
        }

        private void openpdf_btn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog(this) == true)
            {
                // SHOW DIALOG FOR SPECIFYING COLUMNS AND PAGE RANGE HERE


                PDFTextLines.Clear();

                int start = string.IsNullOrWhiteSpace(startpage_tb.Text) ? 0 : int.Parse(startpage_tb.Text);
                int end = string.IsNullOrWhiteSpace(startpage_tb.Text) ? 0 : int.Parse(endpage_tb.Text);

                PDFTable pdftable = PDFEngine.TabifyPDF(openFileDialog.FileName, int.Parse(numcolumns_rg.SelectedValue.ToString()), start, end);
                if (pdftable != null)
                {
                    foreach (PDFRow row in pdftable.Rows)
                    {
                        PDFTextLines.Add(row);
                    }
                    if (PDFTextLines.Count > 0)
                    {
                        UpdateStatus(StatusType.Success, string.Format("Processed {0} pages, found {1} lines.",
                        start, end,
                        PDFTextLines.Count()));
                        return;
                    }
                }
                UpdateStatus(StatusType.Failure, "Error or no data in pdf.");
            }
        }


        public enum StatusType
        {
            Success,
            Failure
        }
        public void UpdateStatus(StatusType type, string msg)
        {
            status_tb.Foreground = type == StatusType.Failure ?
                new SolidColorBrush(Colors.Red) :
                new SolidColorBrush(Colors.Green);
            Console.WriteLine(msg);
        }

        #region Ribbon

        private void exit_btn_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        // switches content in mainframe
        private void Ribbon_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }
        private void ribbon_Loaded(object sender, RoutedEventArgs e)
        {
            // removes quick action toolbar (styling)
            Grid child = VisualTreeHelper.GetChild((DependencyObject)sender, 0) as Grid;
            if (child != null)
            {
                child.RowDefinitions[0].Height = new GridLength(0);
            }

            foreach (RibbonGallery rg in FindVisualChildren<RibbonGallery>(pdf2data_rbn))
            {
                rg.Command = ApplicationCommands.NotACommand;
                rg.Command = null;
            }


            Console.SetOut(new ConsolWriter(status_tb));
            UpdateStatus(StatusType.Success, string.Format("This version of the '{0}' application will expire in {1} days.", Title, DaysLeft));
        }
        private void RibbonApplicationMenu_Loaded(object sender, RoutedEventArgs e)
        {
            // removes 'recent' column in application menu (styling)
            RibbonApplicationMenu am = sender as RibbonApplicationMenu;
            Grid grid = (am.Template.FindName("MainPaneBorder", am) as Border).Parent as Grid;
            grid.ColumnDefinitions[2].Width = new GridLength(0);
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }


        #endregion

        private void page_tb_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = Regex.IsMatch(e.Text, "[^0-9]+");
        }

        private void savexls_btn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Save To Excel";
            saveFileDialog.AddExtension = true;
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.Filter = "XLSX Files|*.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                using (OpenWorkbook owb = new OpenWorkbook(saveFileDialog.FileName))
                {
                    if (owb.AddWorksheet("PDF Export"))
                    {
                        int row = 1;
                        if (includehdr_chk.IsChecked ?? true)
                        {
                            //FUCKING NOPE! --------------------------------------------------------------------------------------------
                            owb.UpdateRow(row++, PDFTextLines.Where(x => x.RowClass == PDFRowClass.header).First().ToString().Split('\t'));
                        }
                        foreach (PDFRow classedpdf in PDFTextLines.Where(x => x.RowClass == PDFRowClass.data))
                        {
                            owb.UpdateRow(row++, classedpdf.ToString().Split('\t'));
                        }
                        owb.ActiveWorksheet.Cells.AutoFitColumns();
                        owb.Save();
                        UpdateStatus(StatusType.Success, string.Format("Table Saved. ({0})", System.IO.Path.GetFileName(saveFileDialog.FileName)));
                    }
                }
            }
        }

        private void setdelete_btn_Click(object sender, RoutedEventArgs e)
        {
            foreach (PDFRow pdfTL in datagrid.SelectedItems)
            {
                PDFTextLines[PDFTextLines.IndexOf(pdfTL)].RowClass = PDFRowClass.delete;
            }
        }

        private void setheader_btn_Click(object sender, RoutedEventArgs e)
        {
            foreach (PDFRow pdfTL in datagrid.SelectedItems)
            {
                PDFTextLines[PDFTextLines.IndexOf(pdfTL)].RowClass = PDFRowClass.header;
            }
        }

        private void setdata_btn_Click(object sender, RoutedEventArgs e)
        {
            foreach(PDFRow pdfTL in datagrid.SelectedItems)
            {
                PDFTextLines[PDFTextLines.IndexOf(pdfTL)].RowClass = PDFRowClass.data;
            }
        }

        private void purgedeleted_btn_Click(object sender, RoutedEventArgs e)
        {
            PDFTextLines.RemoveAll<PDFRow>(x => x.RowClass == PDFRowClass.delete);
        }
    }
    public class ConsolWriter : TextWriter
    {
        private TextBlock textblock;
        public ConsolWriter(TextBlock textbox)
        {
            textblock = textbox;
        }
        public override void Write(string value)
        {
            textblock.Text = value;
        }
        public override void WriteLine(string value)
        {
            textblock.Text = value;
        }

        public override Encoding Encoding
        {
            get { return Encoding.ASCII; }
        }
    }


    public static class ExtensionMethods
    {
        public static int RemoveAll<T>(
            this ObservableCollection<T> coll, Func<T, bool> condition)
        {
            var itemsToRemove = coll.Where(condition).ToList();

            foreach (var itemToRemove in itemsToRemove)
            {
                coll.Remove(itemToRemove);
            }

            return itemsToRemove.Count;
        }

        public static LineIntersectType HorizontalIntersect(this RectangleF rect, RectangleF other)
        {
            if (rect.Left < other.Left && rect.Right > other.Right) return LineIntersectType.Contains;
            if (rect.Left > other.Left && rect.Right < other.Right) return LineIntersectType.Within;
            if (rect.Left < other.Left && rect.Right > other.Left) return LineIntersectType.ContainsStart;
            if (rect.Left < other.Right && rect.Right > other.Right) return LineIntersectType.ContainsEnd;

            return LineIntersectType.DoesNotIntersect;
        }
        public static bool IntersectsHorizontallyWith(this RectangleF rect, RectangleF other)
        {
            return !(rect.Right < other.Left || rect.Left > other.Right);
        }

        public static RectangleF SetX(this RectangleF rect, float x)
        {
            return new RectangleF(x, rect.Y, rect.Width, rect.Height);
        }
        public static RectangleF SetY(this RectangleF rect, float y)
        {
            return new RectangleF(rect.X, y, rect.Width, rect.Height);
        }
        public static RectangleF SetWidth(this RectangleF rect, float width)
        {
            return new RectangleF(rect.X, rect.Y, width, rect.Height);
        }
        public static RectangleF SetHeight(this RectangleF rect, float height)
        {
            return new RectangleF(rect.X, rect.Y, rect.Width, height);
        }
    }
    public enum LineIntersectType
    {
        Contains,
        Within,
        ContainsStart,
        ContainsEnd,
        DoesNotIntersect
    }
    public class LineToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch (value.ToString())
            {
                case "header":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#000"));
                case "data":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#999"));
                case "delete":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E00"));
                case "unknown":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DDD"));
            }


            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(value.ToString()));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
    public class IntegerToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (int)value > 0;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }

}
