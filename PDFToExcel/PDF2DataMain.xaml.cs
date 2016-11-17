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
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PDFToExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class PDF2DataMain : SecuredWindow
    {
        ObservableCollection<PDFTextLine> PDFTextLines { get; set; }


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
            PDFTextLines = new ObservableCollection<PDFTextLine>();
            datagrid.DataContext = PDFTextLines;
            InitializeStaticComboBoxes();
        }

        private void InitializeStaticComboBoxes()
        {
            for (int i=1; i<=20;i++)
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
                PDFTextLines.Clear();

                int start = string.IsNullOrWhiteSpace(startpage_tb.Text) ? 0 : int.Parse(startpage_tb.Text);
                int end = string.IsNullOrWhiteSpace(startpage_tb.Text) ? 0 : int.Parse(endpage_tb.Text);

                // custom sort
                IEnumerable<PDFTextLine> tmp = PDFEngine.ClassifyPDF
                    (openFileDialog.FileName, 
                    int.Parse(numcolumns_rg.SelectedValue.ToString()),
                    start,
                    end);
                foreach (PDFTextLine headerdata in tmp.Where(x => x.LineType == PDFTableClass.header)) PDFTextLines.Add(headerdata);
                foreach (PDFTextLine otherdata in tmp.Where(x => x.LineType != PDFTableClass.header).OrderBy(x => x.LineType).ThenBy(x => x.Index))
                {
                    PDFTextLines.Add(otherdata);
                }


                if (PDFTextLines.Count > 0)
                {
                    UpdateStatus(StatusType.Success, string.Format("Processed {0} pages, found {1} lines.",
                    PDFTextLines.LastOrDefault().PageNumber - PDFTextLines.FirstOrDefault().PageNumber,
                    PDFTextLines.Count()));
                }
                else
                {
                    UpdateStatus(StatusType.Failure, "Error or no data in pdf.");
                }
                
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
                            owb.UpdateRow(row++, PDFTextLines.Where(x => x.LineType == PDFTableClass.header).FirstOrDefault().TextBlocks.Select(x => x.Text).ToArray());
                        }
                        IEnumerable<PDFTextLine> classedpdfs = PDFTextLines.Where(x => x.LineType == PDFTableClass.data);
                        foreach (PDFTextLine classedpdf in classedpdfs)
                        {
                            owb.UpdateRow(row++, classedpdf.TextBlocks.Select(x => x.Text).ToArray());
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
            foreach (PDFTextLine pdfTL in datagrid.SelectedItems)
            {
                PDFTextLines[PDFTextLines.IndexOf(pdfTL)].LineType = PDFTableClass.delete;
            }
        }

        private void setheader_btn_Click(object sender, RoutedEventArgs e)
        {
            foreach (PDFTextLine pdfTL in datagrid.SelectedItems)
            {
                PDFTextLines[PDFTextLines.IndexOf(pdfTL)].LineType = PDFTableClass.header;
            }
        }

        private void setdata_btn_Click(object sender, RoutedEventArgs e)
        {
            foreach(PDFTextLine pdfTL in datagrid.SelectedItems)
            {
                PDFTextLines[PDFTextLines.IndexOf(pdfTL)].LineType = PDFTableClass.data;
            }
        }

        private void purgedeleted_btn_Click(object sender, RoutedEventArgs e)
        {
            PDFTextLines.RemoveAll<PDFTextLine>(x => x.LineType == PDFTableClass.delete);
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

    public enum PDFTableClass
    {
        header,
        data,
        delete
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
    }
    public class LineToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch (value.ToString())
            {
                case "header":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF009BD3"));
                case "data":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#000"));
                case "delete":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E00"));
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
    //public class TextBlocks2Text : IValueConverter
    //{
    //    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        TextBlock[] textblocks = value as TextBlock[];
    //        StringBuilder sb = new StringBuilder();
    //        foreach
    //    }
    //    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        return DependencyProperty.UnsetValue;
    //    }
    //}
}
