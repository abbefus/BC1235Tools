using ABUtils;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace PDFToExcel
{
    class PDFTableExtractDialog : OKCancelDialog
    {
        public int StartPage { get; set; }
        public int EndPage { get; set; }
        public int NumColumns { get; set; }
        public PDFTableExtractDialog(PageRange range) : base ("Extract PDF Table")
        {
            Width = 300;
            Height = Double.NaN;
            SizeToContent = SizeToContent.WidthAndHeight;
            AddComboBoxes(range.StartPage, range.EndPage);
        }

        private void AddComboBoxes(int start, int end)
        {
            DockPanel startpage_dp = new DockPanel
            {
                LastChildFill = false,
                Margin = new Thickness(5, 5, 10, 5),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            TextBlock startLabel_tb = new TextBlock
            {
                Text = "Table starts at page:",
                Margin = new Thickness(10)
            };
            ComboBox start_cb = new ComboBox
            {
                Width = 50,
                Height = 25,
                IsReadOnly = false,
                IsEditable = false,
            };
            Binding startBinding = new Binding
            {
                Source = this,
                Path = new PropertyPath("StartPage"),
                Mode = BindingMode.OneWayToSource,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                NotifyOnTargetUpdated = true,
                Converter = new GenericConverter()
            };
            start_cb.SetBinding(ComboBox.SelectedValueProperty, startBinding);

            DockPanel endpage_dp = new DockPanel
            {
                LastChildFill = false,
                Margin = new Thickness(5,5,10,5),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            TextBlock endLabel_tb = new TextBlock
            {
                Text = "Table ends at page:",
                Margin = new Thickness(10)
            };
            ComboBox end_cb = new ComboBox
            {
                Width = 50,
                Height = 25,
                IsReadOnly = false,
                IsEditable = false,
            };
            Binding endBinding = new Binding
            {
                Source = this,
                Path = new PropertyPath("EndPage"),
                Mode = BindingMode.OneWayToSource,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                NotifyOnTargetUpdated = true,
                Converter = new GenericConverter()
            };
            end_cb.SetBinding(ComboBox.SelectedValueProperty, endBinding);

            DockPanel colpage_dp = new DockPanel
            {
                LastChildFill = false,
                Margin = new Thickness(5, 5, 10, 5),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            TextBlock colLabel_tb = new TextBlock
            {
                Text = "Number of columns:",
                Margin = new Thickness(10)
            };
            ComboBox col_cb = new ComboBox
            {
                Width = 50,
                Height = 25,
                IsReadOnly = false,
                IsEditable = false,
            };
            Binding colBinding = new Binding
            {
                Source = this,
                Path = new PropertyPath("NumColumns"),
                Mode = BindingMode.OneWayToSource,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                NotifyOnTargetUpdated = true,
                Converter = new GenericConverter()
            };
            col_cb.SetBinding(ComboBox.SelectedValueProperty, colBinding);

            for (int i = start; i <= end; i++)
            {
                start_cb.Items.Add(i);
                end_cb.Items.Add(i);
            }
            for (int i = 2; i <= 20; i++)
            {
                col_cb.Items.Add(i);
            }


            start_cb.SelectedIndex = 0;
            end_cb.SelectedIndex = end_cb.Items.Count - 1;
            col_cb.SelectedIndex = 0;


            startpage_dp.Children.Add(startLabel_tb);
            startpage_dp.Children.Add(start_cb);
            endpage_dp.Children.Add(endLabel_tb);
            endpage_dp.Children.Add(end_cb);
            colpage_dp.Children.Add(colLabel_tb);
            colpage_dp.Children.Add(col_cb);


            stack_pnl.Children.Add(startpage_dp);
            stack_pnl.Children.Add(endpage_dp);
            stack_pnl.Children.Add(colpage_dp);
        }

    }
    public class GenericConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null) return value;
            return DependencyProperty.UnsetValue;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null) return value;
            return DependencyProperty.UnsetValue;
        }
    }
}
