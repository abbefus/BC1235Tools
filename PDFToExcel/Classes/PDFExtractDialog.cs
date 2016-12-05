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
        public PDFTableExtractDialog(PageRange range, string title= "Extract PDF Table") : base (title)
        {
            Width = 500;
            Height = Double.NaN;
            SizeToContent = SizeToContent.Height;
            AddComboBoxes(range.StartPage, range.EndPage);

            ColumnDefinition cd = new ColumnDefinition
            {
                Width = new GridLength(250)
            };
            grid.ColumnDefinitions.Add(cd);
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            Grid.SetColumn(ButtonDP, 1);
        }

        private void AddComboBoxes(int start, int end)
        {
            StackPanel labels_sp = new StackPanel();
            Grid.SetColumn(labels_sp, 0);
            TextBlock startLabel_tb = new TextBlock
            {
                Text = "Table starts at page:",
                Margin = new Thickness(10, 13, 0, 10),
                Height = 25,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            TextBlock endLabel_tb = new TextBlock
            {
                Text = "Table ends at page:",
                Margin = new Thickness(10, 10, 0, 10),
                Height = 25,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            TextBlock colLabel_tb = new TextBlock
            {
                Text = "Number of columns:",
                Margin = new Thickness(10, 10, 0, 10),
                Height = 25,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            StackPanel combox_sp = new StackPanel();
            Grid.SetColumn(combox_sp, 1);

            ComboBox start_cb = new ComboBox
            {
                Width = 50,
                Height = 25,
                IsReadOnly = false,
                IsEditable = false,
                Margin = new Thickness(5, 10, 10, 10),
                HorizontalAlignment = HorizontalAlignment.Left,
                IsTabStop = true,
                TabIndex = 0
            };
            Binding startBinding = new Binding
            {
                Source = this,
                Path = new PropertyPath("StartPage"),
                Mode = BindingMode.OneWayToSource,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                NotifyOnTargetUpdated = true,
            };
            start_cb.SetBinding(ComboBox.SelectedValueProperty, startBinding);

            ComboBox end_cb = new ComboBox
            {
                Width = 50,
                Height = 25,
                IsReadOnly = false,
                IsEditable = false,
                Margin = new Thickness(5,10,10,10),
                HorizontalAlignment = HorizontalAlignment.Left,
                IsTabStop = true,
                TabIndex = 1
            };
            Binding endBinding = new Binding
            {
                Source = this,
                Path = new PropertyPath("EndPage"),
                Mode = BindingMode.OneWayToSource,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                NotifyOnTargetUpdated = true,
            };
            end_cb.SetBinding(ComboBox.SelectedValueProperty, endBinding);

            ComboBox col_cb = new ComboBox
            {
                Width = 50,
                Height = 25,
                IsReadOnly = false,
                IsEditable = false,
                Margin = new Thickness(5,10,10,10),
                HorizontalAlignment = HorizontalAlignment.Left,
                IsTabStop = true,
                TabIndex = 2
            };
            Binding colBinding = new Binding
            {
                Source = this,
                Path = new PropertyPath("NumColumns"),
                Mode = BindingMode.OneWayToSource,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                NotifyOnTargetUpdated = true,
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


            combox_sp.Children.Add(start_cb);
            combox_sp.Children.Add(end_cb);
            combox_sp.Children.Add(col_cb);

            labels_sp.Children.Add(startLabel_tb);
            labels_sp.Children.Add(endLabel_tb);
            labels_sp.Children.Add(colLabel_tb);

            grid.Children.Add(combox_sp);
            grid.Children.Add(labels_sp);
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
