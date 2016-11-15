using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using Button = System.Windows.Controls.Button;
using TextBlock = System.Windows.Controls.TextBlock;

namespace ABUtils
{

    public class SecuredWindow : Window
    {
        protected int DaysLeft { get; set; }

        public SecuredWindow(string appfolder = null, string expirydate = null)
        {
            string title = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            if (!string.IsNullOrWhiteSpace(appfolder))
            {
                if (!CheckLocation(appfolder, title)) Application.Current.Shutdown();
            }
            if (!string.IsNullOrWhiteSpace(expirydate))
            {
                if (!CheckExpiry(expirydate, title)) Application.Current.Shutdown();
            }
        }

        private bool CheckLocation(string appfolder, string title)
        {
            string location = Pathing.GetUNCPath(System.Reflection.Assembly.GetExecutingAssembly().Location);

            if (appfolder.ToLowerInvariant() != location.ToLowerInvariant())
            {
                ABModalDialog dialog = new ABModalDialog("DO NOT COPY OR MOVE THIS APPLICATION", false);
                dialog.SizeToContent = SizeToContent.WidthAndHeight;

                TextBlock blk = new TextBlock
                {
                    Text = string.Format("You are attempting to run {0} from {1}.\r\n{2} can only be run from {3}.",
                                            title, Path.GetDirectoryName(location), title, Path.GetDirectoryName(appfolder)),
                    Width = double.NaN,
                    Height = double.NaN,
                    TextAlignment = TextAlignment.Left,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(10),
                    TextWrapping = TextWrapping.Wrap
                };
                Button button = new Button
                {
                    Content = "OK",
                    Width = 100,
                    Height = 30,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(10),
                    IsCancel = true
                };

                dialog.stack_pnl.Children.Add(blk);
                dialog.stack_pnl.Children.Add(button);

                dialog.ShowDialog();
                return false;
            }
            return true;
        }
        private bool CheckExpiry(string expirydate, string title)
        {
            DateTime dt = DateTime.Parse(expirydate);
            if (DateTime.Today > dt)
            {
                ABModalDialog dialog = new ABModalDialog("Version Expired", false);
                dialog.Height = 150;
                dialog.SizeToContent = SizeToContent.Width;

                TextBlock tb = new TextBlock
                {
                    Text = string.Format("This version of {0} expired on\r\n{1}.", title, dt.ToLongDateString()),
                    Width = double.NaN,
                    Height = double.NaN,
                    TextAlignment = TextAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(10)
                };

                Button button = new Button
                {
                    Content = "OK",
                    Width = 100,
                    Height = 30,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(10),
                    IsCancel = true
                };

                dialog.stack_pnl.Children.Add(tb);
                dialog.stack_pnl.Children.Add(button);

                dialog.ShowDialog();
                return false;
            }
            else
            {
                DaysLeft = dt.Subtract(DateTime.Today).Days;
                return true;
            }
        }
    }
    public static class ExtensionMethods
    {
        public static string ReplaceAll(this string seed, char[] chars, string replacement = "")
        {
            return chars.Aggregate(seed, (x, c) => x.Replace(c.ToString(), replacement));
        }
        public static string FirstCharToUpper(this string str)
        {
            if (string.IsNullOrWhiteSpace(str)) return str;
            return char.ToUpper(str[0]) + str.Substring(1);
        }
    }
    public class WindowPositionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return System.Convert.ToDouble(value) > 20 || System.Convert.ToDouble(value) < -200;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new Exception("Never Convert Back");
        }

    }
}
