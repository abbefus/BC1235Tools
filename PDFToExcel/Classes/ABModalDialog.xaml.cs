using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;

namespace ABUtils

{
    /// <summary>
    /// Interaction logic for ABModalDialog.xaml
    /// </summary>
    public partial class ABModalDialog : Window
    {
        private bool stopClose;

        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hwnd, int index);

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hwnd, IntPtr hwndInsertAfter, int x, int y, int width, int height, uint flags);

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

        const int GWL_EXSTYLE = -20;
        const int WS_EX_DLGMODALFRAME = 0x0001;
        const int SWP_NOSIZE = 0x0001;
        const int SWP_NOMOVE = 0x0002;
        const int SWP_NOZORDER = 0x0004;
        const int SWP_FRAMECHANGED = 0x0020;
        const uint WM_SETICON = 0x0080;
        const int WM_SYSCOMMAND = 0x0112;
        const int SC_MOVE = 0xF010;
        private const int GWL_STYLE = -16;
        private const int WS_SYSMENU = 0x80000;
        public ABModalDialog(string title = "Dialog", bool required = true)
        {
            InitializeComponent();
            SourceInitialized += Window_SourceInitialized;
            Title = title;
            stopClose = required;
            Loaded += Window_Loaded;
            Closing += Window_Closing;

        }
        private void Window_SourceInitialized(object sender, EventArgs e)
        {
            WindowInteropHelper helper = new WindowInteropHelper(this);
            IntPtr hwnd = helper.Handle;
            HwndSource source = HwndSource.FromHwnd(hwnd);
            source.AddHook(WndProc);

            // Change the extended window style to not show a window icon
            int extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle | WS_EX_DLGMODALFRAME);

            // Update the window's non-client area to reflect the changes
            SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE |
                  SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);

            SendMessage(hwnd, WM_SETICON, new IntPtr(1), IntPtr.Zero);
            SendMessage(hwnd, WM_SETICON, IntPtr.Zero, IntPtr.Zero);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {

            switch (msg)
            {
                case WM_SYSCOMMAND:
                    int command = wParam.ToInt32() & 0xfff0;
                    if (command == SC_MOVE)
                    {
                        handled = true;
                    }
                    break;
                default:
                    break;
            }
            return IntPtr.Zero;
        }

        protected void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = stopClose;
        }

        public bool Closeable
        {
            get { return !stopClose; }
            set { stopClose = !value; }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Application curApp = Application.Current;
            Window mainWindow = curApp.MainWindow;
            this.Left = mainWindow.Left + (mainWindow.ActualWidth - this.ActualWidth) / 2;
            this.Top = mainWindow.Top + (mainWindow.ActualHeight - this.ActualHeight) / 2;
            Loaded -= Window_Loaded;
        }
    }

    class CheckListDialog : ABModalDialog
    {
        public ListBox CheckList { get; set; }

        public CheckListDialog(string[] checklist,
                                string message = "Do something to the following items?",
                                string title = "Are you Suuuure...?", bool required = false) :
                                base(title, required)
        {
            Width = 300;
            Height = Double.NaN;
            SizeToContent = SizeToContent.Height;
            TextBlock message_tb = new TextBlock
            {
                Text = message,
                Margin = new Thickness(10),
                MaxWidth = 275,
                TextWrapping = TextWrapping.Wrap
            };

            CheckList = new ListBox
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = double.NaN,
                MinWidth = 200,
                MaxHeight = 200,
                ItemsSource = checklist,
                Margin = new Thickness(10),
                Background = new SolidColorBrush(Colors.LightGray)
            };

            stack_pnl.Children.Add(message_tb);
            stack_pnl.Children.Add(CheckList);

            AddOkCancelButtons();
        }

        private void AddOkCancelButtons()
        {
            DockPanel button_dp = new DockPanel
            {
                LastChildFill = false,
                Margin = new Thickness(10),
                Width = 150,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            Button okbtn = new Button
            {
                Content = "OK",
                Width = 60,
                Margin = new Thickness(2),
                IsDefault = true
            };
            Button cancelbtn = new Button
            {
                Content = "CANCEL",
                Width = 60,
                Margin = new Thickness(2),
                IsCancel = true
            };

            okbtn.Click += Okbtn_Click;

            button_dp.Children.Add(okbtn);
            button_dp.Children.Add(cancelbtn);
            DockPanel.SetDock(cancelbtn, Dock.Right);
            DockPanel.SetDock(okbtn, Dock.Right);

            stack_pnl.Children.Add(button_dp);
        }
        private void Okbtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }

    public static class Pathing
    {
        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int WNetGetConnection(
            [MarshalAs(UnmanagedType.LPTStr)] string localName,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder remoteName,
            ref int length);
        /// <summary>
        /// Given a path, returns the UNC path or the original. (No exceptions
        /// are raised by this function directly). For example, "P:\2008-02-29"
        /// might return: "\\networkserver\Shares\Photos\2008-02-09"
        /// </summary>
        /// <param name="originalPath">The path to convert to a UNC Path</param>
        /// <returns>A UNC path. If a network drive letter is specified, the
        /// drive letter is converted to a UNC or network path. If the 
        /// originalPath cannot be converted, it is returned unchanged.</returns>
        public static string GetUNCPath(string originalPath)
        {
            StringBuilder sb = new StringBuilder(512);
            int size = sb.Capacity;

            // look for the {LETTER}: combination ...
            if (originalPath.Length > 2 && originalPath[1] == ':')
            {
                // don't use char.IsLetter here - as that can be misleading
                // the only valid drive letters are a-z && A-Z.
                char c = originalPath[0];
                if ((c >= 'j' && c <= 'z') || (c >= 'J' && c <= 'Z')) //mapped drives should always be greater or equal to J
                {
                    int error = WNetGetConnection(originalPath.Substring(0, 2),
                        sb, ref size);
                    if (error == 0)
                    {
                        DirectoryInfo dir = new DirectoryInfo(originalPath);

                        string path = Path.GetFullPath(originalPath)
                            .Substring(Path.GetPathRoot(originalPath).Length);
                        return Path.Combine(sb.ToString().TrimEnd(), path);
                    }
                }
            }
            return originalPath;
        }
    }
}
