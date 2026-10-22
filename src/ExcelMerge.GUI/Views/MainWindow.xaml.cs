using System.Windows;
using System.Windows.Controls;
using System;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Input;
using ExcelMerge.GUI.Shell;

namespace ExcelMerge.GUI.Views
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private GridLength previousConsoleHeight = new GridLength(0);

        public MainWindow()
        {
            InitializeComponent();
            SourceInitialized += (sender, args) => ApplyTitleBarTheme(App.Instance.Setting.DarkMode);
            DataContextChanged += (sender, args) => AttachApplicationMenu();

            var host = new PowerShellHost();
            Console.PowerShellHost = host;
            host.Open();
        }

        private void AttachApplicationMenu()
        {
            var model = DataContext as ExcelMerge.GUI.ViewModels.MainWindowViewModel;
            var view = model?.Content as DiffView;
            var host = view?.FindName("ApplicationMenuHost") as ContentControl;
            if (host == null) return;
            MainGrid.Children.Remove(ApplicationMenu);
            ApplicationMenu.DataContext = DataContext;
            host.Content = ApplicationMenu;
        }

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        internal int ApplyTitleBarTheme(bool dark)
        {
            var handle = new WindowInteropHelper(this).Handle;
            if (handle == IntPtr.Zero) return 0;
            try
            {
                int enabled = dark ? 1 : 0;
                int caption = dark ? 25 | (30 << 8) | (38 << 16) : -1;
                int text = dark ? 227 | (233 << 8) | (242 << 16) : -1;
                DwmSetWindowAttribute(handle, 20, ref enabled, sizeof(int));
                int captionResult = DwmSetWindowAttribute(handle, 35, ref caption, sizeof(int));
                int textResult = DwmSetWindowAttribute(handle, 36, ref text, sizeof(int));
                return captionResult != 0 ? captionResult : textResult;
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }
            return -1;
        }

        private void MenuItem_Loaded(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem == null)
                return;

            var binding = menuItem.GetBindingExpression(MenuItem.IsEnabledProperty);
            if (binding == null)
                return;

            binding.UpdateTarget();
        }

        private void ConsoleVisibilityChanged(object sender, RoutedEventArgs e)
        {
            if (Console.Visibility == Visibility.Collapsed)
                ShowConsole();
            else
                HideConsole();
        }

        private void ShowConsole()
        {
            Console.Visibility = Visibility.Visible;
            ConsoleGridSplitter.Visibility = Visibility.Visible;

            if (previousConsoleHeight.Value > 0)
            {
                MainGrid.RowDefinitions[3].Height = previousConsoleHeight;
            }
            else
            {
                MainGrid.RowDefinitions[3].Height = new GridLength(Height / 3d);
                previousConsoleHeight = MainGrid.RowDefinitions[3].Height;
            }
        }

        private void HideConsole()
        {
            Console.Visibility = Visibility.Collapsed;
            ConsoleGridSplitter.Visibility = Visibility.Collapsed;
            previousConsoleHeight = new GridLength(MainGrid.RowDefinitions[3].ActualHeight);
            MainGrid.RowDefinitions[3].Height = new GridLength(0);

            UpdateLayout();
        }

        public void WriteToConsole(string message)
        {
            ConsoleVisibilityMenuItem.IsChecked = true;
            ShowConsole();

            Console.Write(message);
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.D:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            if (Console.Visibility == Visibility.Collapsed)
                            {
                                ShowConsole();
                                Console.Focus();
                            }
                            else
                            {
                                HideConsole();
                            }

                            e.Handled = true;
                        }
                    }
                    break;
            }
        }
    }
}
