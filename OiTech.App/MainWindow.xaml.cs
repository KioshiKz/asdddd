using Microsoft.Win32;
using OiTech.App.Services;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace OiTech.App
{
    public partial class MainWindow : Window
    {
        private enum ActiveView { Home, Students, Templates, Reports }
        private ActiveView _currentView = ActiveView.Home;

        public MainWindow()
        {
            InitializeComponent();
            UpdateThemeButton();
            UpdateLangButton();
            LoadDashboard();
        }

        // ═══════════════════ THEME / LANGUAGE ═══════════════════

        private void ToggleTheme(object sender, RoutedEventArgs e)
        {
            ThemeManager.Toggle();
            UpdateThemeButton();
        }

        private void ToggleLang(object sender, RoutedEventArgs e)
        {
            LocalizationManager.Toggle();
            UpdateLangButton();
            UpdateThemeButton();
            RefreshCurrentView();
        }

        private void UpdateThemeButton()
        {
            if (ThemeIcon == null) return;
            if (ThemeManager.IsDark)
            {
                ThemeIcon.Text = ""; // brightness / sun
                ThemeLabel.Text = LocalizationManager.Get("ThemeToLight");
            }
            else
            {
                ThemeIcon.Text = ""; // quiet hours / moon
                ThemeLabel.Text = LocalizationManager.Get("ThemeToDark");
            }
        }

        private void UpdateLangButton()
        {
            if (LangLabel == null) return;
            LangLabel.Text = LocalizationManager.CurrentCode;
        }

        private void RefreshCurrentView()
        {
            switch (_currentView)
            {
                case ActiveView.Home:
                    LoadDashboard();
                    break;
                case ActiveView.Students:
                    MainContent.Content = new StudentsView();
                    break;
                case ActiveView.Templates:
                    MainContent.Content = new TemplateView();
                    break;
                case ActiveView.Reports:
                    MainContent.Content = new ReportsView();
                    break;
            }
        }

        // ═══════════════════ NAVIGATION ═══════════════════

        private void AnimateContent(UIElement element)
        {
            if (element is FrameworkElement fe)
            {
                var sb = fe.TryFindResource("FadeInAnimation") as System.Windows.Media.Animation.Storyboard;
                sb?.Begin(fe);
            }
        }

        private void SetContent(UserControl view, ActiveView viewId)
        {
            if (AppData.Students.Count == 0)
            {
                MessageBox.Show(LocalizationManager.Get("MsgNoData"),
                    LocalizationManager.Get("MsgNoDataTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            _currentView = viewId;
            DashboardPanel.Visibility = Visibility.Collapsed;
            MainContent.Visibility = Visibility.Visible;
            MainContent.Content = view;
            AnimateContent(MainContent);
        }

        private void Nav_Home(object sender, RoutedEventArgs e)
        {
            _currentView = ActiveView.Home;
            DashboardPanel.Visibility = Visibility.Visible;
            MainContent.Visibility = Visibility.Collapsed;
            MainContent.Content = null;
            LoadDashboard();
            AnimateContent(DashboardPanel);
        }

        private void Nav_Students(object sender, RoutedEventArgs e)
        {
            _currentView = ActiveView.Students;
            DashboardPanel.Visibility = Visibility.Collapsed;
            MainContent.Visibility = Visibility.Visible;
            MainContent.Content = new StudentsView();
            AnimateContent(MainContent);
        }
        private void Nav_Import(object sender, RoutedEventArgs e)       => DoImport();
        private void Nav_ImportWord(object sender, RoutedEventArgs e)   => DoImportWord();
        private void Nav_Template(object sender, RoutedEventArgs e)     => SetContent(new TemplateView(), ActiveView.Templates);
        private void Nav_Reports(object sender, RoutedEventArgs e)      => SetContent(new ReportsView(), ActiveView.Reports);

        // ═══════════════════ DASHBOARD ═══════════════════

        private void LoadDashboard()
        {
            if (StudentsCountText != null)
                StudentsCountText.Text = AppData.StudentsCount.ToString("N0");
            if (SubjectsCountText != null)
                SubjectsCountText.Text = AppData.SubjectNames.Count.ToString("N0");
            if (PrintedCountText != null)
                PrintedCountText.Text = AppData.PrintedCount.ToString("N0");

            if (ActionsPanel != null)
            {
                ActionsPanel.Children.Clear();
                if (AppData.Actions.Count == 0)
                {
                    ActionsPanel.Children.Add(new TextBlock
                    {
                        Text = LocalizationManager.Get("ActionLogEmpty"),
                        Foreground = (SolidColorBrush)FindResource("TextDimBrush"),
                        FontSize = 13,
                        Margin = new Thickness(20, 10, 20, 10)
                    });
                }
                else
                {
                    foreach (var a in AppData.Actions)
                        ActionsPanel.Children.Add(MakeActionRow(a));
                }
            }
        }

        private UIElement MakeActionRow(ActionEntry a)
        {
            var g = new Grid();
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var dot = new Ellipse
            {
                Width = 8, Height = 8,
                Fill = (SolidColorBrush)FindResource("AccentBlueBrush"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 12, 0)
            };
            Grid.SetColumn(dot, 0);
            g.Children.Add(dot);

            var txt = new TextBlock
            {
                Text = a.Text,
                Foreground = (SolidColorBrush)FindResource("TextWhiteBrush"),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(txt, 1);
            g.Children.Add(txt);

            var time = new TextBlock
            {
                Text = a.Time.ToString("HH:mm:ss"),
                Foreground = (SolidColorBrush)FindResource("TextDimBrush"),
                FontSize = 12,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0)
            };
            Grid.SetColumn(time, 2);
            g.Children.Add(time);

            return new Border
            {
                BorderBrush = (SolidColorBrush)FindResource("BorderBrush"),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(4, 10, 4, 10),
                Child = g
            };
        }

        // ═══════════════════ IMPORT ═══════════════════

        private void DoImportWord()
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Word (*.docx)|*.docx|Все файлы (*.*)|*.*",
                Title = "Выберите файл Word со списком студентов",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (ofd.ShowDialog() != true) return;

            try
            {
                var svc = new WordService();
                var students = svc.Import(ofd.FileName);

                if (students.Count == 0)
                {
                    MessageBox.Show("Студенты не найдены. Проверьте формат файла.",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                AppData.Students = students;
                AppData.StudentsCount = students.Count;
                AppData.SubjectNames.Clear();
                AppData.AddAction(
                    $"Импорт Word: {students.Count} студентов — {System.IO.Path.GetFileName(ofd.FileName)}");
                LoadDashboard();

                MessageBox.Show($"Импортировано: {students.Count} студентов.",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                Nav_Students(null!, null!);
            }
            catch (Exception ex)
            {
                AppData.AddAction($"Ошибка импорта Word: {ex.Message}");
                LoadDashboard();
                MessageBox.Show($"Ошибка импорта Word:\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void DoImport()
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                Title = "Выберите файл Excel",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (ofd.ShowDialog() != true) return;

            Cursor = Cursors.Wait;
            IsEnabled = false;

            try
            {
                var svc = new ExcelService();
                var students = await Task.Run(() => svc.Import(ofd.FileName));

                AppData.Students = students;
                AppData.StudentsCount = students.Count;
                AppData.AddAction(string.Format(
                    LocalizationManager.Get("ActionImported"),
                    students.Count, AppData.SubjectNames.Count,
                    System.IO.Path.GetFileName(ofd.FileName)));
                LoadDashboard();

                MessageBox.Show(
                    string.Format(LocalizationManager.Get("MsgImportSuccess"), students.Count, AppData.SubjectNames.Count),
                    LocalizationManager.Get("MsgImportSuccessTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Information);

                SetContent(new TemplateView(), ActiveView.Templates);
            }
            catch (Exception ex)
            {
                AppData.AddAction($"Ошибка импорта: {ex.Message}");
                LoadDashboard();
                MessageBox.Show(
                    string.Format(LocalizationManager.Get("MsgImportError"), ex.Message),
                    LocalizationManager.Get("MsgImportErrorTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Cursor = Cursors.Arrow;
                IsEnabled = true;
            }
        }
    }
}
