using Microsoft.Win32;
using OiTech.App.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace OiTech.App
{
    public partial class StudentsView : UserControl
    {
        private List<Student> _allStudents = new();

        public StudentsView()
        {
            InitializeComponent();
            _allStudents = AppData.Students.ToList();
            StudentsGrid.ItemsSource = _allStudents;
            UpdateCount();
        }

        private void UpdateCount()
        {
            var visible = (StudentsGrid.ItemsSource as System.Collections.IList)?.Count ?? 0;
            CountLabel.Text = visible == _allStudents.Count
                ? $"{LocalizationManager.Get("StudentsTotal")}: {_allStudents.Count}"
                : $"{LocalizationManager.Get("StudentsTotal")}: {visible} / {_allStudents.Count}";
        }

        private void ClearSearch(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            SearchBox.Text = "";
        }

        private void Search_Changed(object sender, TextChangedEventArgs e)
        {
            var q = SearchBox.Text;
            PlaceholderText.Visibility = string.IsNullOrEmpty(q)
                ? Visibility.Visible : Visibility.Collapsed;
            ClearBtn.Visibility = string.IsNullOrEmpty(q)
                ? Visibility.Collapsed : Visibility.Visible;

            if (string.IsNullOrWhiteSpace(q))
            {
                StudentsGrid.ItemsSource = _allStudents;
            }
            else
            {
                var ql = q.ToLower();
                StudentsGrid.ItemsSource = _allStudents
                    .Where(s =>
                        s.FIO.ToLower().Contains(ql) ||
                        (s.Group?.ToLower().Contains(ql) ?? false) ||
                        (s.Speciality?.ToLower().Contains(ql) ?? false))
                    .ToList();
            }
            UpdateCount();
        }

        private void ImportWord_Click(object sender, RoutedEventArgs e)
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

                RefreshGrid();

                MessageBox.Show(
                    $"Импортировано: {students.Count} студентов.",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка импорта Word:\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void ImportExcel_Click(object sender, RoutedEventArgs e)
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

                RefreshGrid();

                MessageBox.Show(
                    string.Format(LocalizationManager.Get("MsgImportSuccess"),
                        students.Count, AppData.SubjectNames.Count),
                    LocalizationManager.Get("MsgImportSuccessTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
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

        private void RefreshGrid()
        {
            _allStudents = AppData.Students.ToList();
            SearchBox.Text = "";
            PlaceholderText.Visibility = Visibility.Visible;
            StudentsGrid.ItemsSource = _allStudents;
            UpdateCount();
        }
    }
}
