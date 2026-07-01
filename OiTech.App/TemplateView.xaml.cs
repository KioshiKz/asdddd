using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Printing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace OiTech.App
{
    public partial class TemplateView : UserControl
    {
        private sealed class FieldItem
        {
            public Border Host { get; init; } = null!;
            public TextBlock Text { get; init; } = null!;
            public string Key { get; init; } = "";
        }

        private sealed class FieldOverride
        {
            public string Mode { get; set; } = "";
            public string Key { get; set; } = "";
            public string Text { get; set; } = "";
            public double FontSize { get; set; } = DefaultFontSize;
            public bool HasText { get; set; }
            public bool HasFontSize { get; set; }
        }

        private sealed class CustomTextField
        {
            public string Mode { get; set; } = "";
            public string Key { get; set; } = "";
            public string Text { get; set; } = "Новый текст";
            public double X { get; set; }
            public double Y { get; set; }
            public double Width { get; set; } = 220;
            public double Height { get; set; } = 22;
            public double FontSize { get; set; } = DefaultFontSize;
            public string CanvasName { get; set; } = "vedomost_1";
        }

        private sealed class TableLayout
        {
            public double LeftNoX { get; init; }
            public double LeftSubjectX { get; init; }
            public double LeftHoursX { get; init; }
            public double LeftCreditsX { get; init; }
            public double LeftPercentX { get; init; }
            public double LeftLetterX { get; init; }
            public double LeftGpaX { get; init; }
            public double LeftTraditionalX { get; init; }
            public double LeftBaseY { get; init; }
            public double LeftBottomY { get; init; }
            public double LeftSubjectWidth { get; init; }

            public double RightNoX { get; init; }
            public double RightSubjectX { get; init; }
            public double RightHoursX { get; init; }
            public double RightCreditsX { get; init; }
            public double RightPercentX { get; init; }
            public double RightLetterX { get; init; }
            public double RightGpaX { get; init; }
            public double RightTraditionalX { get; init; }
            public double RightBaseY { get; init; }
            public double RightBottomY { get; init; }
            public double RightSubjectWidth { get; init; }

            public double SecondNoX { get; init; }
            public double SecondSubjectX { get; init; }
            public double SecondHoursX { get; init; }
            public double SecondCreditsX { get; init; }
            public double SecondPercentX { get; init; }
            public double SecondLetterX { get; init; }
            public double SecondGpaX { get; init; }
            public double SecondTraditionalX { get; init; }
            public double SecondBaseY { get; init; }
            public double SecondBottomY { get; init; }
            public double SecondSubjectWidth { get; init; }
            public double SecondRowGap { get; init; }

            // Правая часть 2-й страницы
            public double SecondRightNoX { get; init; }
            public double SecondRightSubjectX { get; init; }
            public double SecondRightHoursX { get; init; }
            public double SecondRightCreditsX { get; init; }
            public double SecondRightPercentX { get; init; }
            public double SecondRightLetterX { get; init; }
            public double SecondRightGpaX { get; init; }
            public double SecondRightTraditionalX { get; init; }
            public double SecondRightBaseY { get; init; }
            public double SecondRightBottomY { get; init; }
            public double SecondRightSubjectWidth { get; init; }

            public int ForcedRightStartIndex { get; init; }
        }

        private readonly List<FieldItem> _fields = new();
        private readonly Dictionary<string, (double X, double Y)> _savedCoords = new();
        private readonly Dictionary<string, CustomTextField> _customTextFields = new();
        private readonly Dictionary<string, FieldOverride> _fieldOverrides = new();
        private readonly HashSet<string> _deletedFieldKeys = new();

        private FieldItem? _selectedField;

        private bool _isReady;
        private bool _updatingCoordBoxes;
        private bool _isDragging;
        private bool _manualBackgroundLoaded;

        private Point _dragStartMouse;
        private double _dragStartX;
        private double _dragStartY;

        private Point _lastRightClickPoint = new Point(120, 120);
        private string _lastRightClickCanvasName = "vedomost_1";

        private string _lastBuildKey = "";
        private int _currentStudentIndex = 0;

        private static readonly double PreviewBgOpacity = 0.85;
        private static readonly double DefaultFontSize = 7.8;
        private static readonly string FixedFontFamily = "Times New Roman";

        private static readonly double MinRowHeight = 18;
        private static readonly double SubjectLineHeight = 9;

        private static readonly double A4LandscapeWidth = 1122.52;
        private static readonly double A4LandscapeHeight = 793.70;

        private static readonly double A5LandscapeWidth = 793.70;
        private static readonly double A5LandscapeHeight = 559.37;

        // РУЧНОЙ СДВИГ ПЕЧАТИ ДИПЛОМА.
        // 0 = без сдвига.
        // 1 = поднять вверх на 1 см.
        // -1 = опустить вниз на 1 см.
        // 0.5 = поднять вверх на 0.5 см.
        // -0.5 = опустить вниз на 0.5 см.
        private static readonly double DiplomaPrintShiftCm = 0.0;


        private static readonly double CmToWpf = 37.8;

        private static readonly string TemplateFolder =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");

        private static readonly string CoordFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OiTech");

        private static readonly string CoordFile =
            Path.Combine(CoordFolder, "template-coordinates.txt");

        private static readonly string CustomFieldsFile =
            Path.Combine(CoordFolder, "custom-text-fields.txt");

        private static readonly string DeletedFieldsFile =
            Path.Combine(CoordFolder, "deleted-fields.txt");

        private static readonly string FieldOverridesFile =
            Path.Combine(CoordFolder, "field-overrides.txt");

        private static string GetTemplatePath(string fileNameWithoutExt)
        {
            string[] extensions = { ".png", ".jpg", ".jpeg", ".bmp" };

            foreach (string ext in extensions)
            {
                string path = Path.Combine(TemplateFolder, fileNameWithoutExt + ext);

                if (File.Exists(path))
                    return path;
            }

            return Path.Combine(TemplateFolder, fileNameWithoutExt + ".png");
        }

        private static readonly TableLayout KazLayout = new()
        {
            LeftNoX = 27,
            LeftSubjectX = 65,
            LeftHoursX = 241,
            LeftCreditsX = 287,

            LeftPercentX = 373,
            LeftLetterX = 337,
            LeftGpaX = 418,
            LeftTraditionalX = 475,

            LeftBaseY = 405,
            LeftBottomY = 745,
            LeftSubjectWidth = 174,

            RightNoX = 578,
            RightSubjectX = 613,
            RightHoursX = 797,
            RightCreditsX = 838,

            RightPercentX = 925,
            RightLetterX = 886,
            RightGpaX = 972,
            RightTraditionalX = 1030,

            RightBaseY = 27,
            RightBottomY = 745,
            RightSubjectWidth = 179,

            SecondNoX = 34,
            SecondSubjectX = 70,
            SecondHoursX = 255,
            SecondCreditsX = 297,

            SecondPercentX = 379,
            SecondLetterX = 341,
            SecondGpaX = 430,
            SecondTraditionalX = 485,

            SecondBaseY = 27,
            SecondBottomY = 750,
            SecondSubjectWidth = 176,
            SecondRowGap = 0,

            SecondRightNoX = 593,
            SecondRightSubjectX = 629,
            SecondRightHoursX = 808,
            SecondRightCreditsX = 849,

            SecondRightPercentX = 936,
            SecondRightLetterX = 908,
            SecondRightGpaX = 991,
            SecondRightTraditionalX = 1043,

            SecondRightBaseY = 31,
            SecondRightBottomY = 700,
            SecondRightSubjectWidth = 175,

            ForcedRightStartIndex = 19
        };

        private static readonly TableLayout RusLayout = new()
        {
            LeftNoX = 27,
            LeftSubjectX = 65,
            LeftHoursX = 241,
            LeftCreditsX = 287,

            LeftPercentX = 373,
            LeftLetterX = 337,
            LeftGpaX = 418,
            LeftTraditionalX = 475,

            LeftBaseY = 405,
            LeftBottomY = 745,
            LeftSubjectWidth = 174,

            RightNoX = 578,
            RightSubjectX = 613,
            RightHoursX = 797,
            RightCreditsX = 838,

            RightPercentX = 925,
            RightLetterX = 886,
            RightGpaX = 972,
            RightTraditionalX = 1030,

            RightBaseY = 27,
            RightBottomY = 745,
            RightSubjectWidth = 179,

            SecondNoX = 34,
            SecondSubjectX = 70,
            SecondHoursX = 255,
            SecondCreditsX = 297,

            SecondPercentX = 379,
            SecondLetterX = 341,
            SecondGpaX = 430,
            SecondTraditionalX = 485,

            SecondBaseY = 27,
            SecondBottomY = 750,
            SecondSubjectWidth = 176,
            SecondRowGap = 0,

            SecondRightNoX = 593,
            SecondRightSubjectX = 629,
            SecondRightHoursX = 808,
            SecondRightCreditsX = 849,

            SecondRightPercentX = 936,
            SecondRightLetterX = 908,
            SecondRightGpaX = 991,
            SecondRightTraditionalX = 1043,

            SecondRightBaseY = 31,
            SecondRightBottomY = 700,
            SecondRightSubjectWidth = 175,

            ForcedRightStartIndex = 19
        };

        private TableLayout CurrentLayout => IsKaz ? KazLayout : RusLayout;

        public TemplateView()
        {
            InitializeComponent();
            Loaded += TemplateView_Loaded;
        }

        private void TemplateView_Loaded(object sender, RoutedEventArgs e)
        {
            if (_isReady)
                return;

            _isReady = true;

            LoadSavedCoordinates();
            LoadCustomTextFields();
            LoadDeletedFields();
            LoadFieldOverrides();

            AttachCanvasContextMenus();

            ApplyMode();
            LoadAutoBackgrounds();
            ScheduleBuild();

            Focusable = true;
            Focus();

            PreviewKeyDown -= TemplateView_PreviewKeyDown;
            PreviewKeyDown += TemplateView_PreviewKeyDown;
        }

        private void TemplateView_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control &&
                e.Key == Key.T)
            {
                AddNewCustomTextField();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Delete)
            {
                DeleteSelectedCustomTextField();
                e.Handled = true;
            }
        }

        private void ScheduleBuild()
        {
            Dispatcher.BeginInvoke(new Action(BuildDocument), DispatcherPriority.Background);
        }

        private string SelectedDocType
        {
            get
            {
                if (DocTypeCombo.SelectedItem is ComboBoxItem item && item.Content is string s)
                    return s.Trim();

                return "Ведомость";
            }
        }

        private string SelectedLang
        {
            get
            {
                if (LangCombo.SelectedItem is ComboBoxItem item && item.Content is string s)
                    return s.Trim();

                return "ҚАЗ";
            }
        }

        private bool IsKaz => SelectedLang == "ҚАЗ";

        private string CurrentModeKey
        {
            get
            {
                if (SelectedDocType == "Диплом")
                    return "Диплом|ALL";

                return $"{SelectedDocType}|{SelectedLang}";
            }
        }

        private Student? CurrentStudent
        {
            get
            {
                if (AppData.Students == null || AppData.Students.Count == 0)
                    return null;

                if (_currentStudentIndex < 0)
                    _currentStudentIndex = 0;

                if (_currentStudentIndex >= AppData.Students.Count)
                    _currentStudentIndex = AppData.Students.Count - 1;

                return AppData.Students[_currentStudentIndex];
            }
        }

        private void DocType_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!_isReady)
                return;

            _lastBuildKey = "";
            _manualBackgroundLoaded = false;

            ApplyMode();
            LoadAutoBackgrounds();
            ScheduleBuild();
        }

        private void Lang_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!_isReady)
                return;

            _lastBuildKey = "";
            _manualBackgroundLoaded = false;

            ApplyMode();
            LoadAutoBackgrounds();
            ScheduleBuild();
        }

        private void ApplyMode()
        {
            bool isDiploma = SelectedDocType == "Диплом";

            DiplomaPage.Visibility = isDiploma ? Visibility.Visible : Visibility.Collapsed;
            VedomostPage.Visibility = isDiploma ? Visibility.Collapsed : Visibility.Visible;
            VedomostPage2.Visibility = isDiploma ? Visibility.Collapsed : Visibility.Visible;

            if (PagePadding != null)
                PagePadding.Visibility = Visibility.Collapsed;

            if (VedomostBg != null)
                VedomostBg.Opacity = isDiploma ? 0 : PreviewBgOpacity;

            if (VedomostBg2 != null)
                VedomostBg2.Opacity = isDiploma ? 0 : PreviewBgOpacity;

            if (DiplomaBgLeft != null)
                DiplomaBgLeft.Opacity = isDiploma ? PreviewBgOpacity : 0;

            if (DiplomaBgRight != null)
                DiplomaBgRight.Opacity = isDiploma ? PreviewBgOpacity : 0;
        }

        private void BuildDocument()
        {
            string studentKey = CurrentStudent?.FIO ?? "no-student";
            string wordKey = CurrentStudent?.WordFIO ?? "";
            string customKey = string.Join(",", _customTextFields.Values.Select(x => $"{x.Key}:{x.Text}:{x.X}:{x.Y}:{x.FontSize}"));
            string deletedKey = string.Join(",", _deletedFieldKeys);
            string overrideKey = string.Join(",", _fieldOverrides.Values.Select(x => $"{x.Mode}:{x.Key}:{x.Text}:{x.FontSize}:{x.HasText}:{x.HasFontSize}"));

            string buildKey =
                $"{SelectedDocType}|{SelectedLang}|{CurrentModeKey}|{_currentStudentIndex}|{studentKey}|{wordKey}|{customKey}|{deletedKey}|{overrideKey}";

            if (_lastBuildKey == buildKey)
                return;

            _lastBuildKey = buildKey;

            ClearAllDynamicFields();

            if (SelectedDocType == "Диплом")
                BuildDiplomaBothSides();
            else if (IsKaz)
                BuildAppendixKaz();
            else
                BuildAppendixRus();

            RenderCustomTextFields();

            UpdateStudentLabel();
        }

        private void UpdateStudentLabel()
        {
            int count = AppData.Students?.Count ?? 0;
            StudentIndexLabel.Text = count == 0 ? "0/0" : $"{_currentStudentIndex + 1}/{count}";
        }

        private void ClearAllDynamicFields()
        {
            VedomostCanvas.Children.Clear();
            VedomostCanvas2.Children.Clear();
            DiplomaLeftCanvas.Children.Clear();
            DiplomaRightCanvas.Children.Clear();

            _fields.Clear();
            _selectedField = null;

            SetCoordBoxes(0, 0);
        }

        private void BuildDiplomaBothSides()
        {
            var s = CurrentStudent;

            string fio = GetBestFio(s);
            string diplomaNo = GetDiplomaNumber(s);
            string speciality = BuildSpecialityText(s);
            string qualification = BuildQualificationText(s);
            string institution = GetInstitutionForPrint();

            AddEditableText(DiplomaLeftCanvas, "dip_kaz_tkb_number", 165, 242, 160, 16, diplomaNo, true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_fio", 95, 308, 395, 16, fio, true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_year_enter", 122, 361, 90, 16, "2021", true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_institution_1", 90, 395, 395, 16, institution, true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_year_finish", 122, 446, 90, 16, "2025", true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_institution_2", 90, 480, 395, 16, institution, true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_speciality", 90, 544, 395, 16, speciality, true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_qualification", 90, 605, 395, 16, qualification, true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_decision_year", 180, 662, 100, 16, "2025", true);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_city", 115, 725, 220, 16, "Өскемен қаласы", false);
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_date", 170, 750, 220, 16, "2025", false);

            // Регистрационный номер внизу убран, чтобы не печаталась лишняя цифра.
            AddEditableText(DiplomaLeftCanvas, "dip_kaz_reg", 185, 775, 180, 16, "", false);

            AddEditableText(DiplomaRightCanvas, "dip_rus_tkb_number", 165, 242, 160, 16, diplomaNo, true);
            AddEditableText(DiplomaRightCanvas, "dip_rus_fio", 95, 308, 395, 16, fio, true);
            AddEditableText(DiplomaRightCanvas, "dip_rus_year_enter", 170, 361, 90, 16, "2021", true);
            AddEditableText(DiplomaRightCanvas, "dip_rus_institution_1", 90, 405, 395, 16, institution, true);
            AddEditableText(DiplomaRightCanvas, "dip_rus_year_finish", 125, 464, 90, 16, "2025", true);
            AddEditableText(DiplomaRightCanvas, "dip_rus_institution_2", 90, 499, 395, 16, institution, true);
            AddEditableText(DiplomaRightCanvas, "dip_rus_speciality", 90, 554, 395, 16, speciality, true);
            AddEditableText(DiplomaRightCanvas, "dip_rus_qualification", 90, 649, 395, 16, qualification, true);
            AddEditableText(DiplomaRightCanvas, "dip_rus_decision_date", 230, 690, 200, 16, "2025", false);
            AddEditableText(DiplomaRightCanvas, "dip_rus_city", 120, 745, 220, 16, "Усть-Каменогорск", false);
            AddEditableText(DiplomaRightCanvas, "dip_rus_date", 170, 770, 220, 16, "2025", false);

            // Регистрационный номер внизу убран, чтобы не печаталась лишняя цифра.
            AddEditableText(DiplomaRightCanvas, "dip_rus_reg", 210, 795, 180, 16, "", false);
        }

        private void BuildAppendixKaz()
        {
            var s = CurrentStudent;

            string fio = GetBestFio(s);
            string diplomaNo = GetDiplomaNumber(s);
            string speciality = BuildSpecialityText(s);
            string qualification = BuildQualificationText(s);
            string institution = GetInstitutionForPrint();

            AddEditableText(VedomostCanvas, "kaz_tkb_number", 237, 61, 170, 16, diplomaNo, false);
            AddEditableText(VedomostCanvas, "kaz_fio", 59, 89, 410, 16, fio, true);
            AddEditableText(VedomostCanvas, "kaz_year_from", 84, 117, 75, 16, "2021", true);
            AddEditableText(VedomostCanvas, "kaz_year_to", 358, 116, 75, 16, "2025", true);
            AddEditableText(VedomostCanvas, "kaz_institution", 46, 142, 470, 16, institution, true);
            AddEditableText(VedomostCanvas, "kaz_speciality", 41, 192, 480, 16, speciality, true);
            AddEditableText(VedomostCanvas, "kaz_qualification", 41, 243, 480, 16, qualification, true);

            FillAppendixGradesSmart();
            AddSecondPageSignatures();
        }

        private void BuildAppendixRus()
        {
            var s = CurrentStudent;

            string fio = GetBestFio(s);
            string diplomaNo = GetDiplomaNumber(s);
            string speciality = BuildSpecialityText(s);
            string qualification = BuildQualificationText(s);
            string institution = GetInstitutionForPrint();

            AddEditableText(VedomostCanvas, "rus_tkb_number", 294, 64, 170, 16, diplomaNo, false);
            AddEditableText(VedomostCanvas, "rus_fio", 94, 98, 410, 16, fio, true);
            AddEditableText(VedomostCanvas, "rus_year_from", 206, 127, 75, 16, "2021", true);
            AddEditableText(VedomostCanvas, "rus_year_to", 422, 125, 75, 16, "2025", true);
            AddEditableText(VedomostCanvas, "rus_institution", 47, 150, 470, 16, institution, true);
            AddEditableText(VedomostCanvas, "rus_speciality", 43, 175, 480, 16, speciality, true);
            AddEditableText(VedomostCanvas, "rus_qualification", 41, 224, 480, 16, qualification, true);

            FillAppendixGradesSmart();
            AddSecondPageSignatures();
        }

        private void FillAppendixGradesSmart()
        {
            var student = CurrentStudent;

            if (student == null || student.Grades == null || student.Grades.Count == 0)
                return;

            TableLayout layout = CurrentLayout;

            double leftY = layout.LeftBaseY;
            double rightY = layout.RightBaseY;
            double secondLeftY = layout.SecondBaseY;
            double secondRightY = layout.SecondRightBaseY;

            int area = 0;

            int maxGradesToPrint = Math.Min(student.Grades.Count, 87);

            for (int i = 0; i < maxGradesToPrint; i++)
            {
                GradeEntry g = student.Grades[i];

                if (area == 0 && i >= layout.ForcedRightStartIndex)
                    area = 1;

                string subjectPrintText = GetSubjectTextForPrint(g);

                // Вложенные строки (7.1, 7.2 ...) печатаются с отступом слева (см.
                // AddGradeRow), поэтому реальная ширина под текст у них меньше —
                // это нужно учитывать при расчёте высоты строки, иначе перенос
                // текста посчитается неверно и строки в таблице разъедутся.
                double indentForRow = IsGradePart(g) ? 14 : 0;

                double subjectWidth = GetSubjectWidth(layout, area);
                double rowHeight = CalculateRealRowHeight(subjectPrintText, Math.Max(40, subjectWidth - indentForRow));

                if (area == 0 && leftY + rowHeight > layout.LeftBottomY)
                    area = 1;

                if (area == 1)
                {
                    subjectWidth = layout.RightSubjectWidth;
                    rowHeight = CalculateRealRowHeight(subjectPrintText, Math.Max(40, subjectWidth - indentForRow));

                    if (rightY + rowHeight > layout.RightBottomY)
                        area = 2;
                }

                if (area == 2)
                {
                    subjectWidth = layout.SecondSubjectWidth;
                    rowHeight = CalculateRealRowHeight(subjectPrintText, Math.Max(40, subjectWidth - indentForRow));

                    if (secondLeftY + rowHeight > layout.SecondBottomY)
                        area = 3;
                }

                if (area == 3)
                {
                    subjectWidth = layout.SecondRightSubjectWidth;
                    rowHeight = CalculateRealRowHeight(subjectPrintText, Math.Max(40, subjectWidth - indentForRow));

                    if (secondRightY + rowHeight > layout.SecondRightBottomY)
                        break;
                }

                if (area == 0)
                {
                    AddGradeRow(
                        VedomostCanvas,
                        i,
                        g,
                        layout.LeftNoX,
                        layout.LeftSubjectX,
                        layout.LeftHoursX,
                        layout.LeftCreditsX,
                        layout.LeftPercentX,
                        layout.LeftLetterX,
                        layout.LeftGpaX,
                        layout.LeftTraditionalX,
                        leftY,
                        rowHeight,
                        layout.LeftSubjectWidth);

                    leftY += rowHeight;
                }
                else if (area == 1)
                {
                    AddGradeRow(
                        VedomostCanvas,
                        i,
                        g,
                        layout.RightNoX,
                        layout.RightSubjectX,
                        layout.RightHoursX,
                        layout.RightCreditsX,
                        layout.RightPercentX,
                        layout.RightLetterX,
                        layout.RightGpaX,
                        layout.RightTraditionalX,
                        rightY,
                        rowHeight,
                        layout.RightSubjectWidth);

                    rightY += rowHeight;
                }
                else if (area == 2)
                {
                    AddGradeRow(
                        VedomostCanvas2,
                        i,
                        g,
                        layout.SecondNoX,
                        layout.SecondSubjectX,
                        layout.SecondHoursX,
                        layout.SecondCreditsX,
                        layout.SecondPercentX,
                        layout.SecondLetterX,
                        layout.SecondGpaX,
                        layout.SecondTraditionalX,
                        secondLeftY,
                        rowHeight,
                        layout.SecondSubjectWidth);

                    secondLeftY += rowHeight + layout.SecondRowGap;
                }
                else
                {
                    AddGradeRow(
                        VedomostCanvas2,
                        i,
                        g,
                        layout.SecondRightNoX,
                        layout.SecondRightSubjectX,
                        layout.SecondRightHoursX,
                        layout.SecondRightCreditsX,
                        layout.SecondRightPercentX,
                        layout.SecondRightLetterX,
                        layout.SecondRightGpaX,
                        layout.SecondRightTraditionalX,
                        secondRightY,
                        rowHeight,
                        layout.SecondRightSubjectWidth);

                    secondRightY += rowHeight + layout.SecondRowGap;
                }
            }
        }

        private double GetSubjectWidth(TableLayout layout, int area)
        {
            if (area == 0)
                return layout.LeftSubjectWidth;

            if (area == 1)
                return layout.RightSubjectWidth;

            if (area == 2)
                return layout.SecondSubjectWidth;

            return layout.SecondRightSubjectWidth;
        }

        private void AddGradeRow(
            Canvas canvas,
            int index,
            GradeEntry g,
            double xNo,
            double xSubject,
            double xHours,
            double xCredits,
            double xPercent,
            double xLetter,
            double xGpa,
            double xTraditional,
            double y,
            double rowHeight,
            double subjectWidth)
        {
            bool isPart = IsGradePart(g);

            string subjectText = GetSubjectTextForPrint(g);

            double indent = isPart ? 14 : 0;
            double subjectX = xSubject + indent;
            double subjectW = Math.Max(40, subjectWidth - indent);

            // Для вложенных строк номер слева не ставим.
            // Пример:
            // 7 Информатика
            //     7.1 Физика
            //     7.2 Химия
            string noText = isPart ? "" : (index + 1).ToString();

            AddEditableText(canvas, $"no_{index}", xNo, y, 28, rowHeight, noText, true);
            AddEditableText(canvas, $"subject_{index}", subjectX, y - 1, subjectW, rowHeight, subjectText, false, true);
            AddEditableText(canvas, $"hours_{index}", xHours, y, 34, rowHeight, g.Hours, true);
            AddEditableText(canvas, $"credits_{index}", xCredits, y, 48, rowHeight, g.Credits, true);

            AddEditableText(canvas, $"percent_{index}", xPercent, y, 32, rowHeight, g.Percentage, true);
            AddEditableText(canvas, $"letter_{index}", xLetter, y, 38, rowHeight, g.LetterGrade, true);
            AddEditableText(canvas, $"gpa_{index}", xGpa, y, 42, rowHeight, g.GPA, true);
            AddEditableText(canvas, $"traditional_{index}", xTraditional, y, 52, rowHeight, g.TraditionalGrade, true);
        }

        private bool IsGradePart(GradeEntry g)
        {
            if (g == null)
                return false;

            if (g.Level > 0)
                return true;

            if (!string.IsNullOrWhiteSpace(g.RowType) &&
                g.RowType.Equals("PART", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(g.SubjectCode) && g.SubjectCode.Contains("."))
                return true;

            return false;
        }

        private string GetSubjectTextForPrint(GradeEntry g)
        {
            if (g == null)
                return "";

            string subject = GetSubjectByLanguage(g);
            string code = g.SubjectCode?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(code))
                return subject;

            // Если название предмета не нашлось при импорте, ExcelService подставляет
            // сам код как название (чтобы не терять строку). Тогда subject == code,
            // и наивная склейка "код + название" печатает код дважды подряд.
            if (string.IsNullOrWhiteSpace(subject) ||
                subject.Trim().StartsWith(code, StringComparison.OrdinalIgnoreCase))
            {
                return code;
            }

            return $"{code} {subject}".Trim();
        }

        private string GetSubjectByLanguage(GradeEntry g)
        {
            if (g == null)
                return "";

            if (IsKaz)
            {
                string kaz = GetPropertyValue(g, "SubjectKaz");

                if (!string.IsNullOrWhiteSpace(kaz))
                    return kaz;

                return g.Subject ?? "";
            }

            string rus = GetPropertyValue(g, "SubjectRus");

            if (!string.IsNullOrWhiteSpace(rus))
                return rus;

            return g.Subject ?? "";
        }

        private static string GetPropertyValue(object obj, string propertyName)
        {
            try
            {
                PropertyInfo? prop = obj.GetType().GetProperty(propertyName);

                if (prop == null)
                    return "";

                object? value = prop.GetValue(obj);

                return value?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }

        private void AddSecondPageSignatures()
        {
            if (SelectedDocType == "Диплом")
                return;

            string curator = CurrentStudent?.Curator ?? "";

            if (IsKaz)
            {
                AddEditableText(VedomostCanvas2, "kaz_second_deputy_name", 970, 505, 230, 18, "", false);
                AddEditableText(VedomostCanvas2, "kaz_second_group_leader_name", 833, 527, 230, 18, curator, false);
            }
            else
            {
                AddEditableText(VedomostCanvas2, "rus_second_deputy_name", 970, 505, 230, 18, "", false);
                AddEditableText(VedomostCanvas2, "rus_second_group_leader_name", 833, 527, 230, 18, curator, false);
            }
        }

        private double CalculateRealRowHeight(string subject, double width)
        {
            double fontSize = DefaultFontSize;

            // Меряем перенос текста по заведомо более узкой ширине, чем реальная
            // колонка. FormattedText (используется здесь) и TextBlock (реально рисует
            // текст в AddEditableText) не всегда переносят длинный текст по строкам
            // одинаково — WPF-квирк, из-за которого точный расчёт может занижать
            // нужную высоту. Меряя по более узкой ширине, мы намеренно считаем строк
            // не меньше, чем реально нужно, — то есть эта оценка гарантированно
            // не меньше настоящей необходимой высоты.
            double safeWidth = width * 0.8;

            double textHeight = MeasureWrappedTextHeight(subject, safeWidth, fontSize);
            double needed = textHeight + 10.0;

            if (needed < MinRowHeight)
                needed = MinRowHeight;

            return needed;
        }

        private double MeasureWrappedTextHeight(string text, double width, double fontSize)
        {
            if (string.IsNullOrWhiteSpace(text))
                return SubjectLineHeight;

            string clean = Regex.Replace(text.Trim(), @"\s+", " ");

            try
            {
                double pixelsPerDip = 1.0;

                try
                {
                    pixelsPerDip = VisualTreeHelper.GetDpi(this).PixelsPerDip;
                }
                catch
                {
                    pixelsPerDip = 1.0;
                }

                var formatted = new FormattedText(
                    clean,
                    CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface(FixedFontFamily),
                    fontSize,
                    Brushes.Black,
                    pixelsPerDip)
                {
                    MaxTextWidth = width,
                    // Должно совпадать с LineHeight реального TextBlock в AddEditableText,
                    // иначе расчёт высоты строки занижается и текст накладывается на соседние строки.
                    LineHeight = Math.Max(SubjectLineHeight + 1.5, fontSize + 1.5)
                };

                return formatted.Height;
            }
            catch
            {
                int charsPerLine = Math.Max(10, (int)(width / 5.6));
                int lines = 1;
                int current = 0;

                foreach (string word in clean.Split(' '))
                {
                    int len = word.Length;

                    if (current == 0)
                    {
                        current = len;
                    }
                    else if (current + 1 + len <= charsPerLine)
                    {
                        current += 1 + len;
                    }
                    else
                    {
                        lines++;
                        current = len;
                    }
                }

                return lines * Math.Max(SubjectLineHeight, fontSize + 0.2);
            }
        }

        private void LoadAutoBackgrounds()
        {
            if (_manualBackgroundLoaded)
                return;

            try
            {
                bool isDiploma = SelectedDocType == "Диплом";

                if (isDiploma)
                {
                    string diplomaPath = GetTemplatePath("diplom_kaz_face");

                    if (!File.Exists(diplomaPath))
                        diplomaPath = GetTemplatePath("diplom_rus_face");

                    if (File.Exists(diplomaPath))
                    {
                        DiplomaBgLeft.Source = LoadBitmapPart(diplomaPath, true);
                        DiplomaBgRight.Source = LoadBitmapPart(diplomaPath, false);
                    }

                    return;
                }

                string facePath = IsKaz
                    ? GetTemplatePath("vedomost_kaz_face")
                    : GetTemplatePath("vedomost_rus_face");

                string backPath = IsKaz
                    ? GetTemplatePath("vedomost_kaz_back")
                    : GetTemplatePath("vedomost_rus_back");

                if (File.Exists(facePath))
                    VedomostBg.Source = LoadBitmap(facePath);

                if (File.Exists(backPath))
                    VedomostBg2.Source = LoadBitmap(backPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка загрузки шаблонов:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private Border AddEditableText(
            Canvas canvas,
            string key,
            double x,
            double y,
            double width,
            double height,
            string text,
            bool centered,
            bool wrap = false)
        {
            if (IsFieldDeleted(key))
                return new Border();

            string finalText = ApplyTextOverride(key, text);
            double fontSize = GetFieldFontSize(key);

            (double X, double Y) saved = GetSavedCoord(key, x, y);

            // Для числовых полей используем центрирование по вертикали
            VerticalAlignment vertAlign = VerticalAlignment.Center;

            // Для центрированных полей добавляем вертикальный отступ
            Thickness padding = new Thickness(1, 1, 1, 1);

            var label = new TextBlock
            {
                Text = finalText,
                FontSize = fontSize,
                FontFamily = new FontFamily(FixedFontFamily),
                Foreground = Brushes.Black,
                TextAlignment = centered ? TextAlignment.Center : TextAlignment.Left,
                VerticalAlignment = vertAlign,
                HorizontalAlignment = centered ? HorizontalAlignment.Center : HorizontalAlignment.Left,
                TextWrapping = wrap ? TextWrapping.Wrap : TextWrapping.NoWrap,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(0),
                Padding = new Thickness(0, 2, 0, 2),
                LineHeight = Math.Max(SubjectLineHeight + 1.5, fontSize + 1.5)
            };

            label.ToolTip = key;

            if (wrap)
            {
                height += 8;
            }

            var border = new Border
            {
                Width = width,
                Height = height,
                Background = Brushes.Transparent,
                Child = label,
                Tag = key,
                Padding = new Thickness(0),
                Margin = new Thickness(0),
                Cursor = Cursors.SizeAll,
                ClipToBounds = true
            };

            Canvas.SetLeft(border, saved.X);
            Canvas.SetTop(border, saved.Y);
            Panel.SetZIndex(border, 999);

            border.MouseLeftButtonDown += Field_MouseLeftButtonDown;
            border.MouseMove += Field_MouseMove;
            border.MouseLeftButtonUp += Field_MouseLeftButtonUp;
            border.ContextMenu = BuildFieldContextMenu(border);

            canvas.Children.Add(border);

            _fields.Add(new FieldItem
            {
                Host = border,
                Text = label,
                Key = key
            });

            return border;
        }

        private string ApplyTextOverride(string key, string defaultText)
        {
            if (key.StartsWith("custom_", StringComparison.OrdinalIgnoreCase))
                return defaultText;

            if (IsDynamicDataField(key))
                return defaultText;

            string overrideKey = MakeFieldOverrideKey(key);

            if (_fieldOverrides.TryGetValue(overrideKey, out FieldOverride? ov) && ov.HasText)
                return ov.Text;

            return defaultText;
        }

        private double GetFieldFontSize(string key)
        {
            if (IsGradeField(key))
                return DefaultFontSize;

            if (key.StartsWith("custom_", StringComparison.OrdinalIgnoreCase))
            {
                if (_customTextFields.TryGetValue(key, out CustomTextField? custom))
                    return custom.FontSize > 0 ? custom.FontSize : DefaultFontSize;

                return DefaultFontSize;
            }

            string overrideKey = MakeFieldOverrideKey(key);

            if (_fieldOverrides.TryGetValue(overrideKey, out FieldOverride? ov) && ov.HasFontSize && ov.FontSize > 0)
                return ov.FontSize;

            return DefaultFontSize;
        }

        private void SaveCurrentFieldOverrideIfNeeded(FieldItem item)
        {
            if (item.Key.StartsWith("custom_", StringComparison.OrdinalIgnoreCase))
            {
                if (_customTextFields.TryGetValue(item.Key, out CustomTextField? custom))
                {
                    custom.Text = item.Text.Text;
                    custom.FontSize = item.Text.FontSize;
                    SaveCustomTextFields();
                }

                return;
            }

            string overrideKey = MakeFieldOverrideKey(item.Key);

            if (!_fieldOverrides.TryGetValue(overrideKey, out FieldOverride? ov))
            {
                ov = new FieldOverride
                {
                    Mode = CurrentModeKey,
                    Key = item.Key
                };

                _fieldOverrides[overrideKey] = ov;
            }

            // Размер шрифта сохраняем всегда.
            // Это важно для ФИО, номера диплома, специальности, квалификации и оценок.
            ov.FontSize = item.Text.FontSize;
            ov.HasFontSize = true;

            // Текст динамических полей не сохраняем,
            // чтобы при переключении студента данные обновлялись из Excel/Word.
            if (IsDynamicDataField(item.Key))
            {
                ov.Text = "";
                ov.HasText = false;
            }
            else
            {
                ov.Text = item.Text.Text;
                ov.HasText = true;
            }

            SaveFieldOverrides();
        }

        private bool IsDynamicDataField(string key)
        {
            if (IsGradeField(key))
                return true;

            string[] prefixes =
            {
                "kaz_fio",
                "rus_fio",
                "dip_kaz_fio",
                "dip_rus_fio",
                "kaz_tkb_number",
                "rus_tkb_number",
                "dip_kaz_tkb_number",
                "dip_rus_tkb_number",
                "kaz_qualification",
                "rus_qualification",
                "dip_kaz_qualification",
                "dip_rus_qualification" 
            };

            return prefixes.Any(p => key.StartsWith(p, StringComparison.OrdinalIgnoreCase));
        }

        private string MakeFieldOverrideKey(string key)
        {
            return $"{CurrentModeKey}|{key}";
        }

        private ContextMenu BuildFieldContextMenu(Border border)
        {
            var menu = new ContextMenu();

            var editItem = new MenuItem
            {
                Header = "Изменить текст"
            };

            editItem.Click += (_, __) =>
            {
                var item = _fields.FirstOrDefault(x => x.Host == border);

                if (item != null)
                {
                    SelectField(item);
                    OpenEditor(item);
                }
            };

            menu.Items.Add(editItem);

            var deleteItem = new MenuItem
            {
                Header = "Удалить"
            };

            deleteItem.Click += (_, __) =>
            {
                var item = _fields.FirstOrDefault(x => x.Host == border);

                if (item == null)
                    return;

                SelectField(item);
                DeleteSelectedFieldSmart();
            };

            menu.Items.Add(deleteItem);

            return menu;
        }

        private void AttachCanvasContextMenus()
        {
            AttachCanvasContextMenu(VedomostCanvas, "vedomost_1");
            AttachCanvasContextMenu(VedomostCanvas2, "vedomost_2");
            AttachCanvasContextMenu(DiplomaLeftCanvas, "diploma_left");
            AttachCanvasContextMenu(DiplomaRightCanvas, "diploma_right");
        }

        private void AttachCanvasContextMenu(Canvas canvas, string canvasName)
        {
            canvas.PreviewMouseRightButtonDown += (_, e) =>
            {
                _lastRightClickPoint = e.GetPosition(canvas);
                _lastRightClickCanvasName = canvasName;
                Focus();
            };

            var menu = new ContextMenu();

            var addItem = new MenuItem
            {
                Header = "Добавить текстовое поле"
            };

            addItem.Click += (_, __) =>
            {
                AddNewCustomTextField(
                    _lastRightClickCanvasName,
                    _lastRightClickPoint.X,
                    _lastRightClickPoint.Y,
                    openEditorImmediately: true);
            };

            menu.Items.Add(addItem);

            canvas.ContextMenu = menu;
        }

        private (double X, double Y) GetSavedCoord(string key, double defaultX, double defaultY)
        {

            if (key.StartsWith("custom_", StringComparison.OrdinalIgnoreCase))
                return (defaultX, defaultY);

            string fullKey = MakeCoordKey(key);

            if (_savedCoords.TryGetValue(fullKey, out var saved))
                return saved;

            return (defaultX, defaultY);
        }

        private bool IsGradeField(string key)
        {
            return Regex.IsMatch(
                key,
                @"^(no|subject|hours|credits|percent|letter|gpa|traditional)_\d+$");
        }

        private string MakeCoordKey(string key)
        {
            return $"{CurrentModeKey}|{key}";
        }

        private string MakeDeletedFieldKey(string key)
        {
            return $"{CurrentModeKey}|{key}";
        }

        private bool IsFieldDeleted(string key)
        {
            if (key.StartsWith("custom_", StringComparison.OrdinalIgnoreCase))
                return false;

            string deletedKey = MakeDeletedFieldKey(key);

            return _deletedFieldKeys.Contains(deletedKey);
        }

        private void Field_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Border border)
                return;

            var item = _fields.FirstOrDefault(x => x.Host == border);

            if (item == null)
                return;

            SelectField(item);

            if (e.ClickCount == 2)
            {
                OpenEditor(item);
                e.Handled = true;
                return;
            }

            _isDragging = true;
            _dragStartMouse = e.GetPosition(border.Parent as UIElement);
            _dragStartX = GetLeft(border);
            _dragStartY = GetTop(border);

            border.CaptureMouse();
            e.Handled = true;
        }

        private void Field_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging || _selectedField == null)
                return;

            if (_selectedField.Host.Parent is not UIElement parent)
                return;

            Point current = e.GetPosition(parent);

            double dx = current.X - _dragStartMouse.X;
            double dy = current.Y - _dragStartMouse.Y;

            double newX = _dragStartX + dx;
            double newY = _dragStartY + dy;

            Canvas.SetLeft(_selectedField.Host, newX);
            Canvas.SetTop(_selectedField.Host, newY);

            SetCoordBoxes(newX, newY);
        }

        private void Field_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDragging)
                return;

            _isDragging = false;

            if (sender is Border border)
                border.ReleaseMouseCapture();

            SaveCoordinatesToMemory();
            SaveCoordinatesToFile();
            SaveCustomPositionIfNeeded();

        }

        private void SelectField(FieldItem item)
        {
            _selectedField = item;

            SetCoordBoxes(GetLeft(item.Host), GetTop(item.Host));

            _updatingCoordBoxes = true;
            FontSizeBox.Text = item.Text.FontSize.ToString("0.##", CultureInfo.InvariantCulture);
            _updatingCoordBoxes = false;

            Focus();
        }

        private void SetCoordBoxes(double x, double y)
        {
            _updatingCoordBoxes = true;

            FieldXBox.Text = x.ToString("0.##", CultureInfo.InvariantCulture);
            FieldYBox.Text = y.ToString("0.##", CultureInfo.InvariantCulture);

            _updatingCoordBoxes = false;
        }

        private void OpenEditor(FieldItem item)
        {
            var win = new Window
            {
                Title = "Редактировать текст",
                Width = 560,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Owner = Window.GetWindow(this),
                Background = Brushes.White
            };

            var panel = new StackPanel { Margin = new Thickness(12) };

            var box = new TextBox
            {
                Text = item.Text.Text,
                FontFamily = new FontFamily(FixedFontFamily),
                FontSize = 16,
                Height = 34,
                VerticalContentAlignment = VerticalAlignment.Center
            };

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };

            var ok = new Button
            {
                Content = "Сохранить",
                Width = 110,
                Height = 30,
                Margin = new Thickness(0, 0, 8, 0)
            };

            var cancel = new Button
            {
                Content = "Отмена",
                Width = 90,
                Height = 30
            };

            ok.Click += (_, __) =>
            {
                item.Text.Text = box.Text;

                SaveCurrentFieldOverrideIfNeeded(item);

                _lastBuildKey = "";

                win.DialogResult = true;
                win.Close();
            };

            cancel.Click += (_, __) =>
            {
                win.DialogResult = false;
                win.Close();
            };

            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            panel.Children.Add(box);
            panel.Children.Add(buttons);

            win.Content = panel;

            win.Loaded += (_, __) =>
            {
                box.Focus();
                box.SelectAll();
            };

            win.ShowDialog();
        }

        private void Canvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is Canvas)
                _selectedField = null;

            Focus();
        }

        private void Canvas_MouseMove(object sender, MouseEventArgs e)
        {
        }

        private void Canvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
        }

        private void FieldCoord_Changed(object sender, TextChangedEventArgs e)
        {
            if (_updatingCoordBoxes || _selectedField == null)
                return;

            if (double.TryParse(FieldXBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double x))
                Canvas.SetLeft(_selectedField.Host, x);

            if (double.TryParse(FieldYBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double y))
                Canvas.SetTop(_selectedField.Host, y);

            SaveCoordinatesToMemory();
            SaveCoordinatesToFile();
            SaveCustomPositionIfNeeded();

        }

        private void Font_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_selectedField == null)
                return;

            _selectedField.Text.FontFamily = new FontFamily(FixedFontFamily);
            SaveCurrentFieldOverrideIfNeeded(_selectedField);
        }

        private void FontSize_Changed(object sender, TextChangedEventArgs e)
        {
            if (_updatingCoordBoxes || _selectedField == null)
                return;

            string raw = FontSizeBox.Text.Trim().Replace(',', '.');

            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double size))
                return;

            if (size < 5)
                size = 5;

            if (size > 40)
                size = 40;

            _selectedField.Text.FontSize = size;
            _selectedField.Text.LineHeight = Math.Max(SubjectLineHeight, size + 0.2);

            SaveCurrentFieldOverrideIfNeeded(_selectedField);

            _lastBuildKey = "";
        }

        private void Margin_Changed(object sender, TextChangedEventArgs e)
        {
            if (!IsLoaded)
                return;

            if (!double.TryParse(MarginTop.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double top))
                top = 5;

            if (!double.TryParse(MarginRight.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double right))
                right = 5;

            if (!double.TryParse(MarginBottom.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double bottom))
                bottom = 5;

            if (!double.TryParse(MarginLeft.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double left))
                left = 5;

            DocumentContent.Margin = new Thickness(left, top, right, bottom);
        }

        private void SaveCoordinatesToMemory()
        {
            foreach (var f in _fields)
            {
                if (f.Key.StartsWith("custom_", StringComparison.OrdinalIgnoreCase))
                    continue;

                string key = MakeCoordKey(f.Key);
                _savedCoords[key] = (GetLeft(f.Host), GetTop(f.Host));
            }
        }

        private void SaveCoordinatesToFile()
        {
            try
            {
                Directory.CreateDirectory(CoordFolder);

                var sb = new StringBuilder();

                foreach (var item in _savedCoords.OrderBy(x => x.Key))
                    sb.AppendLine($"{item.Key}|{item.Value.X.ToString(CultureInfo.InvariantCulture)}|{item.Value.Y.ToString(CultureInfo.InvariantCulture)}");

                File.WriteAllText(CoordFile, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
            }
        }

        private void LoadSavedCoordinates()
        {
            _savedCoords.Clear();

            try
            {
                if (!File.Exists(CoordFile))
                    return;

                foreach (string line in File.ReadAllLines(CoordFile, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] parts = line.Split('|');

                    if (parts.Length < 4)
                        continue;

                    string key;
                    string xText;
                    string yText;

                    if (parts.Length >= 5)
                    {
                        key = $"{parts[0]}|{parts[1]}|{parts[2]}";
                        xText = parts[3];
                        yText = parts[4];
                    }
                    else
                    {
                        key = $"{parts[0]}|{parts[1]}";
                        xText = parts[2];
                        yText = parts[3];
                    }

                    if (!double.TryParse(xText, NumberStyles.Float, CultureInfo.InvariantCulture, out double x))
                        continue;

                    if (!double.TryParse(yText, NumberStyles.Float, CultureInfo.InvariantCulture, out double y))
                        continue;

                    _savedCoords[key] = (x, y);
                }
            }
            catch
            {
            }
        }

        private void SaveDeletedFields()
        {
            try
            {
                Directory.CreateDirectory(CoordFolder);

                File.WriteAllLines(
                    DeletedFieldsFile,
                    _deletedFieldKeys.OrderBy(x => x),
                    Encoding.UTF8);
            }
            catch
            {
            }
        }

        private void LoadDeletedFields()
        {
            _deletedFieldKeys.Clear();

            try
            {
                if (!File.Exists(DeletedFieldsFile))
                    return;

                foreach (string line in File.ReadAllLines(DeletedFieldsFile, Encoding.UTF8))
                {
                    if (!string.IsNullOrWhiteSpace(line))
                        _deletedFieldKeys.Add(line.Trim());
                }
            }
            catch
            {
            }
        }

        private void SaveFieldOverrides()
        {
            try
            {
                Directory.CreateDirectory(CoordFolder);

                var sb = new StringBuilder();

                foreach (var item in _fieldOverrides.Values.OrderBy(x => x.Mode).ThenBy(x => x.Key))
                {
                    string line =
                        $"{Escape(item.Mode)}|" +
                        $"{Escape(item.Key)}|" +
                        $"{Escape(item.Text)}|" +
                        $"{item.FontSize.ToString(CultureInfo.InvariantCulture)}|" +
                        $"{item.HasText}|" +
                        $"{item.HasFontSize}";

                    sb.AppendLine(line);
                }

                File.WriteAllText(FieldOverridesFile, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
            }
        }

        private void LoadFieldOverrides()
        {
            _fieldOverrides.Clear();

            try
            {
                if (!File.Exists(FieldOverridesFile))
                    return;

                foreach (string line in File.ReadAllLines(FieldOverridesFile, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] p = line.Split('|');

                    if (p.Length < 6)
                        continue;

                    var ov = new FieldOverride
                    {
                        Mode = Unescape(p[0]),
                        Key = Unescape(p[1]),
                        Text = Unescape(p[2]),
                        FontSize = ParseDoubleSafe(p[3], DefaultFontSize),
                        HasText = bool.TryParse(p[4], out bool ht) && ht,
                        HasFontSize = bool.TryParse(p[5], out bool hf) && hf
                    };

                    if (!string.IsNullOrWhiteSpace(ov.Mode) && !string.IsNullOrWhiteSpace(ov.Key))
                    {
                        // Для динамических полей текст из файла не берём,
                        // но размер шрифта оставляем.
                        if (IsDynamicDataField(ov.Key))
                        {
                            ov.Text = "";
                            ov.HasText = false;
                        }

                        _fieldOverrides[$"{ov.Mode}|{ov.Key}"] = ov;
                    }
                }
            }
            catch
            {
            }
        }

        private void SaveCustomTextFields()
        {
            try
            {
                Directory.CreateDirectory(CoordFolder);

                var sb = new StringBuilder();

                foreach (var f in _customTextFields.Values.OrderBy(x => x.Mode).ThenBy(x => x.Key))
                {
                    string line =
                        $"{Escape(f.Mode)}|" +
                        $"{Escape(f.Key)}|" +
                        $"{Escape(f.CanvasName)}|" +
                        $"{f.X.ToString(CultureInfo.InvariantCulture)}|" +
                        $"{f.Y.ToString(CultureInfo.InvariantCulture)}|" +
                        $"{f.Width.ToString(CultureInfo.InvariantCulture)}|" +
                        $"{f.Height.ToString(CultureInfo.InvariantCulture)}|" +
                        $"{f.FontSize.ToString(CultureInfo.InvariantCulture)}|" +
                        $"{Escape(f.Text)}";

                    sb.AppendLine(line);
                }

                File.WriteAllText(CustomFieldsFile, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
            }
        }

        private void LoadCustomTextFields()
        {
            _customTextFields.Clear();

            try
            {
                if (!File.Exists(CustomFieldsFile))
                    return;

                foreach (string line in File.ReadAllLines(CustomFieldsFile, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] p = line.Split('|');

                    if (p.Length < 8)
                        continue;

                    var field = new CustomTextField
                    {
                        Mode = Unescape(p[0]),
                        Key = Unescape(p[1]),
                        CanvasName = Unescape(p[2]),
                        X = ParseDoubleSafe(p[3]),
                        Y = ParseDoubleSafe(p[4]),
                        Width = ParseDoubleSafe(p[5], 220),
                        Height = ParseDoubleSafe(p[6], 22),
                        FontSize = p.Length >= 9 ? ParseDoubleSafe(p[7], DefaultFontSize) : DefaultFontSize,
                        Text = p.Length >= 9 ? Unescape(p[8]) : Unescape(p[7])
                    };

                    if (!string.IsNullOrWhiteSpace(field.Key))
                        _customTextFields[field.Key] = field;
                }
            }
            catch
            {
            }
        }

        private void ExportExcel(object sender, RoutedEventArgs e)
        {
            SaveCoordinatesToMemory();
            SaveCoordinatesToFile();
            SaveCustomTextFields();
            SaveDeletedFields();
            SaveFieldOverrides();

            var sb = new StringBuilder();

            sb.AppendLine($"Файл координат: {CoordFile}");
            sb.AppendLine($"Файл доп. текстов: {CustomFieldsFile}");
            sb.AppendLine($"Файл удалённых полей: {DeletedFieldsFile}");
            sb.AppendLine($"Файл изменений текста/шрифта: {FieldOverridesFile}");
            sb.AppendLine($"Режим: {CurrentModeKey}");
            sb.AppendLine($"Студент: {CurrentStudent?.FIO}");
            sb.AppendLine($"Сдвиг диплома при печати, см: {DiplomaPrintShiftCm}");
            sb.AppendLine();

            foreach (var f in _fields)
            {
                double x = GetLeft(f.Host);
                double y = GetTop(f.Host);

                sb.AppendLine($"{MakeCoordKey(f.Key)} | X={x:0.##} | Y={y:0.##} | W={f.Host.Width:0.##} | H={f.Host.Height:0.##} | FS={f.Text.FontSize:0.##} | TEXT={f.Text.Text}");
            }

            Clipboard.SetText(sb.ToString());

            MessageBox.Show(
                "Координаты, текст, шрифты и удалённые поля сохранены.",
                "Координаты",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void PrintSheet(object sender, RoutedEventArgs e)
        {
            BuildDocument();
            UpdateLayout();

            var dialog = new PrintDialog();

            ApplyPrintTicket(dialog);

            if (dialog.ShowDialog() != true)
                return;

            ApplyPrintTicket(dialog);

            try
            {
                FixedDocument document = SelectedDocType == "Диплом"
                    ? BuildDiplomaPrintDocument()
                    : BuildVedomostPrintDocument();

                dialog.PrintDocument(
                    document.DocumentPaginator,
                    SelectedDocType == "Диплом" ? "Диплом A5 Landscape" : "Ведомость A4 Landscape");

                AppData.PrintedCount++;
                AppData.AddAction($"Печать: {SelectedDocType} {SelectedLang}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка печати:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void ApplyPrintTicket(PrintDialog dialog)
        {
            try
            {
                if (SelectedDocType == "Диплом")
                {
                    dialog.PrintTicket.PageMediaSize = new PageMediaSize(A5LandscapeWidth, A5LandscapeHeight);
                    dialog.PrintTicket.PageOrientation = PageOrientation.Landscape;
                }
                else
                {
                    dialog.PrintTicket.PageMediaSize = new PageMediaSize(A4LandscapeWidth, A4LandscapeHeight);
                    dialog.PrintTicket.PageOrientation = PageOrientation.Landscape;
                }

                if (dialog.PrintQueue != null)
                    dialog.PrintQueue.UserPrintTicket = dialog.PrintTicket;
            }
            catch (Exception ex)
            {
                // Принтер может не поддерживать нужный размер/ориентацию страницы —
                // печать всё равно продолжится с настройками по умолчанию, но пользователь должен это видеть.
                AppData.AddAction($"Не удалось применить настройки печати (размер/ориентация страницы): {ex.Message}");
            }
        }

        private FixedDocument BuildDiplomaPrintDocument()
        {
            var document = new FixedDocument();
            document.DocumentPaginator.PageSize = new Size(A5LandscapeWidth, A5LandscapeHeight);

            double printShiftY = -DiplomaPrintShiftCm * CmToWpf;

            AddFrameworkElementAsTextOnlyPrintPage(
                document,
                DiplomaPage,
                A5LandscapeWidth,
                A5LandscapeHeight,
                printShiftY);

            return document;
        }

        private FixedDocument BuildVedomostPrintDocument()
        {
            var document = new FixedDocument();
            document.DocumentPaginator.PageSize = new Size(A4LandscapeWidth, A4LandscapeHeight);

            AddFrameworkElementAsTextOnlyPrintPage(
                document,
                VedomostPage,
                A4LandscapeWidth,
                A4LandscapeHeight,
                0);

            AddFrameworkElementAsTextOnlyPrintPage(
                document,
                VedomostPage2,
                A4LandscapeWidth,
                A4LandscapeHeight,
                0);

            return document;
        }

        private void AddFrameworkElementAsTextOnlyPrintPage(
            FixedDocument document,
            FrameworkElement sourceRoot,
            double pageWidth,
            double pageHeight,
            double shiftY)
        {
            sourceRoot.UpdateLayout();

            double sourceWidth = GetElementPrintWidth(sourceRoot);
            double sourceHeight = GetElementPrintHeight(sourceRoot);

            if (sourceWidth <= 0)
                sourceWidth = pageWidth;

            if (sourceHeight <= 0)
                sourceHeight = pageHeight;

            var fixedPage = new FixedPage
            {
                Width = pageWidth,
                Height = pageHeight,
                Background = Brushes.White
            };

            double scaleX = pageWidth / sourceWidth;
            double scaleY = pageHeight / sourceHeight;
            double scale = Math.Min(scaleX, scaleY);

            double offsetX = (pageWidth - sourceWidth * scale) / 2.0;
            double offsetY = (pageHeight - sourceHeight * scale) / 2.0 + shiftY;

            var printCanvas = new Canvas
            {
                Width = sourceWidth,
                Height = sourceHeight,
                Background = Brushes.Transparent,
                LayoutTransform = new ScaleTransform(scale, scale)
            };

            FixedPage.SetLeft(printCanvas, offsetX);
            FixedPage.SetTop(printCanvas, offsetY);

            CopyTextElementsFromRoot(sourceRoot, sourceRoot, printCanvas, scale);

            fixedPage.Children.Add(printCanvas);

            fixedPage.Measure(new Size(pageWidth, pageHeight));
            fixedPage.Arrange(new Rect(new Size(pageWidth, pageHeight)));
            fixedPage.UpdateLayout();

            var pageContent = new PageContent();
            ((IAddChild)pageContent).AddChild(fixedPage);
            document.Pages.Add(pageContent);
        }

        private void CopyTextElementsFromRoot(
    FrameworkElement root,
    DependencyObject current,
    Canvas printCanvas,
    double scale)
        {
            int count = VisualTreeHelper.GetChildrenCount(current);

            for (int i = 0; i < count; i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(current, i);

                if (child is Border border && border.Child is TextBlock textBlock)
                {
                    AddTextBlockCopyToPrintCanvas(root, border, textBlock, printCanvas, scale);
                }

                CopyTextElementsFromRoot(root, child, printCanvas, scale);
            }
        }

        private void AddTextBlockCopyToPrintCanvas(
    FrameworkElement root,
    Border sourceBorder,
    TextBlock sourceText,
    Canvas printCanvas,
    double scale)
        {
            string text = sourceText.Text ?? "";

            if (string.IsNullOrWhiteSpace(text))
                return;

            double safeScale = scale <= 0 ? 1 : scale;

            Point position;

            try
            {
                GeneralTransform transform = sourceBorder.TransformToAncestor(root);
                position = transform.Transform(new Point(0, 0));
            }
            catch
            {
                position = new Point(GetLeft(sourceBorder), GetTop(sourceBorder));
            }

            // ВАЖНО:
            // printCanvas масштабируется через ScaleTransform.
            // Поэтому FontSize делим на scale, чтобы после масштабирования
            // на печати остался тот размер, который указан в программе.
            double printFontSize = sourceText.FontSize / safeScale;

            double printLineHeight = sourceText.LineHeight;

            if (!double.IsNaN(printLineHeight) && printLineHeight > 0)
                printLineHeight = printLineHeight / safeScale;

            // Копируем padding из исходного текста
            Thickness padding = new Thickness(
                sourceText.Padding.Left,
                sourceText.Padding.Top,
                sourceText.Padding.Right,
                sourceText.Padding.Bottom);

            var copiedText = new TextBlock
            {
                Text = text,
                FontFamily = sourceText.FontFamily,
                FontSize = printFontSize,
                FontWeight = sourceText.FontWeight,
                FontStyle = sourceText.FontStyle,
                Foreground = Brushes.Black,
                TextAlignment = sourceText.TextAlignment,
                VerticalAlignment = sourceText.VerticalAlignment,
                HorizontalAlignment = sourceText.HorizontalAlignment,
                TextWrapping = sourceText.TextWrapping,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Padding = padding,
                Margin = new Thickness(0)
            };

            if (!double.IsNaN(printLineHeight) && printLineHeight > 0)
                copiedText.LineHeight = printLineHeight;

            var copiedBorder = new Border
            {
                Width = sourceBorder.Width,
                Height = sourceBorder.Height,
                Background = Brushes.Transparent,
                Padding = new Thickness(0),
                Margin = new Thickness(0),
                Child = copiedText,
                ClipToBounds = true
            };

            Canvas.SetLeft(copiedBorder, position.X);
            Canvas.SetTop(copiedBorder, position.Y);
            Panel.SetZIndex(copiedBorder, Panel.GetZIndex(sourceBorder));

            printCanvas.Children.Add(copiedBorder);
        }

        private double GetElementPrintWidth(FrameworkElement element)
        {
            if (element.ActualWidth > 0 && !double.IsNaN(element.ActualWidth))
                return element.ActualWidth;

            if (element.Width > 0 && !double.IsNaN(element.Width))
                return element.Width;

            element.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));

            if (element.DesiredSize.Width > 0)
                return element.DesiredSize.Width;

            return 1;
        }

        private double GetElementPrintHeight(FrameworkElement element)
        {
            if (element.ActualHeight > 0 && !double.IsNaN(element.ActualHeight))
                return element.ActualHeight;

            if (element.Height > 0 && !double.IsNaN(element.Height))
                return element.Height;

            element.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));

            if (element.DesiredSize.Height > 0)
                return element.DesiredSize.Height;

            return 1;
        }

        private void LoadBackground(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp",
                Title = "Выберите фон лицевой стороны"
            };

            if (dialog.ShowDialog() != true)
                return;

            _manualBackgroundLoaded = true;

            string path = dialog.FileName;

            if (SelectedDocType == "Диплом")
            {
                DiplomaBgLeft.Source = LoadBitmapPart(path, true);
                DiplomaBgRight.Source = LoadBitmapPart(path, false);
            }
            else
            {
                VedomostBg.Source = LoadBitmap(path);
            }

            ApplyMode();
            _lastBuildKey = "";
            ScheduleBuild();
        }

        private void LoadBackground2(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp",
                Title = "Выберите фон оборотной стороны"
            };

            if (dialog.ShowDialog() != true)
                return;

            _manualBackgroundLoaded = true;

            if (SelectedDocType == "Диплом")
                DiplomaBgRight.Source = LoadBitmap(dialog.FileName);
            else
                VedomostBg2.Source = LoadBitmap(dialog.FileName);

            ApplyMode();
            _lastBuildKey = "";
            ScheduleBuild();
        }

        private static BitmapImage LoadBitmap(string path)
        {
            var img = new BitmapImage();
            img.BeginInit();
            img.CacheOption = BitmapCacheOption.OnLoad;
            img.UriSource = new Uri(path, UriKind.Absolute);
            img.EndInit();
            img.Freeze();
            return img;
        }

        private static ImageSource LoadBitmapPart(string path, bool leftPart)
        {
            var img = LoadBitmap(path);

            if (img.PixelWidth <= img.PixelHeight)
                return img;

            int halfWidth = img.PixelWidth / 2;
            int x = leftPart ? 0 : halfWidth;

            var cropped = new CroppedBitmap(img, new Int32Rect(x, 0, halfWidth, img.PixelHeight));
            cropped.Freeze();

            return cropped;
        }

        private void PrevStudent(object sender, RoutedEventArgs e)
        {
            int count = AppData.Students?.Count ?? 0;

            if (count == 0)
            {
                UpdateStudentLabel();
                return;
            }

            _currentStudentIndex--;

            if (_currentStudentIndex < 0)
                _currentStudentIndex = count - 1;

            _lastBuildKey = "";
            ScheduleBuild();
        }

        private void NextStudent(object sender, RoutedEventArgs e)
        {
            int count = AppData.Students?.Count ?? 0;

            if (count == 0)
            {
                UpdateStudentLabel();
                return;
            }

            _currentStudentIndex++;

            if (_currentStudentIndex >= count)
                _currentStudentIndex = 0;

            _lastBuildKey = "";
            ScheduleBuild();
        }

        private void AddNewCustomTextField()
        {
            Canvas canvas = GetActiveCustomCanvas(out string canvasName);

            AddNewCustomTextField(
                canvasName,
                120,
                120,
                openEditorImmediately: true);
        }

        private void AddNewCustomTextField(string canvasName)
        {
            AddNewCustomTextField(
                canvasName,
                120,
                120,
                openEditorImmediately: true);
        }

        private void AddNewCustomTextField(
            string canvasName,
            double x,
            double y,
            bool openEditorImmediately)
        {
            string mode = CurrentModeKey;
            string key = "custom_" + DateTime.Now.Ticks;

            Canvas canvas = GetCanvasByName(canvasName);

            var field = new CustomTextField
            {
                Mode = mode,
                Key = key,
                Text = "Новый текст",
                X = x,
                Y = y,
                Width = 220,
                Height = 22,
                FontSize = DefaultFontSize,
                CanvasName = canvasName
            };

            _customTextFields[key] = field;

            SaveCustomTextFields();

            Border border = AddEditableText(
                canvas,
                field.Key,
                field.X,
                field.Y,
                field.Width,
                field.Height,
                field.Text,
                false,
                false);

            var item = _fields.FirstOrDefault(f => f.Host == border);

            if (item != null)
            {
                SelectField(item);

                if (openEditorImmediately)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        OpenEditor(item);
                    }), DispatcherPriority.Background);
                }
            }
        }

        private void DeleteSelectedCustomTextField()
        {
            DeleteSelectedFieldSmart();
        }

        private void DeleteSelectedFieldSmart()
        {
            if (_selectedField == null)
                return;

            string key = _selectedField.Key;

            if (key.StartsWith("custom_", StringComparison.OrdinalIgnoreCase))
            {
                if (_selectedField.Host.Parent is Canvas customCanvas)
                    customCanvas.Children.Remove(_selectedField.Host);

                _fields.Remove(_selectedField);
                _customTextFields.Remove(key);

                _selectedField = null;

                SaveCustomTextFields();

                _lastBuildKey = "";
                ScheduleBuild();

                return;
            }

            string deletedKey = MakeDeletedFieldKey(key);
            _deletedFieldKeys.Add(deletedKey);

            if (_selectedField.Host.Parent is Canvas canvas)
                canvas.Children.Remove(_selectedField.Host);

            _fields.Remove(_selectedField);
            _selectedField = null;

            SaveDeletedFields();

            _lastBuildKey = "";
            ScheduleBuild();
        }

        private void RenderCustomTextFields()
        {
            string mode = CurrentModeKey;

            foreach (var item in _customTextFields.Values.Where(x => x.Mode == mode))
            {
                Canvas canvas = GetCanvasByName(item.CanvasName);

                AddEditableText(
                    canvas,
                    item.Key,
                    item.X,
                    item.Y,
                    item.Width,
                    item.Height,
                    item.Text,
                    false,
                    false);
            }
        }

        private Canvas GetActiveCustomCanvas(out string canvasName)
        {
            if (SelectedDocType == "Диплом")
            {
                canvasName = "diploma_left";
                return DiplomaLeftCanvas;
            }

            canvasName = "vedomost_1";
            return VedomostCanvas;
        }

        private Canvas GetCanvasByName(string canvasName)
        {
            return canvasName switch
            {
                "vedomost_2" => VedomostCanvas2,
                "diploma_left" => DiplomaLeftCanvas,
                "diploma_right" => DiplomaRightCanvas,
                _ => VedomostCanvas
            };
        }

        private void SaveCustomPositionIfNeeded()
        {
            if (_selectedField == null)
                return;

            if (!_selectedField.Key.StartsWith("custom_", StringComparison.OrdinalIgnoreCase))
                return;

            if (!_customTextFields.TryGetValue(_selectedField.Key, out CustomTextField? custom))
                return;

            custom.X = GetLeft(_selectedField.Host);
            custom.Y = GetTop(_selectedField.Host);
            custom.Width = _selectedField.Host.Width;
            custom.Height = _selectedField.Host.Height;
            custom.FontSize = _selectedField.Text.FontSize;
            custom.Text = _selectedField.Text.Text;

            SaveCustomTextFields();
        }

        private static string Escape(string text)
        {
            if (text == null)
                return "";

            return Convert.ToBase64String(Encoding.UTF8.GetBytes(text));
        }

        private static string Unescape(string text)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(text))
                    return "";

                return Encoding.UTF8.GetString(Convert.FromBase64String(text));
            }
            catch
            {
                return text;
            }
        }

        private static double ParseDoubleSafe(string text, double fallback = 0)
        {
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                return value;

            return fallback;
        }

        private static string GetBestFio(Student? s)
        {
            if (s == null)
                return "";

            if (!string.IsNullOrWhiteSpace(s.WordFIO))
                return s.WordFIO;

            return s.FIO ?? "";
        }

        private static string BuildSpecialityText(Student? s)
        {
            if (s == null)
                return "";

            if (!string.IsNullOrWhiteSpace(s.SpecialityCode) &&
                !string.IsNullOrWhiteSpace(s.SpecialityName))
            {
                return $"{s.SpecialityCode} - {s.SpecialityName}";
            }

            if (!string.IsNullOrWhiteSpace(s.SpecialityName))
                return s.SpecialityName;

            if (!string.IsNullOrWhiteSpace(s.SpecialityCode))
                return s.SpecialityCode;

            return CleanSpeciality(s.Speciality);
        }

        private static string BuildQualificationText(Student? s)
        {
            if (s == null)
                return "";

            if (!string.IsNullOrWhiteSpace(s.QualificationCode) &&
                !string.IsNullOrWhiteSpace(s.QualificationName))
            {
                return $"{s.QualificationCode} - {s.QualificationName}";
            }

            if (!string.IsNullOrWhiteSpace(s.QualificationName))
                return s.QualificationName;

            if (!string.IsNullOrWhiteSpace(s.QualificationCode))
                return s.QualificationCode;

            return ExtractQualification(s.Speciality, s.Group);
        }

        private static string GetInstitutionForPrint()
        {
            return "Усть-Каменогорский колледж экономики и техники";
        }

        private static string GetDiplomaNumber(Student? s)
        {
            if (s == null)
                return "";

            object obj = s;

            string[] names =
            {
                "DiplomaNumber",
                "DiplomaNo",
                "TkbNumber",
                "TkbNo",
                "DocNumber"
            };

            foreach (string name in names)
            {
                PropertyInfo? prop = obj.GetType().GetProperty(name);

                if (prop == null)
                    continue;

                object? value = prop.GetValue(obj);

                if (value == null)
                    continue;

                string text = value.ToString() ?? "";

                if (!string.IsNullOrWhiteSpace(text) && text != "0")
                    return text.Trim();
            }

            return "";
        }

        private static string CleanSpeciality(string? speciality)
        {
            if (string.IsNullOrWhiteSpace(speciality))
                return "";

            string text = speciality.Trim();

            int idx = text.IndexOf("квалификация", StringComparison.OrdinalIgnoreCase);

            if (idx > 0)
                text = text.Substring(0, idx).Trim().Trim(',', ';');

            idx = text.IndexOf("біліктілік", StringComparison.OrdinalIgnoreCase);

            if (idx > 0)
                text = text.Substring(0, idx).Trim().Trim(',', ';');

            return text;
        }

        private static string ExtractQualification(string? speciality, string? fallback)
        {
            if (string.IsNullOrWhiteSpace(speciality))
                return fallback ?? "";

            string text = speciality.Trim();

            int idx = text.IndexOf("квалификация", StringComparison.OrdinalIgnoreCase);

            if (idx >= 0)
                return text.Substring(idx + "квалификация".Length).Trim().Trim(':', ',', ';');

            idx = text.IndexOf("біліктілік", StringComparison.OrdinalIgnoreCase);

            if (idx >= 0)
                return text.Substring(idx + "біліктілік".Length).Trim().Trim(':', ',', ';');

            return fallback ?? "";
        }

        private static double GetLeft(UIElement element)
        {
            double value = Canvas.GetLeft(element);
            return double.IsNaN(value) ? 0 : value;
        }

        private static double GetTop(UIElement element)
        {
            double value = Canvas.GetTop(element);
            return double.IsNaN(value) ? 0 : value;
        }
    }
}