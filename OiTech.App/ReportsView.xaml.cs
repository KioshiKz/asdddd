using OiTech.App.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace OiTech.App
{
    public partial class ReportsView : UserControl
    {
        public ReportsView()
        {
            InitializeComponent();
            LoadReport();
        }

        private void LoadReport()
        {
            TotalStudents.Text = AppData.Students.Count.ToString();
            TotalSubjects.Text = AppData.SubjectNames.Count.ToString();
            PrintedCount.Text = AppData.PrintedCount.ToString();

            var allGrades = new List<double>();
            var subjectGrades = new Dictionary<string, List<double>>();

            foreach (var s in AppData.Students)
            {
                foreach (var g in s.Grades)
                {
                    if (double.TryParse(g.Grade, out double val))
                    {
                        allGrades.Add(val);
                        if (!subjectGrades.ContainsKey(g.Subject))
                            subjectGrades[g.Subject] = new List<double>();
                        subjectGrades[g.Subject].Add(val);
                    }
                }
            }

            AvgGrade.Text = allGrades.Count > 0
                ? Math.Round(allGrades.Average(), 2).ToString("F2")
                : "—";

            SubjectBarsPanel.Children.Clear();
            foreach (var kvp in subjectGrades)
            {
                double avg = Math.Round(kvp.Value.Average(), 2);
                double barWidth = Math.Min(avg / 5.0 * 100, 100);

                var row = new DockPanel { Margin = new Thickness(0, 0, 0, 8) };

                var label = new TextBlock
                {
                    Text = kvp.Key,
                    Width = 200,
                    Foreground = (SolidColorBrush)FindResource("TextDimBrush"),
                    FontSize = 13,
                    VerticalAlignment = VerticalAlignment.Center
                };
                DockPanel.SetDock(label, Dock.Left);
                row.Children.Add(label);

                var valLabel = new TextBlock
                {
                    Text = avg.ToString("F2"),
                    Width = 50,
                    Foreground = (SolidColorBrush)FindResource("TextWhiteBrush"),
                    FontSize = 13,
                    FontWeight = FontWeights.SemiBold,
                    TextAlignment = TextAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center
                };
                DockPanel.SetDock(valLabel, Dock.Right);
                row.Children.Add(valLabel);

                var barBg = new Border
                {
                    Height = 10,
                    CornerRadius = new CornerRadius(5),
                    Background = new SolidColorBrush(Color.FromRgb(30, 41, 59)),
                    VerticalAlignment = VerticalAlignment.Center
                };
                var barFill = new Border
                {
                    Height = 10,
                    CornerRadius = new CornerRadius(5),
                    Width = barWidth * 3,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Background = (SolidColorBrush)FindResource("AccentBlueBrush")
                };
                var barGrid = new Grid();
                barGrid.Children.Add(barBg);
                barGrid.Children.Add(barFill);
                row.Children.Add(barGrid);

                SubjectBarsPanel.Children.Add(row);
            }

            if (subjectGrades.Count == 0)
            {
                SubjectBarsPanel.Children.Add(new TextBlock
                {
                    Text = LocalizationManager.Get("NoGradesData"),
                    Foreground = (SolidColorBrush)FindResource("TextDimBrush"),
                    FontSize = 13
                });
            }
        }
    }
}
