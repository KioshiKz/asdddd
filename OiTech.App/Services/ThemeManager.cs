using System;
using System.Linq;
using System.Windows;

namespace OiTech.App.Services
{
    public static class ThemeManager
    {
        public static bool IsDark { get; private set; } = true;

        public static void Toggle()
        {
            IsDark = !IsDark;
            Apply(IsDark ? "DarkTheme" : "LightTheme");
        }

        private static void Apply(string name)
        {
            var dicts = Application.Current.Resources.MergedDictionaries;
            var existing = dicts.FirstOrDefault(d => d.Source?.OriginalString.Contains("Theme") == true);
            var newDict = new ResourceDictionary
            {
                Source = new Uri($"Themes/{name}.xaml", UriKind.Relative)
            };
            if (existing != null)
            {
                int idx = dicts.IndexOf(existing);
                dicts.Remove(existing);
                dicts.Insert(idx, newDict);
            }
            else
            {
                dicts.Insert(0, newDict);
            }
        }
    }
}
