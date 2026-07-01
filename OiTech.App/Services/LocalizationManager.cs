using System;
using System.Linq;
using System.Windows;

namespace OiTech.App.Services
{
    public static class LocalizationManager
    {
        private static readonly string[] DictNames  = { "Strings_RU", "Strings_KK", "Strings_EN" };
        private static readonly string[] LangCodes  = { "РУС", "ҚАЗ", "ENG" };
        private static int _index = 0;

        public static bool IsRussian => _index == 0;

        public static string CurrentCode => LangCodes[_index];

        public static void Toggle()
        {
            _index = (_index + 1) % DictNames.Length;
            Apply(DictNames[_index]);
        }

        public static string Get(string key)
            => Application.Current.TryFindResource(key) as string ?? key;

        private static void Apply(string name)
        {
            var dicts = Application.Current.Resources.MergedDictionaries;
            var existing = dicts.FirstOrDefault(d => d.Source?.OriginalString.Contains("Strings_") == true);
            var newDict = new ResourceDictionary
            {
                Source = new Uri($"Resources/{name}.xaml", UriKind.Relative)
            };
            if (existing != null)
            {
                int idx = dicts.IndexOf(existing);
                dicts.Remove(existing);
                dicts.Insert(idx, newDict);
            }
            else
            {
                dicts.Add(newDict);
            }
        }
    }
}
