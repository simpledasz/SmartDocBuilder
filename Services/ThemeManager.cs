
using System;
using Avalonia;
using Avalonia.Markup.Xaml;

namespace SmartDocBuilder.Services
{
    public class ThemeManager
    {
        public enum Theme { Light, Dark }

        public static void SetTheme(Theme theme)
        {
            var app = Application.Current;
            if (app == null) return;

            var themeUri = theme switch
            {
                Theme.Light => new Uri("avares://SmartDocBuilderGUI/Themes/LightTheme.axaml"),
                Theme.Dark => new Uri("avares://SmartDocBuilderGUI/Themes/DarkTheme.axaml"),
                _ => throw new ArgumentOutOfRangeException(nameof(theme), theme, null)
            };

            if (app.Styles[0] is IResourceDictionary currentTheme)
            {
                currentTheme.MergedDictionaries[0] = new ResourceInclude(themeUri)
                {
                    Source = themeUri
                };
            }
        }
    }
}
