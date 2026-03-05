using System.Windows;
using System.Windows.Media;
using Wpf.Ui.Appearance;

namespace SCCMAdPrep;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Global accent color: #006ABB instead of default purple
        ApplicationAccentColorManager.Apply(
            Color.FromRgb(0x00, 0x6A, 0xBB),
            ApplicationTheme.Light
        );
    }
}
