using System.Windows;
using Velopack;

namespace SuneungMarker;

public partial class App : Application
{
    [STAThread]
    public static void Main(string[] args)
    {
        VelopackApp.Build().Run();

        var app = new App();
        app.InitializeComponent();
        app.StartupUri = new Uri("MainWindow.xaml", UriKind.Relative);
        app.Run();
    }
}
