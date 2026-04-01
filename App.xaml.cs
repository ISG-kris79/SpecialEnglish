using System.Windows;
using System.Windows.Threading;
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

        app.DispatcherUnhandledException += (s, e) =>
        {
            MessageBox.Show($"오류 발생:\n{e.Exception.Message}\n\n{e.Exception.InnerException?.Message}",
                "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        };

        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
        {
            if (e.ExceptionObject is Exception ex)
                MessageBox.Show($"심각한 오류:\n{ex.Message}", "오류");
        };

        TaskScheduler.UnobservedTaskException += (s, e) =>
        {
            e.SetObserved();
        };

        app.Run();
    }
}
