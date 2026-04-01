using System.IO;
using System.Windows;
using Newtonsoft.Json;

namespace SuneungMarker;

public partial class SettingsWindow : Window
{
    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SuneungMarker", "settings.json");

    public SettingsWindow()
    {
        InitializeComponent();
        LoadSettings();
    }

    private void LoadSettings()
    {
        var s = AppSettings.Load();
        TxtClaude.Text = s.ClaudeKey;
        TxtGpt.Text = s.GptKey;
        TxtGemini.Text = s.GeminiKey;
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        var s = new AppSettings
        {
            ClaudeKey = TxtClaude.Text.Trim(),
            GptKey = TxtGpt.Text.Trim(),
            GeminiKey = TxtGemini.Text.Trim()
        };
        s.Save();

        // AiAnalyzer에 반영
        AiAnalyzer.ClaudeKey = s.ClaudeKey;
        AiAnalyzer.GptKey = s.GptKey;
        AiAnalyzer.GeminiKey = s.GeminiKey;

        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}

public class AppSettings
{
    public string ClaudeKey { get; set; } = "";
    public string GptKey { get; set; } = "";
    public string GeminiKey { get; set; } = "";

    private static readonly string Dir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SuneungMarker");
    private static readonly string FilePath = Path.Combine(Dir, "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonConvert.DeserializeObject<AppSettings>(File.ReadAllText(FilePath)) ?? new();
        }
        catch { }
        return new AppSettings();
    }

    public void Save()
    {
        Directory.CreateDirectory(Dir);
        File.WriteAllText(FilePath, JsonConvert.SerializeObject(this, Formatting.Indented));
    }
}
