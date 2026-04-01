using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Win32;

namespace SuneungMarker;

public class PassageData
{
    public string Key { get; set; } = "";
    public int Chapter { get; set; }
    public int Number { get; set; }
    public string English { get; set; } = "";
    public string Korean { get; set; } = "";
    public string Footnotes { get; set; } = "";
    public string TypeHint { get; set; } = "";
    public List<string> VocabWords { get; set; } = new();
    public string Numbered { get; set; } = "";
    public List<MarkItem> Marks { get; set; } = new();
}

public class MarkItem
{
    [JsonProperty("type")] public string Type { get; set; } = "";
    [JsonProperty("start")] public int Start { get; set; }
    [JsonProperty("end")] public int End { get; set; }
    [JsonProperty("text")] public string Text { get; set; } = "";
}

public partial class MainWindow : Window
{
    private static readonly string[] Circled = { "①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩","⑪","⑫","⑬","⑭","⑮","⑯","⑰","⑱","⑲","⑳" };
    private static readonly Regex SplitRx = new(@"(?<=[.!?])\s+(?=[A-Z""'\(])|(?<=\.[\u201d\u2019'""'])\s+(?=[A-Z])", RegexOptions.Compiled);

    private List<PassageData> _passages = new();
    private int _currentIndex = 0;
    private string _dataFolder = "";
    private string _file1Path = "";
    private string _file2Path = "";

    public MainWindow()
    {
        InitializeComponent();
        LoadApiKeys();
    }

    private void LoadApiKeys()
    {
        var s = AppSettings.Load();
        AiAnalyzer.ClaudeKey = s.ClaudeKey;
        AiAnalyzer.GptKey = s.GptKey;
        AiAnalyzer.GeminiKey = s.GeminiKey;
    }

    // ═══════════════════════════════
    //  업로드 화면
    // ═══════════════════════════════

    private void BtnLoad1_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "한글 파일|*.hwp|텍스트 파일|*.txt|모든 파일|*.*",
            Title = "1.hwp (영어 지문) 선택"
        };
        if (dlg.ShowDialog() == true)
        {
            _file1Path = dlg.FileName;
            TxtFile1.Text = $"✅ {Path.GetFileName(_file1Path)}";
            TxtFile1.Foreground = new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50));
            CheckReady();
        }
    }

    private void BtnLoad2_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "한글 파일|*.hwp|텍스트 파일|*.txt|모든 파일|*.*",
            Title = "2.hwp (어휘) 선택"
        };
        if (dlg.ShowDialog() == true)
        {
            _file2Path = dlg.FileName;
            TxtFile2.Text = $"✅ {Path.GetFileName(_file2Path)}";
            TxtFile2.Foreground = new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50));
            CheckReady();
        }
    }

    private void CheckReady()
    {
        BtnStart.IsEnabled = !string.IsNullOrEmpty(_file1Path) && !string.IsNullOrEmpty(_file2Path);
    }

    private int _currentPercent = 0;

    private async Task AnimateProgress(int targetPercent, string message)
    {
        while (_currentPercent < targetPercent)
        {
            _currentPercent++;
            var pct = _currentPercent;
            var msg = message;
            Dispatcher.Invoke(() =>
            {
                PrgLoad.Visibility = Visibility.Visible;
                PrgIndeterminate.Visibility = pct < 100 ? Visibility.Visible : Visibility.Collapsed;
                PrgLoad.Value = pct;
                TxtLoadStatus.Text = $"{pct}% - {msg}";
                TxtLoadStatus.Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66));
            });
            await Task.Delay(15);
        }
    }

    private async void BtnStart_Click(object sender, RoutedEventArgs e)
    {
        BtnStart.IsEnabled = false;
        _currentPercent = 0;
        PrgLoad.Visibility = Visibility.Visible;
        PrgIndeterminate.Visibility = Visibility.Visible;
        PrgLoad.Value = 0;

        try
        {
            _dataFolder = Path.GetDirectoryName(_file1Path) ?? "";
            string fullText, vocabText;

            if (_file1Path.EndsWith(".hwp", StringComparison.OrdinalIgnoreCase))
            {
                _ = AnimateProgress(15, "영어 지문 HWP 읽는 중...");
                fullText = await Task.Run(() => HwpReader.ExtractText(_file1Path));
                await AnimateProgress(40, "영어 지문 변환 완료");
            }
            else
            {
                await AnimateProgress(5, "영어 지문 로딩...");
                fullText = await File.ReadAllTextAsync(_file1Path, Encoding.UTF8);
                await AnimateProgress(40, "영어 지문 로드 완료");
            }

            if (_file2Path.EndsWith(".hwp", StringComparison.OrdinalIgnoreCase))
            {
                _ = AnimateProgress(55, "어휘 HWP 읽는 중...");
                vocabText = await Task.Run(() => HwpReader.ExtractText(_file2Path));
                await AnimateProgress(70, "어휘 변환 완료");
            }
            else
            {
                await AnimateProgress(50, "어휘 로딩...");
                vocabText = await File.ReadAllTextAsync(_file2Path, Encoding.UTF8);
                await AnimateProgress(70, "어휘 로드 완료");
            }

            var f1txt = Path.Combine(_dataFolder, "1_full.txt");
            var f2txt = Path.Combine(_dataFolder, "2_text.txt");
            if (!File.Exists(f1txt)) await File.WriteAllTextAsync(f1txt, fullText, Encoding.UTF8);
            if (!File.Exists(f2txt)) await File.WriteAllTextAsync(f2txt, vocabText, Encoding.UTF8);

            await AnimateProgress(80, "데이터 파싱 중...");
            LoadDataFromText(fullText, vocabText);

            await AnimateProgress(90, "마킹 데이터 확인 중...");
            var mf = Path.Combine(_dataFolder, "marks_v2.json");
            if (File.Exists(mf)) LoadMarks(mf);

            await AnimateProgress(100, "완료!");

            // 에디터 화면 전환
            PnlWelcome.Visibility = Visibility.Collapsed;
            PnlEditor.Visibility = Visibility.Visible;
            PnlToolbar.Visibility = Visibility.Visible;
            PnlNavButtons.Visibility = Visibility.Visible;
            PnlStatusBar.Visibility = Visibility.Visible;

            ShowPassage(0);
        }
        catch (Exception ex)
        {
            PrgLoad.Visibility = Visibility.Collapsed;
            PrgIndeterminate.Visibility = Visibility.Collapsed;
            TxtLoadStatus.Text = $"오류: {ex.Message}";
            TxtLoadStatus.Foreground = new SolidColorBrush(Color.FromRgb(0xE8, 0x38, 0x4F));
            BtnStart.IsEnabled = true;
        }
    }

    private static string FindPython()
    {
        // 직접 설치된 Python 찾기
        var candidates = new[]
        {
            @"C:\Users\COFFE\AppData\Local\Python\pythoncore-3.14-64\python.exe",
            @"C:\Python314\python.exe",
            @"C:\Python312\python.exe",
            @"C:\Python311\python.exe",
        };
        foreach (var p in candidates)
            if (File.Exists(p)) return p;

        // PATH에서 찾기
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "where", Arguments = "python",
                RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true
            };
            var proc = Process.Start(psi);
            var output = proc?.StandardOutput.ReadToEnd() ?? "";
            proc?.WaitForExit(5000);
            foreach (var line in output.Split('\n'))
            {
                var path = line.Trim();
                if (path.EndsWith(".exe") && !path.Contains("WindowsApps") && File.Exists(path))
                    return path;
            }
        }
        catch { }

        return "python"; // fallback
    }

    private string ExtractHwpText(string hwpPath)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"hwp_{Guid.NewGuid():N}");

        try
        {
            Directory.CreateDirectory(tempDir);

            var script = $@"
import sys, os
sys.argv = ['hwp5html', '--output', r'{tempDir}', r'{hwpPath}']
from hwp5.hwp5html import main
main()
";
            var scriptPath = Path.Combine(tempDir, "convert.py");
            File.WriteAllText(scriptPath, script, Encoding.UTF8);

            // python 절대경로 찾기
            var pythonPath = FindPython();

            var psi = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = $"\"{scriptPath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = tempDir
            };

            var proc = Process.Start(psi);
            var stderr = proc?.StandardError.ReadToEnd() ?? "";
            proc?.WaitForExit(120000);

            var xhtmlPath = Path.Combine(tempDir, "index.xhtml");
            if (!File.Exists(xhtmlPath))
                throw new Exception($"HWP 변환 실패 (python={pythonPath}, exit={proc?.ExitCode}, err={stderr[..Math.Min(200, stderr.Length)]})");

            var html = File.ReadAllText(xhtmlPath, Encoding.UTF8);

            var paras = Regex.Matches(html, @"<p[^>]*>(.*?)</p>", RegexOptions.Singleline);
            var sb = new StringBuilder();
            foreach (Match m in paras)
            {
                var text = Regex.Replace(m.Groups[1].Value, @"<[^>]+>", "").Replace("&#13;", "").Trim();
                if (!string.IsNullOrEmpty(text))
                    sb.AppendLine(text);
            }
            return sb.ToString();
        }
        finally
        {
            // 파일 잠김 방지: 약간 대기 후 삭제
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Thread.Sleep(500);
                Directory.Delete(tempDir, true);
            }
            catch { /* 삭제 실패해도 무시 */ }
        }
    }

    // ═══════════════════════════════
    //  데이터 로딩
    // ═══════════════════════════════

    private void LoadDataFromText(string fullText, string vocabText)
    {
        // 지문 파싱
        var parts = Regex.Split(fullText, @"(\[2027수능특강\(영어\)\s+\d+강\s+\d+번\])");
        var entries = new Dictionary<string, PassageData>();

        for (int i = 1; i < parts.Length; i += 2)
        {
            var m = Regex.Match(parts[i], @"(\d+)강\s+(\d+)번");
            if (!m.Success) continue;
            int ch = int.Parse(m.Groups[1].Value);
            int num = int.Parse(m.Groups[2].Value);
            string key = $"{ch}-{num}";

            string body = i + 1 < parts.Length ? parts[i + 1].Trim() : "";
            var lines = body.Split('\n');
            string typeHint = "";
            if (lines.Length > 0 && lines[0].Trim().StartsWith("다음"))
            {
                typeHint = lines[0].Trim();
                body = string.Join('\n', lines.Skip(1)).Trim();
            }

            var engLines = new List<string>();
            var fnLines = new List<string>();
            foreach (var line in body.Split('\n'))
            {
                var l = line.Trim();
                if (l.StartsWith("*")) fnLines.Add(l);
                else if (!Regex.IsMatch(l, @"^\d+강$") && l.Length > 0) engLines.Add(l);
            }

            entries[key] = new PassageData
            {
                Key = key, Chapter = ch, Number = num,
                English = string.Join(" ", engLines).Trim(),
                Footnotes = string.Join(" ", fnLines).Trim(),
                TypeHint = typeHint
            };
        }

        // 한국어 해석 파싱
        string[] textLines = fullText.Split('\n');
        int curCh = 0;
        for (int i = 0; i < textLines.Length; i++)
        {
            var l = textLines[i].Trim();
            var cm = Regex.Match(l, @"^(\d+)강$");
            if (cm.Success) { curCh = int.Parse(cm.Groups[1].Value); continue; }

            var nm = Regex.Match(l, @"^(\d{1,2})$");
            if (nm.Success && curCh > 0)
            {
                int num2 = int.Parse(nm.Groups[1].Value);
                string key2 = $"{curCh}-{num2}";
                // [해석] 찾기
                for (int j = i + 1; j < textLines.Length; j++)
                {
                    if (textLines[j].Trim() == "[해석]")
                    {
                        var sb = new StringBuilder();
                        for (int k = j + 1; k < textLines.Length; k++)
                        {
                            var tl = textLines[k].Trim();
                            if (tl is "[해설]" or "[정답]" || Regex.IsMatch(tl, @"^\d{1,2}$") || Regex.IsMatch(tl, @"^\d+강$"))
                                break;
                            if (tl.Length > 0) sb.Append(tl + " ");
                        }
                        if (entries.ContainsKey(key2))
                            entries[key2].Korean = sb.ToString().Trim();
                        break;
                    }
                    var tl2 = textLines[j].Trim();
                    if (Regex.IsMatch(tl2, @"^\d{1,2}$") || Regex.IsMatch(tl2, @"^\d+강$"))
                        break;
                }
            }
        }

        // 어휘 파싱
        var vocab = new Dictionary<string, List<string>>();
        int vch = 0, vnum = 0;
        var vwords = new List<string>();
        foreach (var line in vocabText.Split('\n'))
        {
            var l = line.Trim();
            var vm = Regex.Match(l, @"^(\d+)강\s*$");
            if (vm.Success)
            {
                if (vch > 0 && vnum > 0) vocab[$"{vch}-{vnum}"] = new List<string>(vwords);
                vch = int.Parse(vm.Groups[1].Value); vnum = 0; vwords.Clear(); continue;
            }
            var vn = Regex.Match(l, @"^(\d{1,2})$");
            if (vn.Success && vch > 0)
            {
                if (vnum > 0) vocab[$"{vch}-{vnum}"] = new List<string>(vwords);
                vnum = int.Parse(vn.Groups[1].Value); vwords.Clear(); continue;
            }
            if (l is "Words & Phrases" or "part 1" or "part 2" or "유형편" or "소재편" or "") continue;
            if (l.Length > 0 && vch > 0 && vnum > 0)
                vwords.Add(l);
        }
        if (vch > 0 && vnum > 0) vocab[$"{vch}-{vnum}"] = new List<string>(vwords);

        _passages = entries.Values
            .OrderBy(p => p.Chapter).ThenBy(p => p.Number)
            .ToList();

        foreach (var p in _passages)
        {
            if (vocab.TryGetValue(p.Key, out var vw)) p.VocabWords = vw;
            p.Numbered = MakeNumbered(p.English, new HashSet<int>());
        }

        TxtStatus.Text = $"{_passages.Count}개 지문 로드 완료";
    }

    private string MakeNumbered(string english, HashSet<int> orderPos)
    {
        var sents = SplitRx.Split(english);
        var sb = new StringBuilder();
        for (int si = 0; si < sents.Length; si++)
        {
            if (orderPos.Contains(si + 1)) sb.Append("// ");
            if (si < Circled.Length) sb.Append(Circled[si] + " " + sents[si] + " ");
            else sb.Append(sents[si] + " ");
        }
        return sb.ToString().Trim();
    }

    private void LoadMarks(string path)
    {
        try
        {
            var json = File.ReadAllText(path, Encoding.UTF8);
            var data = JsonConvert.DeserializeObject<Dictionary<string, JObject>>(json);
            if (data == null) return;

            foreach (var p in _passages)
            {
                if (!data.TryGetValue(p.Key, out var obj)) continue;
                var marks = obj["marks"]?.ToObject<List<MarkItem>>() ?? new();
                var orderPos = obj["order_positions"]?.ToObject<List<int>>() ?? new();
                p.Numbered = MakeNumbered(p.English, new HashSet<int>(orderPos));
                p.Marks = marks;
            }
            TxtStatus.Text += $" | 마킹 {data.Count}개 로드";
        }
        catch (Exception ex)
        {
            TxtStatus.Text = $"마킹 로드 실패: {ex.Message}";
        }
    }

    // ═══════════════════════════════
    //  지문 표시
    // ═══════════════════════════════

    private void ShowPassage(int index)
    {
        if (index < 0 || index >= _passages.Count) return;
        _currentIndex = index;
        var p = _passages[index];

        TxtPageInfo.Text = $"{index + 1:D3} / {_passages.Count}  |  수특 {p.Key}  |  {GetPrimaryType(p)}";
        RenderPassage(p);

        // 우측 해석 (문장 번호 추가)
        if (!string.IsNullOrEmpty(p.Korean))
        {
            var koSents = Regex.Split(p.Korean, @"(?<=[.다])\s+");
            var sb = new StringBuilder();
            for (int si = 0; si < koSents.Length; si++)
            {
                if (si < Circled.Length) sb.Append(Circled[si] + " ");
                sb.AppendLine(koSents[si].Trim());
            }
            TxtKorean.Text = sb.ToString().Trim();
        }
        else
        {
            TxtKorean.Text = "(해석 없음)";
        }
        TxtVocab.Text = string.Join("\n", p.VocabWords.Select(v => "· " + v));
    }

    private void RenderPassage(PassageData p)
    {
        var doc = new FlowDocument { FontFamily = new FontFamily("맑은 고딕"), FontSize = 14, LineHeight = 26 };
        var para = new Paragraph();

        string text = p.Numbered;
        if (string.IsNullOrEmpty(text)) { RtbPassage.Document = doc; return; }

        var charType = new string?[text.Length];
        var inCore = new bool[text.Length];

        foreach (var m in p.Marks.Where(x => x.Type == "핵심"))
            for (int i = m.Start; i < Math.Min(m.End, text.Length); i++)
            { charType[i] = "핵심"; inCore[i] = true; }

        var priority = new Dictionary<string, int> { {"어법",6}, {"빈칸",5}, {"어휘",4}, {"연결어",3}, {"순서",2}, {"밑줄의미",2}, {"서술형",2}, {"핵심",1} };
        foreach (var m in p.Marks.Where(x => x.Type != "핵심"))
            for (int i = m.Start; i < Math.Min(m.End, text.Length); i++)
            {
                var cur = charType[i];
                if (cur == null || priority.GetValueOrDefault(m.Type) > priority.GetValueOrDefault(cur))
                    charType[i] = m.Type;
            }

        int pos = 0;
        while (pos < text.Length)
        {
            var curType = charType[pos];
            var curCore = inCore[pos];
            int start = pos;
            while (pos < text.Length && charType[pos] == curType && inCore[pos] == curCore)
                pos++;

            var segment = text[start..pos];
            var run = new Run(segment);

            if (curType == "어법")
            {
                run.Foreground = new SolidColorBrush(Color.FromRgb(0xE8, 0x38, 0x4F));
                if (curCore) { run.FontWeight = FontWeights.Bold; run.TextDecorations = TextDecorations.Underline; }
            }
            else if (curType == "빈칸")
            {
                run.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xF9, 0xC4));
                run.Foreground = new SolidColorBrush(Color.FromRgb(0xF5, 0x7F, 0x17));
                run.FontWeight = FontWeights.Bold;
            }
            else if (curType == "어휘")
            {
                run.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xDC, 0xA8));
            }
            else if (curType == "연결어")
            {
                run.Background = new SolidColorBrush(Color.FromRgb(0xE0, 0xF2, 0xE0));
                run.Foreground = new SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0x32));
                run.FontWeight = FontWeights.Bold;
            }
            else if (curType == "순서")
            {
                run.Foreground = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3));
                run.FontWeight = FontWeights.Bold;
            }
            else if (curType == "밑줄의미")
            {
                run.TextDecorations = TextDecorations.Underline;
                run.Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0xBC, 0xD4));
            }
            else if (curType == "서술형")
            {
                run.Background = new SolidColorBrush(Color.FromRgb(0xEF, 0xEB, 0xE9));
                run.FontWeight = FontWeights.Bold;
            }
            else if (curType == "핵심" || curCore)
            {
                run.FontWeight = FontWeights.Bold;
                run.TextDecorations = TextDecorations.Underline;
            }

            para.Inlines.Add(run);
        }

        if (!string.IsNullOrEmpty(p.Footnotes))
        {
            para.Inlines.Add(new LineBreak());
            para.Inlines.Add(new LineBreak());
            para.Inlines.Add(new Run(p.Footnotes) { FontSize = 11, Foreground = Brushes.Gray });
        }

        doc.Blocks.Add(para);
        RtbPassage.Document = doc;
    }

    // ═══════════════════════════════
    //  마킹 조작
    // ═══════════════════════════════

    private void BtnMark_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        var markType = btn.Tag?.ToString() ?? "";
        var p = _passages[_currentIndex];
        var sel = RtbPassage.Selection;

        if (sel.IsEmpty)
        {
            TxtSelection.Text = "텍스트를 드래그해서 선택하세요!";
            return;
        }

        string selectedText = sel.Text.Trim();
        if (string.IsNullOrEmpty(selectedText)) return;

        var beforeRange = new TextRange(RtbPassage.Document.ContentStart, sel.Start);
        int offset = beforeRange.Text.Replace("\r\n", "\n").Replace("\r", "").Length;

        int numberedStart = p.Numbered.IndexOf(selectedText, Math.Max(0, offset - 5));
        if (numberedStart < 0) numberedStart = p.Numbered.IndexOf(selectedText);
        if (numberedStart < 0)
        {
            TxtSelection.Text = $"\"{selectedText}\" 위치를 찾을 수 없습니다.";
            return;
        }
        int numberedEnd = numberedStart + selectedText.Length;

        if (markType == "제거")
        {
            p.Marks.RemoveAll(m => m.Start >= numberedStart && m.End <= numberedEnd);
            p.Marks.RemoveAll(m => m.Start == numberedStart && m.End == numberedEnd);
            TxtSelection.Text = $"\"{selectedText}\" 마킹 제거됨";
        }
        else
        {
            p.Marks.RemoveAll(m => m.Start == numberedStart && m.End == numberedEnd && m.Type == markType);
            p.Marks.Add(new MarkItem { Type = markType, Start = numberedStart, End = numberedEnd, Text = selectedText });
            TxtSelection.Text = $"\"{selectedText}\" → {markType}";
        }

        RenderPassage(p);
    }

    private void RtbPassage_SelectionChanged(object sender, RoutedEventArgs e)
    {
        var sel = RtbPassage.Selection;
        if (!sel.IsEmpty)
            TxtSelection.Text = $"선택: \"{sel.Text.Trim()}\"";
    }

    // ═══════════════════════════════
    //  네비게이션 + 저장
    // ═══════════════════════════════

    private void BtnPrev_Click(object sender, RoutedEventArgs e) => ShowPassage(_currentIndex - 1);
    private void BtnNext_Click(object sender, RoutedEventArgs e) => ShowPassage(_currentIndex + 1);

    private async void BtnAiAnalyze_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(AiAnalyzer.ClaudeKey) && string.IsNullOrEmpty(AiAnalyzer.GptKey) && string.IsNullOrEmpty(AiAnalyzer.GeminiKey))
        {
            MessageBox.Show("API 키가 설정되지 않았습니다.\n설정 버튼에서 API 키를 등록하세요.", "AI 분석", MessageBoxButton.OK, MessageBoxImage.Warning);
            BtnSettings_Click(sender, e);
            return;
        }

        var p = _passages[_currentIndex];
        var vocabStr = string.Join(", ", p.VocabWords.Select(v => v.Split(' ')[0]).Take(10));

        // 오버레이 표시
        PnlAiOverlay.Visibility = Visibility.Visible;
        PrgAi.Value = 0;
        TxtAiPercent.Text = "0%";
        TxtAiStatus.Text = $"수특 {p.Key} 분석 준비 중...";

        var analyzer = new AiAnalyzer();
        analyzer.OnProgress += (pct, msg) => Dispatcher.Invoke(() =>
        {
            PrgAi.Value = pct;
            TxtAiPercent.Text = $"{pct}%";
            TxtAiStatus.Text = msg;
            PrgAiIndeterminate.Visibility = pct < 100 ? Visibility.Visible : Visibility.Collapsed;
        });

        try
        {
            p.Marks.Clear();
            var marks = await analyzer.AnalyzeAsync(p.Numbered, p.English, vocabStr, p.Footnotes);
            p.Marks = marks;
            RenderPassage(p);
            TxtStatus.Text = $"AI 분석 완료: {marks.Count}개 마킹";
        }
        catch (Exception ex)
        {
            TxtStatus.Text = $"AI 분석 오류: {ex.Message}";
        }
        finally
        {
            await Task.Delay(500);
            PnlAiOverlay.Visibility = Visibility.Collapsed;
        }
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        var path = Path.Combine(_dataFolder, "marks_v2.json");
        var data = new Dictionary<string, object>();
        foreach (var p in _passages)
        {
            if (p.Marks.Count == 0) continue;
            var orderPositions = new List<int>();
            // // 마크에서 order_positions 역추론
            var sents = SplitRx.Split(p.English);
            foreach (var m in p.Marks.Where(x => x.Type == "순서"))
            {
                // numbered에서 // 위치로부터 문장 번호 계산
                var before = p.Numbered[..m.Start];
                int sentNum = 1;
                for (int ci = 0; ci < Circled.Length && ci < sents.Length; ci++)
                {
                    if (before.Contains(Circled[ci])) sentNum = ci + 2;
                }
                if (!orderPositions.Contains(sentNum)) orderPositions.Add(sentNum);
            }
            orderPositions.Sort();
            data[p.Key] = new { marks = p.Marks, order_positions = orderPositions };
        }
        File.WriteAllText(path, JsonConvert.SerializeObject(data, Formatting.Indented), Encoding.UTF8);
        TxtStatus.Text = $"저장 완료: {path} ({data.Count}개 지문)";
    }

    private void BtnExportDocx_Click(object sender, RoutedEventArgs e)
    {
        BtnSave_Click(sender, e);
        var dlg = new SaveFileDialog
        {
            Filter = "Word 문서|*.docx|모든 파일|*.*",
            FileName = "수특영어_종합.docx",
            Title = "DOCX 내보내기 (한글에서 HWP 변환 가능)"
        };
        if (dlg.ShowDialog() == true)
        {
            try
            {
                ExportService.ExportDocx(_passages, dlg.FileName);
                TxtStatus.Text = $"DOCX 저장 완료: {dlg.FileName}";
                MessageBox.Show($"저장 완료!\n{dlg.FileName}\n\n한글에서 열어서 HWP로 저장 가능합니다.", "내보내기");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"저장 실패: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void BtnSettings_Click(object sender, RoutedEventArgs e)
    {
        var win = new SettingsWindow { Owner = this };
        win.ShowDialog();
    }

    private string GetPrimaryType(PassageData p)
    {
        if (p.TypeHint.Contains("어법")) return "어법";
        int ch = p.Chapter;
        return ch switch
        {
            >= 11 and <= 16 => "빈칸",
            >= 17 and <= 18 => "문장삽입",
            >= 21 and <= 24 => "순서배열",
            >= 25 and <= 28 => "장문독해",
            >= 29 and <= 30 => "복합문",
            >= 1 and <= 2 => "실용문",
            >= 3 and <= 6 => "주제/요지",
            >= 7 and <= 10 => "세부정보",
            _ => "독해"
        };
    }
}
