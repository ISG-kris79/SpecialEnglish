using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SuneungMarker;

public class AiAnalyzer
{
    public static string ClaudeKey { get; set; } = "";
    public static string GptKey { get; set; } = "";
    public static string GeminiKey { get; set; } = "";

    private static readonly Regex SplitRx = new(@"(?<=[.!?])\s+(?=[A-Z""'\(])|(?<=\.[\u201d\u2019'""'])\s+(?=[A-Z])", RegexOptions.Compiled);
    private static readonly string[] Circled = { "①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩","⑪","⑫","⑬","⑭","⑮","⑯","⑰","⑱","⑲","⑳" };

    public event Action<int, string>? OnProgress;

    private static readonly string PROMPT_TEMPLATE = @"You are a Korean CSAT (수능) English marking expert.

[EXAMPLE]
Passage (5 sentences):
① Sharing meanings and experiences is not a simple, automatic process: whenever two or more people observe or participate in the same ongoing event they experience it from different perspectives, with different histories, with different background knowledge.
② These differences are most glaring when the individuals involved include an infant or very young child and an adult, although an adult arranging an event for an infant may find it difficult to keep these differences in mind.
③ When giving a bath, for example, a parent whose experience with water play has been fun, whose knowledge about the source of the water and its disposition provides confidence, and whose body control ensures that the baby is safely held and not in danger, expects the infant to enjoy splashing water.
④ Yet the baby may make it clear (perhaps by terrified crying) that her perspective is quite different from that of the adult.
⑤ The two share an experience, but not a meaning.
Vocab: perspective, infant, arrange, confidence | Fn: *glaring **disposition

Answer:
{""vocab"":[{""w"":""automatic"",""s"":1},{""w"":""glaring"",""s"":2},{""w"":""difficult"",""s"":2}],""grammar"":[{""w"":""whenever"",""s"":1},{""w"":""involved"",""s"":2},{""w"":""include"",""s"":2},{""w"":""whose"",""s"":3},{""w"":""expects"",""s"":3},{""w"":""that"",""s"":4},{""w"":""that"",""s"":4}],""connector"":[{""w"":""for example"",""s"":3},{""w"":""Yet"",""s"":4}],""core"":[{""p"":""share an experience, but not a meaning"",""s"":5}],""order"":[2,3,4]}

WHY:
- vocab: theme keywords (automatic=core argument, difficult=key concept) + hard footnote (glaring). NOT simple words (infant/perspective/confidence/disposition).
- grammar: CSAT test points only. whose=1st only (others are parallel). that=skip 1st trivial 'ensures that', mark 2nd+3rd. expects=distant subject agreement. involved=participle adj.
- connector: discourse markers only.
- core: last sentence key phrase, exclude simple subject 'The two'.
- order: // before sentences that start new logical sections.

[ANALYZE THIS]
Passage ({SENT_COUNT} sentences):
{PASSAGE}
Vocab: {VOCAB} | Fn: {FOOTNOTES}

Return ONLY JSON. Use ""w"" for word, ""s"" for sentence number, ""p"" for phrase:
{""vocab"":[],""grammar"":[],""connector"":[],""core"":[],""order"":[]}";

    public async Task<List<MarkItem>> AnalyzeAsync(string numbered, string english, string vocabList, string footnotes)
    {
        var sents = SplitRx.Split(english);

        // 문장 번호 붙인 텍스트
        var sb = new StringBuilder();
        for (int i = 0; i < sents.Length; i++)
        {
            if (i < Circled.Length) sb.AppendLine($"{Circled[i]} {sents[i]}");
            else sb.AppendLine(sents[i]);
        }

        var prompt = PROMPT_TEMPLATE
            .Replace("{SENT_COUNT}", sents.Length.ToString())
            .Replace("{PASSAGE}", sb.ToString().Trim())
            .Replace("{VOCAB}", vocabList)
            .Replace("{FOOTNOTES}", footnotes);

        // Claude → GPT → Gemini 순서로 시도
        OnProgress?.Invoke(10, "Claude AI 분석 중...");
        string? result = await CallClaude(prompt);

        if (string.IsNullOrEmpty(result) || result.Contains("error"))
        {
            OnProgress?.Invoke(30, "GPT AI 분석 중...");
            result = await CallGpt(prompt);
        }

        if (string.IsNullOrEmpty(result) || result.Contains("error"))
        {
            OnProgress?.Invoke(50, "Gemini AI 분석 중...");
            result = await CallGemini(prompt);
        }

        OnProgress?.Invoke(70, "결과 변환 중...");
        var marks = ParseResult(result ?? "{}", numbered, english);

        OnProgress?.Invoke(100, "완료!");
        return marks;
    }

    private List<MarkItem> ParseResult(string json, string numbered, string english)
    {
        var marks = new List<MarkItem>();
        try
        {
            var clean = json.Trim();
            if (clean.StartsWith("```")) clean = clean.Split('\n', 2).Last();
            if (clean.Contains("```")) clean = clean.Split("```")[0];
            clean = clean.Trim();

            var obj = JObject.Parse(clean);
            var sents = SplitRx.Split(english);

            // vocab → 어휘
            ParseArray(obj, "vocab", "어휘", sents, numbered, marks);
            // grammar → 어법
            ParseArray(obj, "grammar", "어법", sents, numbered, marks);
            // connector → 연결어
            ParseArray(obj, "connector", "연결어", sents, numbered, marks);
            // core → 핵심
            ParseCoreArray(obj, "core", sents, numbered, marks);
            // order → 순서 (// 마크는 numbered에서 이미 있으면 스킵)
        }
        catch { }
        return marks;
    }

    private void ParseArray(JObject obj, string key, string markType, string[] sents, string numbered, List<MarkItem> marks)
    {
        var arr = obj[key] as JArray;
        if (arr == null) return;

        foreach (var item in arr)
        {
            var word = item["w"]?.ToString() ?? "";
            var sentNum = item["s"]?.Value<int>() ?? 0;
            if (string.IsNullOrEmpty(word) || sentNum < 1 || sentNum > sents.Length) continue;

            var pos = FindInNumbered(word, sentNum, sents, numbered);
            if (pos != null)
            {
                pos.Type = markType;
                marks.Add(pos);
            }
        }
    }

    private void ParseCoreArray(JObject obj, string key, string[] sents, string numbered, List<MarkItem> marks)
    {
        var arr = obj[key] as JArray;
        if (arr == null) return;

        foreach (var item in arr)
        {
            var phrase = item["p"]?.ToString() ?? "";
            var sentNum = item["s"]?.Value<int>() ?? 0;
            if (string.IsNullOrEmpty(phrase) || sentNum < 1 || sentNum > sents.Length) continue;

            var pos = FindInNumbered(phrase, sentNum, sents, numbered);
            if (pos != null)
            {
                pos.Type = "핵심";
                marks.Add(pos);
            }
        }
    }

    private MarkItem? FindInNumbered(string word, int sentNum, string[] sents, string numbered)
    {
        var targetSent = sents[sentNum - 1];

        // 해당 문장에서 단어 위치 (마지막 출현 우선 - 어휘는 보통 독립적 사용이 뒤에 있음)
        var matches = Regex.Matches(targetSent, @"\b" + Regex.Escape(word) + @"\b", RegexOptions.IgnoreCase);
        int wordIdx;
        if (matches.Count > 1)
        {
            // 여러 개면 마지막 것
            wordIdx = matches[^1].Index;
        }
        else if (matches.Count == 1)
        {
            wordIdx = matches[0].Index;
        }
        else
        {
            wordIdx = targetSent.LastIndexOf(word, StringComparison.OrdinalIgnoreCase);
            if (wordIdx < 0) return null;
        }

        // numbered에서 해당 문장 시작 위치
        var sentChunk = targetSent[..Math.Min(25, targetSent.Length)];
        var sentStart = numbered.IndexOf(sentChunk);
        if (sentStart < 0) return null;

        var absStart = sentStart + wordIdx;
        var absEnd = absStart + word.Length;

        if (absStart < 0 || absEnd > numbered.Length) return null;

        return new MarkItem
        {
            Start = absStart,
            End = absEnd,
            Text = numbered[absStart..absEnd]
        };
    }

    // ═══════════════════════════
    //  API 호출
    // ═══════════════════════════

    private async Task<string> CallClaude(string prompt)
    {
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            http.DefaultRequestHeaders.Add("x-api-key", ClaudeKey);
            http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
            var body = new
            {
                model = "claude-sonnet-4-20250514",
                max_tokens = 1500,
                messages = new[] { new { role = "user", content = prompt } }
            };
            var response = await http.PostAsync("https://api.anthropic.com/v1/messages",
                new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json"));
            var json = await response.Content.ReadAsStringAsync();
            var obj = JObject.Parse(json);
            if (obj["error"] != null) return $"{{\"error\":\"{obj["error"]?["message"]}\"}}";
            return obj["content"]?[0]?["text"]?.ToString() ?? "";
        }
        catch (Exception ex) { return $"{{\"error\":\"{ex.Message}\"}}"; }
    }

    private async Task<string> CallGpt(string prompt)
    {
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            http.DefaultRequestHeaders.Add("Authorization", $"Bearer {GptKey}");
            var body = new
            {
                model = "gpt-4.1-mini",
                messages = new[] { new { role = "user", content = prompt } },
                max_tokens = 1500,
                temperature = 0.1
            };
            var response = await http.PostAsync("https://api.openai.com/v1/chat/completions",
                new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json"));
            var json = await response.Content.ReadAsStringAsync();
            var obj = JObject.Parse(json);
            if (obj["error"] != null) return $"{{\"error\":\"{obj["error"]?["message"]}\"}}";
            return obj["choices"]?[0]?["message"]?["content"]?.ToString() ?? "";
        }
        catch (Exception ex) { return $"{{\"error\":\"{ex.Message}\"}}"; }
    }

    private async Task<string> CallGemini(string prompt)
    {
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GeminiKey}";
            var body = new { contents = new[] { new { parts = new[] { new { text = prompt } } } } };
            var response = await http.PostAsync(url,
                new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json"));
            var json = await response.Content.ReadAsStringAsync();
            var obj = JObject.Parse(json);
            return obj["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString() ?? "";
        }
        catch (Exception ex) { return $"{{\"error\":\"{ex.Message}\"}}"; }
    }
}
