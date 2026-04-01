using System.IO;
using System.Text.RegularExpressions;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace SuneungMarker;

public static class ExportService
{
    private static readonly Regex SplitRx = new(@"(?<=[.!?])\s+(?=[A-Z""'\(])|(?<=\.[\u201d\u2019'""'])\s+(?=[A-Z])", RegexOptions.Compiled);
    private static readonly string[] Circled = { "①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩","⑪","⑫","⑬","⑭","⑮","⑯","⑰","⑱","⑲","⑳" };

    public static void ExportDocx(List<PassageData> passages, string outputPath)
    {
        var doc = new XWPFDocument();

        // 지문별 생성

        foreach (var p in passages)
        {
            if (string.IsNullOrEmpty(p.English)) continue;

            // 헤더
            var header = doc.CreateParagraph();
            var hRun = header.CreateRun();
            hRun.SetText($"수특 {p.Key}  |  {GetType(p)}");
            hRun.FontSize = 14;
            hRun.IsBold = true;
            hRun.SetColor("C0645C");
            hRun.FontFamily = "맑은 고딕";

            // 영어 지문 (마킹 적용)
            var engPara = doc.CreateParagraph();
            engPara.SpacingBetween = 1.5;
            RenderMarkedText(engPara, p.Numbered, p.Marks);

            // 각주
            if (!string.IsNullOrEmpty(p.Footnotes))
            {
                var fnPara = doc.CreateParagraph();
                var fnRun = fnPara.CreateRun();
                fnRun.SetText(p.Footnotes);
                fnRun.FontSize = 9;
                fnRun.SetColor("888888");
                fnRun.FontFamily = "맑은 고딕";
            }

            // 구분선
            var divPara = doc.CreateParagraph();
            var divRun = divPara.CreateRun();
            divRun.SetText("─────────────────────────────────");
            divRun.SetColor("DDDDDD");
            divRun.FontSize = 6;

            // 해석
            if (!string.IsNullOrEmpty(p.Korean))
            {
                var koHeader = doc.CreateParagraph();
                var khRun = koHeader.CreateRun();
                khRun.SetText("해석 및 어휘");
                khRun.FontSize = 10;
                khRun.IsBold = true;
                khRun.SetColor("C0645C");
                khRun.FontFamily = "맑은 고딕";

                // 문장 번호 추가
                var koSents = Regex.Split(p.Korean, @"(?<=[.다])\s+");
                var koPara = doc.CreateParagraph();
                koPara.SpacingBetween = 1.3;
                for (int si = 0; si < koSents.Length; si++)
                {
                    var koRun = koPara.CreateRun();
                    var prefix = si < Circled.Length ? Circled[si] + " " : "";
                    koRun.SetText(prefix + koSents[si].Trim() + " ");
                    koRun.FontSize = 9;
                    koRun.FontFamily = "맑은 고딕";
                }
            }

            // 어휘
            if (p.VocabWords.Count > 0)
            {
                var vPara = doc.CreateParagraph();
                foreach (var vw in p.VocabWords)
                {
                    var vRun = vPara.CreateRun();
                    vRun.SetText("· " + vw + "  ");
                    vRun.FontSize = 8;
                    vRun.FontFamily = "맑은 고딕";
                }
            }

            // 페이지 나눔
            var breakPara = doc.CreateParagraph();
            breakPara.IsPageBreak = true;
        }

        using var fs = new FileStream(outputPath, FileMode.Create);
        doc.Write(fs);
    }

    private static void RenderMarkedText(XWPFParagraph para, string text, List<MarkItem> marks)
    {
        if (string.IsNullOrEmpty(text) || marks.Count == 0)
        {
            var run = para.CreateRun();
            run.SetText(text ?? "");
            run.FontSize = 11;
            run.FontFamily = "맑은 고딕";
            return;
        }

        // 문자별 마킹 맵
        var charType = new string?[text.Length];
        var inCore = new bool[text.Length];
        var priority = new Dictionary<string, int> { {"어법",6}, {"빈칸",5}, {"어휘",4}, {"연결어",3}, {"순서",2}, {"밑줄의미",2}, {"서술형",2}, {"핵심",1} };

        foreach (var m in marks.Where(x => x.Type == "핵심"))
            for (int i = m.Start; i < Math.Min(m.End, text.Length); i++)
            { charType[i] = "핵심"; inCore[i] = true; }

        foreach (var m in marks.Where(x => x.Type != "핵심"))
            for (int i = m.Start; i < Math.Min(m.End, text.Length); i++)
            {
                var cur = charType[i];
                if (cur == null || priority.GetValueOrDefault(m.Type) > priority.GetValueOrDefault(cur))
                    charType[i] = m.Type;
            }

        // 세그먼트별 Run 생성
        int pos = 0;
        while (pos < text.Length)
        {
            var curType = charType[pos];
            var curCore = inCore[pos];
            int start = pos;
            while (pos < text.Length && charType[pos] == curType && inCore[pos] == curCore)
                pos++;

            var segment = text[start..pos];
            var run = para.CreateRun();
            run.SetText(segment);
            run.FontSize = 11;
            run.FontFamily = "맑은 고딕";

            if (curType == "어법")
            {
                run.SetColor("E8384F");
                if (curCore) { run.IsBold = true; run.GetCTR().AddNewRPr().u = new CT_Underline { val = ST_Underline.single }; }
            }
            else if (curType == "빈칸")
            {
                run.SetColor("E8384F");
                run.IsBold = true;
                // NPOI doesn't have highlight easily, use bold+color
            }
            else if (curType == "어휘")
            {
                // 주황 형광
                run.GetCTR().AddNewRPr().highlight = new CT_Highlight { val = ST_HighlightColor.yellow };
            }
            else if (curType == "연결어")
            {
                run.SetColor("2E7D32");
                run.IsBold = true;
            }
            else if (curType == "순서")
            {
                run.SetColor("2196F3");
                run.IsBold = true;
            }
            else if (curType == "밑줄의미")
            {
                run.SetColor("00BCD4");
                run.GetCTR().AddNewRPr().u = new CT_Underline { val = ST_Underline.single };
            }
            else if (curType == "서술형")
            {
                run.IsBold = true;
            }
            else if (curCore)
            {
                run.IsBold = true;
                run.GetCTR().AddNewRPr().u = new CT_Underline { val = ST_Underline.single };
            }
        }
    }

    // ═══════════════════════════
    //  PDF 내보내기 (QuestPDF)
    // ═══════════════════════════

    public static void ExportPdf(List<PassageData> passages, string outputPath)
    {
        QuestPDF.Settings.License = LicenseType.Community;
        var coralColor = Color.FromHex("#C0645C");

        QuestPDF.Fluent.Document.Create(container =>
        {
            for (int pi = 0; pi < passages.Count; pi++)
            {
                var p = passages[pi];
                if (string.IsNullOrEmpty(p.English)) continue;

                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(10).FontFamily("맑은 고딕"));

                    // 상단 배너
                    page.Header().Background(coralColor).Padding(10).Row(row =>
                    {
                        row.RelativeItem().Text(text =>
                        {
                            text.Span("수특 영어를 휩쓸다").FontSize(14).Bold().FontColor(Colors.White);
                        });
                        row.ConstantItem(200).AlignRight().Text(text =>
                        {
                            text.Span($"수특 {p.Key}  |  {GetType(p)}").FontSize(10).FontColor(Colors.White);
                        });
                    });

                    page.Content().Column(col =>
                    {
                        col.Spacing(6);
                        col.Item().Height(5);

                        // 영어 지문 (마킹)
                        col.Item().Text(text => RenderPdfMarked(text, p.Numbered, p.Marks));

                        // 각주
                        if (!string.IsNullOrEmpty(p.Footnotes))
                            col.Item().Text(p.Footnotes).FontSize(8).FontColor(Colors.Grey.Medium);

                        // 구분선
                        col.Item().PaddingVertical(4).LineHorizontal(0.5f).LineColor(Colors.Grey.Lighten3);

                        // 해석 및 어휘 헤더
                        col.Item().Text("해석 및 어휘 ✈").Bold().FontSize(11).FontColor(coralColor);

                        // 해석 (문장 번호 추가)
                        if (!string.IsNullOrEmpty(p.Korean))
                        {
                            var koSents = Regex.Split(p.Korean, @"(?<=[.다])\s+");
                            col.Item().Text(text =>
                            {
                                for (int si = 0; si < koSents.Length; si++)
                                {
                                    var prefix = si < Circled.Length ? Circled[si] + " " : "";
                                    text.Span(prefix + koSents[si].Trim() + "\n").FontSize(9).LineHeight(1.5f);
                                }
                            });
                        }

                        // 어휘
                        if (p.VocabWords.Count > 0)
                        {
                            col.Item().PaddingTop(4).Text(text =>
                            {
                                foreach (var vw in p.VocabWords)
                                    text.Span("· " + vw + "\n").FontSize(8).LineHeight(1.4f);
                            });
                        }
                    });

                    // 하단 페이지 번호
                    page.Footer().AlignCenter().Text(text =>
                        text.CurrentPageNumber().FontSize(8).FontColor(Colors.Grey.Medium));
                });
            }
        }).GeneratePdf(outputPath);
    }

    private static void RenderPdfMarked(TextDescriptor text, string content, List<MarkItem> marks)
    {
        if (string.IsNullOrEmpty(content) || marks.Count == 0)
        {
            text.Span(content ?? "");
            return;
        }

        var charType = new string?[content.Length];
        var inCore = new bool[content.Length];
        var priority = new Dictionary<string, int> { {"어법",6}, {"빈칸",5}, {"어휘",4}, {"연결어",3}, {"순서",2}, {"핵심",1} };

        foreach (var m in marks.Where(x => x.Type == "핵심"))
            for (int i = m.Start; i < Math.Min(m.End, content.Length); i++)
            { charType[i] = "핵심"; inCore[i] = true; }

        foreach (var m in marks.Where(x => x.Type != "핵심"))
            for (int i = m.Start; i < Math.Min(m.End, content.Length); i++)
            {
                var cur = charType[i];
                if (cur == null || priority.GetValueOrDefault(m.Type) > priority.GetValueOrDefault(cur))
                    charType[i] = m.Type;
            }

        int pos = 0;
        while (pos < content.Length)
        {
            var curType = charType[pos];
            var curCore = inCore[pos];
            int start = pos;
            while (pos < content.Length && charType[pos] == curType && inCore[pos] == curCore)
                pos++;

            var segment = content[start..pos];

            if (curType == "어법")
            {
                var s = text.Span(segment).FontColor(Color.FromHex("#E8384F"));
                if (curCore) s.Bold().Underline();
            }
            else if (curType == "빈칸")
                text.Span(segment).Bold().FontColor(Color.FromHex("#F57F17")).BackgroundColor(Color.FromHex("#FFF9C4"));
            else if (curType == "어휘")
                text.Span(segment).BackgroundColor(Color.FromHex("#FFDCA8"));
            else if (curType == "연결어")
                text.Span(segment).Bold().FontColor(Color.FromHex("#2E7D32")).BackgroundColor(Color.FromHex("#E0F2E0"));
            else if (curType == "순서")
                text.Span(segment).Bold().FontColor(Color.FromHex("#2196F3"));
            else if (curCore)
                text.Span(segment).Bold().Underline();
            else
                text.Span(segment);
        }
    }

    private static string GetType(PassageData p)
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
