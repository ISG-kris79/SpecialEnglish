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

        // 페이지 여백 축소
        var sectPr = doc.Document.body.sectPr ?? (doc.Document.body.sectPr = new NPOI.OpenXmlFormats.Wordprocessing.CT_SectPr());
        var pgMar = sectPr.pgMar ?? (sectPr.pgMar = new NPOI.OpenXmlFormats.Wordprocessing.CT_PageMar());
        pgMar.top = 600; pgMar.bottom = 600; pgMar.left = 700; pgMar.right = 700;

        foreach (var p in passages)
        {
            if (string.IsNullOrEmpty(p.English)) continue;

            // 헤더
            var header = doc.CreateParagraph();
            header.SpacingAfter = 100;
            var hRun = header.CreateRun();
            hRun.SetText($"수특 {p.Key}  |  {GetType(p)}");
            hRun.FontSize = 12;
            hRun.IsBold = true;
            hRun.SetColor("C0645C");
            hRun.FontFamily = "맑은 고딕";

            // 좌우 2단 테이블
            var table = doc.CreateTable(1, 2);
            // 전체 폭 = A4 (약 9638 twips) - 좌우 여백(1400) = 8238
            table.Width = 9000;
            table.SetColumnWidth(0, 5000);  // 좌 55%
            table.SetColumnWidth(1, 4000);  // 우 45%

            // 테두리: 좌우 셀 사이 세로선만
            var tblPr = table.GetCTTbl().tblPr;
            var borders = new CT_TblBorders();
            borders.top = new CT_Border { val = ST_Border.none };
            borders.bottom = new CT_Border { val = ST_Border.none };
            borders.left = new CT_Border { val = ST_Border.none };
            borders.right = new CT_Border { val = ST_Border.none };
            borders.insideH = new CT_Border { val = ST_Border.none };
            borders.insideV = new CT_Border { val = ST_Border.single, color = "DDDDDD", sz = 4 };
            tblPr.tblBorders = borders;

            // ── 좌측: 영어 지문 ──
            var leftCell = table.GetRow(0).GetCell(0);
            var leftFirst = leftCell.Paragraphs[0];
            leftFirst.SpacingBetween = 1.4;
            RenderMarkedText(leftFirst, p.Numbered, p.Marks, 9);

            if (!string.IsNullOrEmpty(p.Footnotes))
            {
                var fnPara = leftCell.AddParagraph();
                fnPara.SpacingBefore = 80;
                var fnRun = fnPara.CreateRun();
                fnRun.SetText(p.Footnotes);
                fnRun.FontSize = 7;
                fnRun.SetColor("888888");
                fnRun.FontFamily = "맑은 고딕";
            }

            // ── 우측: 해석 + 어휘 ──
            var rightCell = table.GetRow(0).GetCell(1);
            // 셀 패딩
            // 셀 좌측 패딩

            var koHeader2 = rightCell.Paragraphs[0];
            koHeader2.SpacingAfter = 60;
            var khRun = koHeader2.CreateRun();
            khRun.SetText("해석 및 어휘 ✈");
            khRun.FontSize = 9;
            khRun.IsBold = true;
            khRun.SetColor("C0645C");
            khRun.FontFamily = "맑은 고딕";

            if (!string.IsNullOrEmpty(p.Korean))
            {
                var koSents = Regex.Split(p.Korean, @"(?<=[.다])\s+");
                for (int si = 0; si < koSents.Length; si++)
                {
                    var koPara = rightCell.AddParagraph();
                    koPara.SpacingBetween = 1.2;
                    koPara.SpacingAfter = 20;
                    var koRun = koPara.CreateRun();
                    var prefix = si < Circled.Length ? Circled[si] + " " : "";
                    koRun.SetText(prefix + koSents[si].Trim());
                    koRun.FontSize = 8;
                    koRun.FontFamily = "맑은 고딕";
                }
            }

            // 구분선
            var divPara = rightCell.AddParagraph();
            divPara.SpacingBefore = 40;
            divPara.SpacingAfter = 40;
            var divRun = divPara.CreateRun();
            divRun.SetText("─────────────────");
            divRun.SetColor("DDDDDD");
            divRun.FontSize = 5;

            if (p.VocabWords.Count > 0)
            {
                foreach (var vw in p.VocabWords)
                {
                    var vPara = rightCell.AddParagraph();
                    vPara.SpacingAfter = 10;
                    var vRun = vPara.CreateRun();
                    vRun.SetText("· " + vw);
                    vRun.FontSize = 7;
                    vRun.FontFamily = "맑은 고딕";
                }
            }

            var breakPara = doc.CreateParagraph();
            breakPara.IsPageBreak = true;
        }

        using var fs = new FileStream(outputPath, FileMode.Create);
        doc.Write(fs);
    }

    private static void RenderMarkedText(XWPFParagraph para, string text, List<MarkItem> marks, int fontSize = 11)
    {
        if (string.IsNullOrEmpty(text) || marks.Count == 0)
        {
            var run = para.CreateRun();
            run.SetText(text ?? "");
            run.FontSize = fontSize;
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
            run.FontSize = fontSize;
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
        var coral = Color.FromHex("#C0645C");

        QuestPDF.Fluent.Document.Create(container =>
        {
            for (int pi = 0; pi < passages.Count; pi++)
            {
                var p = passages[pi];
                if (string.IsNullOrEmpty(p.English)) continue;

                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.MarginVertical(20);
                    page.MarginHorizontal(25);
                    page.DefaultTextStyle(x => x.FontSize(10).FontFamily("맑은 고딕"));

                    // 상단 배너
                    page.Header().Height(30).Background(coral).Padding(6).Row(row =>
                    {
                        row.RelativeItem().AlignMiddle().Text("수특 영어를 휩쓸다")
                            .FontSize(13).Bold().FontColor(Colors.White);
                        row.ConstantItem(180).AlignMiddle().AlignRight()
                            .Text($"수특 {p.Key}  |  {GetType(p)}")
                            .FontSize(9).FontColor(Colors.White);
                    });

                    // 좌우 2단
                    page.Content().PaddingTop(10).Row(row =>
                    {
                        // ── 좌측: 영어 지문 (55%) ──
                        row.RelativeItem(55).PaddingRight(10).Column(left =>
                        {
                            left.Spacing(6);

                            left.Item().Text(text =>
                            {
                                text.DefaultTextStyle(x => x.FontSize(10).LineHeight(1.7f));
                                RenderPdfMarked(text, p.Numbered, p.Marks);
                            });

                            if (!string.IsNullOrEmpty(p.Footnotes))
                                left.Item().PaddingTop(8).Text(p.Footnotes)
                                    .FontSize(8).FontColor(Colors.Grey.Medium);
                        });

                        // 세로 구분선
                        row.ConstantItem(1).Background(Colors.Grey.Lighten3);

                        // ── 우측: 해석 + 어휘 (45%) ──
                        row.RelativeItem(45).PaddingLeft(10).Column(right =>
                        {
                            right.Spacing(6);

                            right.Item().Text("해석 및 어휘 ✈").Bold().FontSize(11).FontColor(coral);

                            if (!string.IsNullOrEmpty(p.Korean))
                            {
                                var koSents = Regex.Split(p.Korean, @"(?<=[.다])\s+");
                                right.Item().Text(text =>
                                {
                                    text.DefaultTextStyle(x => x.FontSize(9).LineHeight(1.6f));
                                    for (int si = 0; si < koSents.Length; si++)
                                    {
                                        var prefix = si < Circled.Length ? Circled[si] + " " : "";
                                        text.Span(prefix + koSents[si].Trim() + "\n");
                                    }
                                });
                            }

                            right.Item().PaddingVertical(5).LineHorizontal(0.5f).LineColor(Colors.Grey.Lighten3);

                            if (p.VocabWords.Count > 0)
                            {
                                right.Item().Text(text =>
                                {
                                    text.DefaultTextStyle(x => x.FontSize(8).LineHeight(1.5f));
                                    foreach (var vw in p.VocabWords)
                                        text.Span("· " + vw + "\n");
                                });
                            }
                        });
                    });

                    page.Footer().AlignCenter().Text(text =>
                        text.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium));
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
