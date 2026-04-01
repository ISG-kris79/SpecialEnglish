using System.IO;
using System.Text.RegularExpressions;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Font;
using iText.IO.Font;
using iText.Kernel.Colors;

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
    //  PDF 내보내기
    // ═══════════════════════════

    public static void ExportPdf(List<PassageData> passages, string outputPath)
    {
        var writer = new PdfWriter(outputPath);
        var pdf = new PdfDocument(writer);
        var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);
        doc.SetMargins(30, 30, 30, 30);

        PdfFont font;
        try { font = PdfFontFactory.CreateFont(@"C:\Windows\Fonts\malgun.ttf", PdfEncodings.IDENTITY_H); }
        catch { font = PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA); }

        PdfFont fontBold;
        try { fontBold = PdfFontFactory.CreateFont(@"C:\Windows\Fonts\malgunbd.ttf", PdfEncodings.IDENTITY_H); }
        catch { fontBold = font; }

        foreach (var p in passages)
        {
            if (string.IsNullOrEmpty(p.English)) continue;

            // 헤더
            doc.Add(new Paragraph($"수특 {p.Key}  |  {GetType(p)}")
                .SetFont(fontBold).SetFontSize(14)
                .SetFontColor(new DeviceRgb(192, 100, 92)));

            // 영어 지문 (마킹)
            var engPara = new Paragraph().SetFont(font).SetFontSize(10).SetFixedLeading(18);
            RenderPdfMarked(engPara, p.Numbered, p.Marks, font, fontBold);
            doc.Add(engPara);

            // 각주
            if (!string.IsNullOrEmpty(p.Footnotes))
                doc.Add(new Paragraph(p.Footnotes).SetFont(font).SetFontSize(8)
                    .SetFontColor(new DeviceRgb(136, 136, 136)));

            // 구분선
            doc.Add(new Paragraph("─────────────────────────────").SetFont(font).SetFontSize(6)
                .SetFontColor(new DeviceRgb(221, 221, 221)));

            // 해석
            if (!string.IsNullOrEmpty(p.Korean))
            {
                doc.Add(new Paragraph("해석 및 어휘").SetFont(fontBold).SetFontSize(10)
                    .SetFontColor(new DeviceRgb(192, 100, 92)));

                var koSents = Regex.Split(p.Korean, @"(?<=[.다])\s+");
                var koPara = new Paragraph().SetFont(font).SetFontSize(9).SetFixedLeading(16);
                for (int si = 0; si < koSents.Length; si++)
                {
                    var prefix = si < Circled.Length ? Circled[si] + " " : "";
                    koPara.Add(new Text(prefix + koSents[si].Trim() + " "));
                }
                doc.Add(koPara);
            }

            // 어휘
            if (p.VocabWords.Count > 0)
            {
                var vPara = new Paragraph().SetFont(font).SetFontSize(8).SetFixedLeading(14);
                foreach (var vw in p.VocabWords)
                    vPara.Add(new Text("· " + vw + "  "));
                doc.Add(vPara);
            }

            doc.Add(new AreaBreak(iText.Layout.Properties.AreaBreakType.NEXT_PAGE));
        }

        doc.Close();
    }

    private static void RenderPdfMarked(Paragraph para, string text, List<MarkItem> marks, PdfFont font, PdfFont fontBold)
    {
        if (string.IsNullOrEmpty(text) || marks.Count == 0)
        {
            para.Add(new Text(text ?? "").SetFont(font));
            return;
        }

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

        int pos = 0;
        while (pos < text.Length)
        {
            var curType = charType[pos];
            var curCore = inCore[pos];
            int start = pos;
            while (pos < text.Length && charType[pos] == curType && inCore[pos] == curCore)
                pos++;

            var segment = text[start..pos];
            var t = new Text(segment);

            if (curType == "어법")
            {
                t.SetFont(curCore ? fontBold : font).SetFontColor(new DeviceRgb(0xE8, 0x38, 0x4F));
                if (curCore) t.SetUnderline();
            }
            else if (curType == "빈칸")
            {
                t.SetFont(fontBold).SetFontColor(new DeviceRgb(0xF5, 0x7F, 0x17))
                    .SetBackgroundColor(new DeviceRgb(0xFF, 0xF9, 0xC4));
            }
            else if (curType == "어휘")
            {
                t.SetFont(font).SetBackgroundColor(new DeviceRgb(0xFF, 0xDC, 0xA8));
            }
            else if (curType == "연결어")
            {
                t.SetFont(fontBold).SetFontColor(new DeviceRgb(0x2E, 0x7D, 0x32))
                    .SetBackgroundColor(new DeviceRgb(0xE0, 0xF2, 0xE0));
            }
            else if (curType == "순서")
            {
                t.SetFont(fontBold).SetFontColor(new DeviceRgb(0x21, 0x96, 0xF3));
            }
            else if (curCore)
            {
                t.SetFont(fontBold).SetUnderline();
            }
            else
            {
                t.SetFont(font);
            }

            para.Add(t);
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
