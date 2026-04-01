using System.IO;
using System.IO.Compression;
using System.Text;
using OpenMcdf;

namespace SuneungMarker;

/// <summary>
/// HWP 파일에서 텍스트 추출 (Python 없이 C# 순수 구현)
/// HWP = OLE Compound File → FileHeader, BodyText/SectionN, PrvText
/// </summary>
public static class HwpReader
{
    public static string ExtractText(string hwpPath)
    {
        using var cf = new CompoundFile(hwpPath);

        // 1. FileHeader에서 압축 여부 확인
        var headerStream = cf.RootStorage.GetStream("FileHeader");
        var headerData = headerStream.GetData();
        bool isCompressed = (headerData[36] & 1) != 0;

        // 2. BodyText 섹션들에서 텍스트 추출
        var sb = new StringBuilder();
        var bodyStorage = cf.RootStorage.GetStorage("BodyText");

        int sectionIdx = 0;
        while (true)
        {
            try
            {
                var sectionStream = bodyStorage.GetStream($"Section{sectionIdx}");
                var data = sectionStream.GetData();

                if (isCompressed)
                {
                    data = DecompressZlib(data);
                }

                var text = ExtractTextFromSection(data);
                if (!string.IsNullOrEmpty(text))
                    sb.AppendLine(text);

                sectionIdx++;
            }
            catch
            {
                break; // 더 이상 섹션 없음
            }
        }

        // BodyText가 비면 PrvText(미리보기 텍스트) 사용
        if (sb.Length == 0)
        {
            try
            {
                var prvStream = cf.RootStorage.GetStream("PrvText");
                var prvData = prvStream.GetData();
                sb.Append(Encoding.Unicode.GetString(prvData));
            }
            catch { }
        }

        return sb.ToString();
    }

    private static byte[] DecompressZlib(byte[] data)
    {
        // zlib decompress (raw deflate, skip zlib header)
        using var input = new MemoryStream(data);
        using var output = new MemoryStream();
        using var deflate = new DeflateStream(input, CompressionMode.Decompress);
        deflate.CopyTo(output);
        return output.ToArray();
    }

    private static string ExtractTextFromSection(byte[] data)
    {
        var sb = new StringBuilder();
        int pos = 0;

        while (pos + 4 <= data.Length)
        {
            uint header = BitConverter.ToUInt32(data, pos);
            int tagId = (int)(header & 0x3FF);
            int size = (int)((header >> 20) & 0xFFF);

            if (size == 0xFFF)
            {
                if (pos + 8 > data.Length) break;
                size = (int)BitConverter.ToUInt32(data, pos + 4);
                pos += 8;
            }
            else
            {
                pos += 4;
            }

            if (pos + size > data.Length) break;

            // Tag 67 = HWPTAG_PARA_TEXT
            if (tagId == 67)
            {
                var text = ParseParaText(data, pos, size);
                if (!string.IsNullOrWhiteSpace(text))
                    sb.AppendLine(text.Trim());
            }

            pos += size;
        }

        return sb.ToString();
    }

    private static string ParseParaText(byte[] data, int offset, int size)
    {
        var chars = new StringBuilder();
        int i = offset;
        int end = offset + size;

        while (i + 1 < end)
        {
            ushort code = BitConverter.ToUInt16(data, i);
            i += 2;

            if (code < 32)
            {
                if (code == 0 || code == 10 || code == 13)
                {
                    chars.Append('\n');
                }
                else
                {
                    // 확장 컨트롤 문자 - 14바이트 스킵
                    i += 14;
                }
            }
            else
            {
                chars.Append((char)code);
            }
        }

        return chars.ToString();
    }
}
