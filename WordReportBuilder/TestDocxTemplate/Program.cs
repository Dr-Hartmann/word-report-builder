using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using System.Reflection;

namespace WordReportBuilder;

public class Program
{
    private static Dictionary<string, string> values = new() {
        { "{{institute}}", "Институт информационных технологий" },
        { "{{faculty}}", "МПО ЭВМ" },
        { "{{type}}", "КУРСОВАЯ РАБОТА" },
        { "{{class}}", "С#-программирование" },
        { "{{plot}}", "Blazor, Rest API, Entity Framework, PostrgreSQL, сериализация, работа с файлами" },
        { "{{group}}", "1ПИб-02-1оп-22" },
        { "{{code}}", "09.03.04" },
        { "{{specialization}}", "Программная инженерия" },
        { "{{student's_fullname}}", "Микуцких Григорий Андреевич" },
        { "{{teacher's_fullname}}", "Шаханов Н.И." },
        { "{{teacher_position}}", "доцент" },
        { "{{year}}", "2025" },
        { "{{annotation}}", "Курсовая работа посвящена (SOME TEXT) разработке...\r\n" +
            "В ходе работы было ...\r\nВ работе присутствует введение в предметную область, …, " +
            "сопровождение графическим материалом и диаграммами, код итоговой программы " +
            "и результаты её тестирования." },
        { "{{введение}}", "YOLOOOOOOOOOOOOOOOOO" }
    };

    public static async Task Main(string[] args)
    {
        string? path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        string inputPath = $@"{path}/template.docx";

        var stopWatch = new Stopwatch();
        var count = 5;

        for (var y = 0; y < count; ++y)
        {
            for (var i = 0; i < count / 2; ++i)
            {
                string outputPath = $@"{path}/template{i}.docx";
                File.Copy(inputPath, outputPath, true);
                stopWatch.Restart();
                var result = await MakeDocument(outputPath, WordprocessingDocument.Open(outputPath, true));
                stopWatch.Stop();
                Console.Write($"{i}): {stopWatch.ElapsedMilliseconds}\t");
                if (result != string.Empty)
                    Console.WriteLine(result);
            }

            string
                outputPath1 = $@"{path}/template10.docx",
                outputPath2 = $@"{path}/template11.docx",
                outputPath3 = $@"{path}/template12.docx";

            File.Copy(inputPath, outputPath1, true);
            File.Copy(inputPath, outputPath2, true);
            File.Copy(inputPath, outputPath3, true);

            stopWatch.Restart();
            var t10 = MakeDocument(outputPath1, WordprocessingDocument.Open(outputPath1, true));
            var t11 = MakeDocument(outputPath2, WordprocessingDocument.Open(outputPath2, true));
            var t12 = MakeDocument(outputPath3, WordprocessingDocument.Open(outputPath3, true));
            await Task.WhenAll(t10, t11, t12);
            stopWatch.Stop();

            Console.WriteLine($"t10+11+12: {stopWatch.ElapsedMilliseconds}");

            if (t10.Result != string.Empty
                || t11.Result != string.Empty
                || t12.Result != string.Empty)
                Console.WriteLine($"{t10.Result}\n\n{t11.Result}\n\n{t12.Result}");
        }
    }

    private static Task<string> MakeDocument(string outputPath, WordprocessingDocument doc)
    {
        try
        {
            var body = doc.MainDocumentPart?.Document.Body;
            if (body is null)
                return Task.FromResult("Тело пусто");

            foreach (var text in body.Descendants<Text>())
            {
                var pair = values
                    .FirstOrDefault(e => text.InnerText.Contains(e.Key));
                if (pair.Key is null || pair.Value is null) continue;

                if (!pair.Value.Contains("\r\n"))
                {
                    text.Text = text.Text.Replace(pair.Key, pair.Value);
                    continue;
                }

                if (text.Parent is not Run run
                    || run.Parent is not Paragraph paragraph) continue;

                var pPr = paragraph.ParagraphProperties;
                if (pPr is null) continue;
                var rPr = run.RunProperties;
                if (rPr is null) continue;

                var appendList = pair.Value
                    .Split("\r\n", StringSplitOptions.None)
                    .Select(newText =>
                    {
                        var newPar = new Paragraph();
                        newPar.ParagraphProperties = (ParagraphProperties)pPr.CloneNode(true);
                        var newRun = new Run();
                        newRun.RunProperties = (RunProperties)rPr.CloneNode(true);
                        newRun.Append(new Text(newText));
                        newPar.Append(newRun);
                        return newPar;
                    });

                var index = body.ToList().IndexOf(paragraph);
                paragraph.Remove();
                appendList.Select(e => body.InsertAt(e, index++));
            }

            doc.Save();
        }
        catch (Exception ex)
        {
            return Task.FromResult(ex.Message);
        }
        finally
        {
            doc.Dispose();
        }

        return Task.FromResult("");
    }
}
