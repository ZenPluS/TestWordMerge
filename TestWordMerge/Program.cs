using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;

public class WordMerger
{
    public static void Main()
    {
        var main64 = ReadFileAsBase64(@"C:\Users\zenpl\Desktop\Main.docx");
        var insert64 = ReadFileAsBase64(@"C:\Users\zenpl\Desktop\Insert.docx");
        var merged = MergeDocumentsBase64(main64, insert64, "<<CONTENT>>");
        WriteBase64ToFile(merged, @"C:\Users\zenpl\Desktop\Main.docx");
        Console.WriteLine("Merge completato con successo!");
    }

    public static string MergeDocumentsBase64(string mainBase64, string insertBase64, string placeholder)
    {
        byte[] mainBytes = Convert.FromBase64String(mainBase64);
        byte[] insertBytes = Convert.FromBase64String(insertBase64);

        using (MemoryStream mainStream = new MemoryStream())
        using (MemoryStream insertStream = new MemoryStream(insertBytes))
        {
            mainStream.Write(mainBytes, 0, mainBytes.Length);
            mainStream.Position = 0;

            using (WordprocessingDocument mainDoc = WordprocessingDocument.Open(mainStream, true))
            using (WordprocessingDocument insertDoc = WordprocessingDocument.Open(insertStream, false))
            {
                CopyStyles(mainDoc, insertDoc);
                //CopyNumbering(mainDoc, insertDoc);
                CopyImages(mainDoc, insertDoc);

                var mainBody = mainDoc.MainDocumentPart.Document.Body;
                var insertBody = insertDoc.MainDocumentPart.Document.Body;

                var placeholderParagraph = mainBody.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.Contains(placeholder));

                if (placeholderParagraph != null)
                {
                    var parent = placeholderParagraph.Parent;
                    var elements = parent.Elements().ToList();
                    int index = elements.IndexOf(placeholderParagraph);

                    placeholderParagraph.Remove();

                    foreach (var element in insertBody.Elements())
                    {
                        var imported = element.CloneNode(true);

                        // Forza l'importazione degli stili locali di ogni paragrafo
                        if (imported is Paragraph para)
                        {
                            foreach (var run in para.Elements<Run>())
                            {
                                var runProps = run.RunProperties ?? new RunProperties();
                                if (run.RunProperties == null)
                                    run.RunProperties = runProps;

                                // Copia font se non definito
                                if (!runProps.Elements<RunFonts>().Any())
                                {
                                    runProps.Append(new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" });
                                }

                                // Imposta il colore del testo se non presente
                                if (!runProps.Elements<Color>().Any())
                                {
                                    runProps.Append(new Color { Val = "000000" }); // nero
                                }
                            }

                            var paraProps = para.ParagraphProperties ?? new ParagraphProperties();
                            if (para.ParagraphProperties == null)
                                para.ParagraphProperties = paraProps;

                            // Imposta colore sfondo paragrafo se mancante
                            if (!paraProps.Elements<Shading>().Any())
                            {
                                paraProps.Append(new Shading
                                {
                                    Val = ShadingPatternValues.Clear,
                                    Color = "auto",
                                    Fill = "FFFFFF" // bianco
                                });
                            }
                        }

                        parent.InsertAt(imported, index++);
                    }
                }

                mainDoc.MainDocumentPart.Document.Save();
            }
            return Convert.ToBase64String(mainStream.ToArray());
        }
    }

    private static void CopyStyles(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
    {
        var mainStylePart = mainDoc.MainDocumentPart.StyleDefinitionsPart ?? mainDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        var insertStylePart = insertDoc.MainDocumentPart.StyleDefinitionsPart;

        if (insertStylePart != null)
        {
            var insertStyles = insertStylePart.Styles;
            var mainStyles = mainStylePart.Styles ?? new Styles();
            if (mainStylePart.Styles == null)
                mainStylePart.Styles = mainStyles;

            foreach (var style in insertStyles.Elements<Style>())
            {
                if (!mainStyles.Elements<Style>().Any(s => s.StyleId == style.StyleId))
                {
                    mainStyles.Append(style.CloneNode(true));
                }
            }
            mainStylePart.Styles.Save();
        }
    }

    private static void CopyNumbering(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
    {
        var insertPart = insertDoc.MainDocumentPart.NumberingDefinitionsPart;
        if (insertPart == null) return;

        var mainPart = mainDoc.MainDocumentPart.NumberingDefinitionsPart ?? mainDoc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
        var insertNumbering = insertPart.Numbering;
        var mainNumbering = mainPart.Numbering;// ?? new Numbering();

        foreach (var num in insertNumbering.Elements())
        {
            mainNumbering.Append(num.CloneNode(true));
        }
        mainPart.Numbering = mainNumbering;
        mainPart.Numbering.Save();
    }

    private static void CopyImages(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
    {
        foreach (var imagePart in insertDoc.MainDocumentPart.ImageParts)
        {
            mainDoc.MainDocumentPart.AddPart(imagePart);
        }
    }

    public static string ReadFileAsBase64(string path)
    {
        return Convert.ToBase64String(File.ReadAllBytes(path));
    }

    public static void WriteBase64ToFile(string base64, string path)
    {
        File.WriteAllBytes(path, Convert.FromBase64String(base64));
    }
}
