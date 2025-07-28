using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

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
        var mainBytes = Convert.FromBase64String(mainBase64);
        var insertBytes = Convert.FromBase64String(insertBase64);

        using (var mainStream = new MemoryStream())
        using (var insertStream = new MemoryStream(insertBytes))
        {
            mainStream.Write(mainBytes, 0, mainBytes.Length);
            mainStream.Position = 0;

            using (var mainDoc = WordprocessingDocument.Open(mainStream, true))
            using (var insertDoc = WordprocessingDocument.Open(insertStream, false))
            {
                CopyStyles(mainDoc, insertDoc);
                var numberingMap = CopyNumbering(mainDoc, insertDoc);
                var imageMap = CopyImages(mainDoc, insertDoc);

                var mainBody = mainDoc.MainDocumentPart.Document.Body;
                var insertBody = insertDoc.MainDocumentPart.Document.Body;

                var placeholderParagraph = mainBody
                    .Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.Contains(placeholder));

                if (placeholderParagraph == null)
                    throw new InvalidOperationException("Placeholder non trovato nel documento principale.");

                // Solo se il parent è Body
                if (placeholderParagraph.Parent is Body parentBody)
                {
                    var elements = parentBody.Elements().ToList();
                    var index = elements.IndexOf(placeholderParagraph);

                    placeholderParagraph.Remove();

                    foreach (var element in insertBody.Elements())
                    {
                        var imported = element.CloneNode(true);

                        // Aggiorna riferimenti immagini e numbering
                        UpdateImageReferences(imported, imageMap);
                        UpdateNumberingReferences(imported, numberingMap);

                        parentBody.InsertAt(imported, index++);
                    }
                }
                else
                {
                    throw new InvalidOperationException("Il placeholder non si trova nel body principale.");
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

        if (insertStylePart == null)
            return;

        var insertStyles = insertStylePart.Styles;
        var mainStyles = mainStylePart.Styles ?? new Styles();
        if (mainStylePart.Styles == null)
            mainStylePart.Styles = mainStyles;

        foreach (var style in insertStyles.Elements<Style>())
        {
            // Rimuovi lo stile esistente con lo stesso StyleId (così viene sovrascritto)
            var existing = mainStyles.Elements<Style>().FirstOrDefault(s => s.StyleId == style.StyleId);
            if (existing != null)
                existing.Remove();

            mainStyles.Append(style.CloneNode(true));
        }
        mainStylePart.Styles.Save();

        // Copia anche il ThemePart se presente (per i colori)
        if (insertDoc.MainDocumentPart.ThemePart != null)
        {
            var mainThemePart = mainDoc.MainDocumentPart.ThemePart;
            if (mainThemePart != null)
                mainDoc.MainDocumentPart.DeletePart(mainThemePart);

            mainDoc.MainDocumentPart.AddNewPart<ThemePart>();
            mainDoc.MainDocumentPart.ThemePart.FeedData(insertDoc.MainDocumentPart.ThemePart.GetStream());
        }
    }

    // Restituisce una mappa oldNumId -> newNumId
    private static Dictionary<int, int> CopyNumbering(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
    {
        var insertPart = insertDoc.MainDocumentPart.NumberingDefinitionsPart;
        if (insertPart == null) return new Dictionary<int, int>();

        var mainPart = mainDoc.MainDocumentPart.NumberingDefinitionsPart ?? mainDoc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
        if (mainPart.Numbering == null)
            mainPart.Numbering = new Numbering();

        var mainNumbering = mainPart.Numbering;
        var insertNumbering = insertPart.Numbering;

        var maxAbstractNumId = mainNumbering.Elements<AbstractNum>()
            .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(0).Max();
        var maxNumId = mainNumbering.Elements<NumberingInstance>()
            .Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max();

        var abstractNumMap = new Dictionary<int, int>();
        foreach (var abstractNum in insertNumbering.Elements<AbstractNum>())
        {
            maxAbstractNumId++;
            var oldId = abstractNum.AbstractNumberId.Value;
            var newAbstractNum = (AbstractNum)abstractNum.CloneNode(true);
            newAbstractNum.AbstractNumberId = new Int32Value(maxAbstractNumId);
            mainNumbering.AppendChild(newAbstractNum);
            abstractNumMap[oldId] = maxAbstractNumId;
        }

        var numMap = new Dictionary<int, int>();
        foreach (var num in insertNumbering.Elements<NumberingInstance>())
        {
            maxNumId++;
            var oldId = num.NumberID.Value;
            var newNum = (NumberingInstance)num.CloneNode(true);
            newNum.NumberID = new Int32Value(maxNumId);

            // Aggiorna AbstractNumId se necessario
            var absNumId = newNum.Descendants<AbstractNumId>().FirstOrDefault();
            if (absNumId != null && abstractNumMap.ContainsKey(absNumId.Val.Value))
                absNumId.Val = abstractNumMap[absNumId.Val.Value];

            mainNumbering.AppendChild(newNum);
            numMap[oldId] = maxNumId;
        }

        mainPart.Numbering.Save();
        return numMap;
    }

    // Restituisce una mappa oldRelId -> newRelId
    private static Dictionary<string, string> CopyImages(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
    {
        var imageMap = new Dictionary<string, string>();
        var insertImageParts = insertDoc.MainDocumentPart.ImageParts.ToList();
        foreach (var imagePart in insertImageParts)
        {
            var oldRelId = insertDoc.MainDocumentPart.GetIdOfPart(imagePart);
            var imageStream = imagePart.GetStream();
            var newImagePart = mainDoc.MainDocumentPart.AddImagePart(imagePart.ContentType);
            newImagePart.FeedData(imageStream);
            var newRelId = mainDoc.MainDocumentPart.GetIdOfPart(newImagePart);
            imageMap[oldRelId] = newRelId;
        }
        return imageMap;
    }

    // Aggiorna i riferimenti alle immagini nei nodi inseriti
    private static void UpdateImageReferences(OpenXmlElement element, Dictionary<string, string> imageMap)
    {
        foreach (var drawing in element.Descendants<Drawing>())
        {
            foreach (var blip in drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>())
            {
                var embed = blip.Embed?.Value;
                if (embed != null && imageMap.ContainsKey(embed))
                {
                    blip.Embed.Value = imageMap[embed];
                }
            }
        }
    }

    // Aggiorna i riferimenti di numberingId nei paragrafi inseriti
    private static void UpdateNumberingReferences(OpenXmlElement element, Dictionary<int, int> numberingMap)
    {
        foreach (var numPr in element.Descendants<NumberingProperties>())
        {
            var numId = numPr.NumberingId;
            if (numId != null && numberingMap.ContainsKey(numId.Val.Value))
            {
                numId.Val = numberingMap[numId.Val.Value];
            }
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