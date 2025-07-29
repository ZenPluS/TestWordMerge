using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;

namespace WordMerge
{
    internal static class WordsMergerHelper
    {
        internal static void CopyStyles(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
        {
            var mainStylePart = mainDoc.MainDocumentPart?.StyleDefinitionsPart ?? mainDoc.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();
            var insertStylePart = insertDoc.MainDocumentPart?.StyleDefinitionsPart;

            if (insertStylePart == null)
                return;

            var insertStyles = insertStylePart.Styles;
            var mainStyles = mainStylePart?.Styles ?? new Styles();
            if (mainStylePart?.Styles == null && mainStylePart != null)
                    mainStylePart.Styles = mainStyles;

            foreach (var style in insertStyles?.Elements<Style>().ToList() ?? new List<Style>())
            {
                var existing = mainStyles.Elements<Style>().FirstOrDefault(s => s.StyleId == style.StyleId);
                existing?.Remove();
                mainStyles.Append(style.CloneNode(true));
            }

            mainStylePart?.Styles.Save();

            if (insertDoc.MainDocumentPart.ThemePart == null)
                return;

            var mainThemePart = mainDoc.MainDocumentPart?.ThemePart;
            if (mainThemePart != null)
                mainDoc.MainDocumentPart.DeletePart(mainThemePart);

            mainDoc.MainDocumentPart?.AddNewPart<ThemePart>();
            mainDoc.MainDocumentPart?.ThemePart?.FeedData(insertDoc.MainDocumentPart.ThemePart.GetStream());
        }

        internal static Dictionary<int, int> CopyNumbering(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
        {
            var insertPart = insertDoc.MainDocumentPart?.NumberingDefinitionsPart;
            if (insertPart == null) return new Dictionary<int, int>();

            var mainPart = mainDoc.MainDocumentPart?.NumberingDefinitionsPart ?? mainDoc.MainDocumentPart?.AddNewPart<NumberingDefinitionsPart>();

            var mainNumbering = mainPart?.Numbering ?? new Numbering();
            var insertNumbering = insertPart.Numbering;

            var maxAbstractNumId = mainNumbering.Elements<AbstractNum>()
                .ToList()
                .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(0).Max();

            var maxNumId = mainNumbering.Elements<NumberingInstance>()
                .Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max();

            var abstractNumMap = new Dictionary<int, int>();
            foreach (var abstractNum in insertNumbering.Elements<AbstractNum>())
            {
                maxAbstractNumId++;
                var oldId = abstractNum.AbstractNumberId?.Value ?? 0;
                var newAbstractNum = (AbstractNum)abstractNum.CloneNode(true);
                newAbstractNum.AbstractNumberId = new Int32Value(maxAbstractNumId);
                mainNumbering.AppendChild(newAbstractNum);
                abstractNumMap[oldId] = maxAbstractNumId;
            }

            var numMap = new Dictionary<int, int>();
            foreach (var num in insertNumbering.Elements<NumberingInstance>())
            {
                maxNumId++;
                var oldId = num.NumberID?.Value ?? 0;
                var newNum = (NumberingInstance)num.CloneNode(true);
                newNum.NumberID = new Int32Value(maxNumId);

                var absNumId = newNum.Descendants<AbstractNumId>().FirstOrDefault();
                if (absNumId != null && abstractNumMap.TryGetValue(absNumId.Val?.Value ?? -1, out var value))
                    absNumId.Val = value;

                mainNumbering.AppendChild(newNum);
                numMap[oldId] = maxNumId;
            }

            mainPart?.Numbering.Save();
            return numMap;
        }

        internal static Dictionary<string, string> CopyImages(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
        {
            var imageMap = new Dictionary<string, string>();
            var insertImageParts = insertDoc.MainDocumentPart?.ImageParts.ToList();
            foreach (var imagePart in insertImageParts ?? new List<ImagePart>())
            {
                var oldRelId = insertDoc.MainDocumentPart?.GetIdOfPart(imagePart);
                var imageStream = imagePart.GetStream();
                var newImagePart = mainDoc.MainDocumentPart?.AddImagePart(imagePart.ContentType);
                newImagePart?.FeedData(imageStream);
                if (newImagePart == null)
                    continue;

                var newRelId = mainDoc.MainDocumentPart?.GetIdOfPart(newImagePart);
                if (oldRelId != null) imageMap[oldRelId] = newRelId;
            }
            return imageMap;
        }

        internal static void UpdateImageReferences(OpenXmlElement element, Dictionary<string, string> imageMap)
        {
            foreach (var drawing in element.Descendants<Drawing>().ToList())
            {
                foreach (var blip in drawing.Descendants<Blip>().ToList())
                {
                    var embed = blip.Embed?.Value;
                    if (embed != null && imageMap.TryGetValue(embed, out var value))
                    {
                        blip.Embed.Value = value;
                    }
                }
            }
        }

        internal static void UpdateNumberingReferences(OpenXmlElement element, Dictionary<int, int> numberingMap)
        {
            foreach (var numPr in element.Descendants<NumberingProperties>())
            {
                var numId = numPr.NumberingId;
                if (numId != null && numId.Val != null && numberingMap.TryGetValue(numId.Val.Value, out var value))
                {
                    numId.Val = value;
                }
            }
        }
    }
}
