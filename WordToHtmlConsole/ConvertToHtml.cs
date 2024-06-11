using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using OpenXmlPowerTools;
using System.Drawing.Imaging;
using System.Xml.Linq;

namespace WordToHtmlConsole
{
    public class ConvertToHtml
    {
        public string ParseDOCX(string filePath, string fileName = "TestFileName", bool isHeaderAvailable = false, bool isFooterAvailable = false, int noOfHeaderPara = 0,int noOfFooterPara = 0)
        {
            try
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true))
                {
                    int imageCounter = 0;
                    var pageTitle = fileName;
                    var part = wDoc.CoreFilePropertiesPart;
                    if (part != null)
                        pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fileName;

                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                        PageTitle = pageTitle,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            ++imageCounter;
                            Console.WriteLine("Image Processing - ", imageCounter);
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png") imageFormat = ImageFormat.Png;
                            else if (extension == "gif") imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp") imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg") imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            if (imageFormat == null)
                                return null;

                            string base64 = null;
                            try
                            {
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    imageInfo.Bitmap.Save(ms, imageFormat);
                                    var ba = ms.ToArray();
                                    base64 = System.Convert.ToBase64String(ba);
                                }
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            { return null; }

                            ImageFormat format = imageInfo.Bitmap.RawFormat;
                            ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders().First(c => c.FormatID == format.Guid);
                            string mimeType = codec.MimeType;

                            string imageSource = string.Format("data:{0};base64,{1}", mimeType, base64);

                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageSource),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };

                    XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);


                    //var drawingHtml = ExtractHeaderContent3("C:\\Users\\visha\\Downloads\\WordToHtmlCore\\WordToHtml\\WordToHtml\\Statement-of-Work.docx");

                    var html = new XDocument(new XDocumentType("html", null, null, null), htmlElement);
                    var htmlString = html.ToString(SaveOptions.DisableFormatting);

                    HtmlDocument document = new HtmlDocument();
                    document.LoadHtml(htmlString);

                    HtmlNode lastFooterPara = null;

                    if (isFooterAvailable)
                    {
                        MergeLastParagraphs(document, noOfFooterPara);
                        lastFooterPara = GetLastParagraph(document);
                        //AddHeaderAtFirstParagraph(document, lastPara);
                        AddParagraphAtPageBreaksForFooter(document, lastFooterPara);
                        RemoveLastParagraph(document);
                        RemoveLastParagraphsFromWord(wDoc, noOfFooterPara);
                    }

                    if (isHeaderAvailable)
                    {
                        MergeLastParagraphs(document,noOfHeaderPara);
                        var lastPara = GetLastParagraph(document);
                        AddHeaderAtFirstParagraph(document, lastPara);

                        AddParagraphAtPageBreaks(document, lastPara);

                        RemoveLastParagraphs(document,noOfHeaderPara);

                        RemoveLastParagraphsFromWord(wDoc,noOfHeaderPara);
                    }
                    if (isFooterAvailable && lastFooterPara != null)
                    { 
                        AddFooterAtLastParagraph(document, lastFooterPara);
                    }
                    var str = document.DocumentNode.OuterHtml;

                    return str;
                }
            }
            catch (Exception ex)
            {
                return "File contains corrupt data";
            }
        }

        public void RemoveLastParagraphsFromWord(WordprocessingDocument document, int noOfPara)
        {
            // Get the main document part
            var mainPart = document.MainDocumentPart;

            if (mainPart != null)
            {
                // Get the body element
                var body = mainPart.Document.Body;

                if (body != null)
                {
                    // Get all paragraph elements
                    var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();

                    // Check if there are enough paragraphs to remove
                    if (paragraphs.Count >= noOfPara)
                    {
                        // Remove the last 'noOfPara' paragraphs
                        for (int i = 0; i < noOfPara; i++)
                        {
                            paragraphs[paragraphs.Count - 1 - i].Remove();
                        }

                        // Save the changes to the document
                        mainPart.Document.Save();
                    }
                }
            }
        }

        public void AddHeaderAtFirstParagraph(HtmlDocument document, HtmlNode para)
        {
            // Get the body element
            HtmlNode body = document.DocumentNode.SelectSingleNode("//body");
            if (body != null)
            {
                // Find the first div within the body
                HtmlNode div = body.SelectSingleNode("//div");
                if (div != null)
                {
                    // Find the first paragraph within the div
                    HtmlNode firstParagraph = div.SelectSingleNode("p");
                    if (firstParagraph != null)
                    {
                        // Clone the header node
                        HtmlNode paragraphClone = para.Clone();

                        // Insert the cloned header node before the first paragraph
                        firstParagraph.ParentNode.InsertBefore(paragraphClone, firstParagraph);
                    }
                }
            }
        }

        public void AddFooterAtLastParagraph(HtmlDocument document, HtmlNode para)
        {
            // Get the body element
            HtmlNode body = document.DocumentNode.SelectSingleNode("//body");
            if (body != null)
            {
                // Find the last div within the body
                HtmlNode div = body.SelectSingleNode("//div");
                if (div != null)
                {
                    // Find the last paragraph within the div
                    HtmlNode lastParagraph = div.SelectNodes("p").LastOrDefault();
                    if (lastParagraph != null)
                    {
                        // Clone the footer node
                        HtmlNode paragraphClone = para.Clone();
                        div.AppendChild(paragraphClone);
                    }
                }
            }
        }


        public void RemoveLastParagraphs(HtmlDocument document, int noOfPara)
        {
            // Get the body element
            HtmlNode body = document.DocumentNode.SelectSingleNode("//body");

            if (body != null)
            {
                // Find the div element
                HtmlNode div = body.SelectSingleNode("div");

                if (div != null)
                {
                    // Get all paragraph nodes inside the div
                    HtmlNodeCollection paragraphs = div.SelectNodes("p");

                    if (paragraphs != null && paragraphs.Count >= noOfPara)
                    {
                        // Remove the last 'noOfPara' paragraphs
                        for (int i = 0; i < noOfPara; i++)
                        {
                            HtmlNode lastParagraph = paragraphs[paragraphs.Count - 1 - i];
                            lastParagraph.Remove();
                        }
                    }
                }
            }
        }

        //public void MergeLastParagraphs(HtmlDocument document, int noOfPara)
        //{
        //    // Get the body element
        //    HtmlNode body = document.DocumentNode.SelectSingleNode("//body");

        //    if (body != null)
        //    {
        //        // Find the div element
        //        HtmlNode div = body.SelectSingleNode("div");

        //        if (div != null)
        //        {
        //            // Get all paragraph nodes inside the div
        //            HtmlNodeCollection paragraphs = div.SelectNodes("p");

        //            if (paragraphs != null && paragraphs.Count >= noOfPara)
        //            {
        //                // Collect the content of the last 'noOfPara' paragraphs
        //                string mergedContent = string.Join(" ", paragraphs.Skip(paragraphs.Count - noOfPara).Select(p => p.InnerHtml));

        //                // Remove the last 'noOfPara' paragraphs
        //                for (int i = 0; i < noOfPara; i++)
        //                {
        //                    paragraphs[paragraphs.Count - 1 - i].Remove();
        //                }

        //                // Create a new paragraph with the merged content
        //                HtmlNode newParagraph = HtmlNode.CreateNode($"<p>{mergedContent}</p>");

        //                // Add the new paragraph to the div
        //                div.AppendChild(newParagraph);
        //            }
        //        }
        //    }
        //}
        public void MergeLastParagraphs(HtmlDocument document, int noOfPara)
        {
            // Get the body element
            HtmlNode body = document.DocumentNode.SelectSingleNode("//body");

            if (body != null)
            {
                // Find the div element
                HtmlNode div = body.SelectSingleNode("div");

                if (div != null)
                {
                    // Get all paragraph nodes inside the div
                    HtmlNodeCollection paragraphs = div.SelectNodes("p");

                    if (paragraphs != null && paragraphs.Count >= noOfPara)
                    {
                        // Collect the content of the last 'noOfPara' paragraphs
                        string mergedContent = string.Join(" ", paragraphs.Skip(paragraphs.Count - noOfPara).Select(p => p.InnerHtml));

                        // Collect the classes from the last 'noOfPara' paragraphs
                        var classes = new HashSet<string>();
                        foreach (var p in paragraphs.Skip(paragraphs.Count - noOfPara))
                        {
                            string classAttr = p.GetAttributeValue("class", string.Empty);
                            if (!string.IsNullOrEmpty(classAttr))
                            {
                                foreach (var cls in classAttr.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                                {
                                    classes.Add(cls);
                                }
                            }
                        }

                        // Remove the last 'noOfPara' paragraphs
                        for (int i = 0; i < noOfPara; i++)
                        {
                            paragraphs[paragraphs.Count - 1 - i].Remove();
                        }

                        // Create a new paragraph with the merged content and collected classes
                        HtmlNode newParagraph = HtmlNode.CreateNode($"<p>{mergedContent}</p>");
                        if (classes.Count > 0)
                        {
                            newParagraph.SetAttributeValue("class", string.Join(" ", classes));
                        }

                        // Add the new paragraph to the div
                        div.AppendChild(newParagraph);
                    }
                }
            }
        }


        public static void RemoveLastParagraph(HtmlDocument document)
        {
            // Get the body element
            HtmlNode body = document.DocumentNode.SelectSingleNode("//body");

            if (body != null)
            {
                // Find the last paragraph in the body
                HtmlNode lastParagraph = body.SelectNodes("//p").LastOrDefault();

                lastParagraph.Remove();
            }
        }

        public static HtmlNode GetLastParagraph(HtmlDocument document)
        {
            // Get the body element
            HtmlNode body = document.DocumentNode.SelectSingleNode("//body");

            if (body != null)
            {
                // Find the last paragraph in the body
                HtmlNode lastParagraph = body.SelectNodes("//p").LastOrDefault();

                return lastParagraph;
            }

            return null;
        }

        private static HtmlNode GetValidLastParagraph(HtmlNode paragraph)
        {
            while (paragraph != null)
            {
                // Check if the paragraph contains only spaces or empty lines
                if (string.IsNullOrWhiteSpace(paragraph.InnerText))
                {
                    HtmlNode previousNode = paragraph.PreviousSibling;
                    // Remove the current paragraph
                    //paragraph.Remove();
                    // Move to the previous paragraph
                    paragraph = previousNode;
                }
                else
                {
                    // If the paragraph is valid, return it
                    return paragraph;
                }
            }

            // If no valid paragraph is found, return null
            return null;
        }



        public static void AddParagraphAtPageBreaks(HtmlDocument document, HtmlNode para)
        {
            // Get all <hr> elements that represent page breaks
            var pageBreaks = document.DocumentNode.SelectNodes("//br");

            if (pageBreaks != null)
            {
                foreach (var pageBreak in pageBreaks)
                {
                    // Create a new text node with the specified text
                    HtmlNode paragraphClone = para.Clone();
                    HtmlNode hrNode = HtmlNode.CreateNode("<hr>");


                    // Insert the text node before the page break
                    pageBreak.ParentNode.InsertAfter(paragraphClone, pageBreak);
                    pageBreak.ParentNode.InsertAfter(hrNode, pageBreak);
                }
            }
        }

        public static void AddParagraphAtPageBreaksForFooter(HtmlDocument document, HtmlNode para)
        {
            // Get all <hr> elements that represent page breaks
            var pageBreaks = document.DocumentNode.SelectNodes("//br");

            if (pageBreaks != null)
            {
                foreach (var pageBreak in pageBreaks)
                {
                    // Create a new text node with the specified text
                    HtmlNode paragraphClone = para.Clone();
                    HtmlNode hrNode = HtmlNode.CreateNode("<hr>");


                    // Insert the text node before the page break
                    pageBreak.ParentNode.InsertBefore(paragraphClone, pageBreak);
                    pageBreak.ParentNode.InsertBefore(hrNode, pageBreak);
                }
            }
        }

        public void AddSpanToLastParagraph(HtmlDocument document, HtmlNode spanToAdd)
        {
            // Get the body element
            HtmlNode body = document.DocumentNode.SelectSingleNode("//body");

            if (body != null)
            {
                // Find the last paragraph
                HtmlNode div
                    = body.SelectNodes("div").FirstOrDefault();
                // Find the new last paragraph
                HtmlNode lastParagraph = div.SelectNodes("p").LastOrDefault();

                if (lastParagraph != null)
                {
                    // Append the saved span to the last paragraph
                    lastParagraph.AppendChild(spanToAdd);
                }
                else
                {
                    // If no paragraph exists, create a new one and append the span
                    HtmlNode newParagraph = document.CreateElement("p");
                    newParagraph.AppendChild(spanToAdd);
                    body.AppendChild(newParagraph);
                }
            }
        }


        public HtmlNode GetImagePara(HtmlDocument document)
        {
            // Get the body element
            HtmlNode body = document.DocumentNode.SelectSingleNode("//body");

            if (body != null)
            {
                // Find the last paragraph
                HtmlNode div
                    = body.SelectNodes("div").FirstOrDefault();

                HtmlNode lastParagraph = div.SelectNodes("p").LastOrDefault();

                if (lastParagraph != null)
                {
                    // Find the span containing the img tag
                    HtmlNode spanWithImg = lastParagraph.SelectNodes(".//span")
                                                         .FirstOrDefault(span => span.SelectSingleNode(".//img") != null);

                    if (spanWithImg != null)
                    {
                        // Save the span with img tag to a variable
                        HtmlNode savedSpan = spanWithImg.Clone();

                        // Remove the last paragraph
                        //lastParagraph.Remove();

                        // Return the saved span
                        return savedSpan;
                    }
                }
            }

            // Return null if no paragraph or span with img tag is found
            return null;
        }
    }
}
