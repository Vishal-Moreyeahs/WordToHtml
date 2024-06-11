using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System.Xml;
using System.Xml.Linq;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace WordToHtmlConsole
{
    public class ProcessWordHeader
    {
        public ProcessingModel AppendHeaderFromSourseToDestinationFile(string sourceFilePath, string destinationFilePath = "")
        {
            ProcessingModel processing = new ProcessingModel();
            ImageAdderInDocx img = new ImageAdderInDocx();
            // Open the source Word document
            using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(sourceFilePath, true))
            {
                // Access the first header part
                HeaderPart firstHeaderPart = sourceDoc.MainDocumentPart.HeaderParts.FirstOrDefault();

                if (firstHeaderPart != null && firstHeaderPart.Header != null)
                {
                    // Extract the content from the header
                    Header header = firstHeaderPart.Header;
                    var headerContent = header.Elements().ToList(); // Get all elements in the header

                    //Update Model with number of header
                    processing.NoOfHeaderElement = headerContent.Count;

                    // Open the destination Word document
                    //using (WordprocessingDocument destinationDoc = WordprocessingDocument.Open(destinationFilePath, true))
                    //{
                    //    // Access the body of the destination document
                    //    Body destinationBody = destinationDoc.MainDocumentPart.Document.Body;

                        // Iterate through header content and add paragraphs to destination document
                    foreach (var element in headerContent)
                    {
                        // Clone the element to avoid modifying the original header content
                        var clonedElement = element.CloneNode(true);

                        sourceDoc.MainDocumentPart.Document.Body.Append(clonedElement);
                    }

                    var imageModel = GetImagesFromHeaderParts(sourceDoc);
                    processing.Images = imageModel;
                    processing.NoOfImages = imageModel.Count;

                    // Handle images
                    foreach (var draw in imageModel)
                    {
                        var xName = CheckDrawingType(draw.Drawing);
                        if (xName == "Anchor")
                        {
                            var anchorDrawing = img.ConvertAnchorToInline(sourceDoc, draw.Drawing, draw.ImageData);
                            sourceDoc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(anchorDrawing)));
                            sourceDoc.Save();
                        }
                        if (xName == "Inline")
                        {
                            var inlineDrawing = img.GenerateInlineDrawing(sourceDoc, draw.Drawing, draw.ImageData);
                            sourceDoc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(inlineDrawing)));
                            sourceDoc.Save();
                        }
                        if (xName == "None")
                        {
                            continue;
                        }
                    }
                }
            }
            return processing;
        }

        public ProcessingModel AppendFooterFromSourseToDestinationFile(string sourceFilePath, string destinationFilePath = "")
        {
            ProcessingModel processing = new ProcessingModel();
            ImageAdderInDocx img = new ImageAdderInDocx();
            // Open the source Word document
            using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(sourceFilePath, true))
            {
                // Access the first header part
                FooterPart firstFooterPart = sourceDoc.MainDocumentPart.FooterParts.FirstOrDefault();

                if (firstFooterPart != null && firstFooterPart.Footer != null)
                {
                    // Extract the content from the header
                    Footer footer = firstFooterPart.Footer;
                    var footerContent = footer.Elements().ToList(); // Get all elements in the header
                    processing.NoOfHeaderElement = footerContent.Count;
                    //Update Model with number of header


                    // Open the destination Word document
                    //using (WordprocessingDocument destinationDoc = WordprocessingDocument.Open(destinationFilePath, true))
                    //{
                    //    // Access the body of the destination document
                    //    Body destinationBody = destinationDoc.MainDocumentPart.Document.Body;

                    // Iterate through header content and add paragraphs to destination document
                    foreach (var element in footerContent)
                    {
                        int i = 0;
                        // Clone the element to avoid modifying the original header content
                        var clonedElement = element.CloneNode(true);

                        //var isElementEmpty = IsElementEmpty(clonedElement);

                        //if (!isElementEmpty)
                        //{
                        //    i++;
                        sourceDoc.MainDocumentPart.Document.Body.Append(clonedElement);

                        //}
                        
                        List<byte[]> images = new List<byte[]>();
                    }
                    var imageModel = GetImagesFromFooterParts(sourceDoc);
                    processing.Images = imageModel;
                    processing.NoOfImages = imageModel.Count;

                    //Handle images
                    foreach (var draw in imageModel)
                    {
                        var xName = CheckDrawingType(draw.Drawing);
                        if (xName == "Anchor")
                        {
                            var anchorDrawing = img.ConvertAnchorToInline(sourceDoc, draw.Drawing, draw.ImageData);
                            sourceDoc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(anchorDrawing)));
                            sourceDoc.Save();
                        }
                        if (xName == "Inline")
                        {
                            var inlineDrawing = img.GenerateInlineDrawing(sourceDoc, draw.Drawing, draw.ImageData);
                            sourceDoc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(inlineDrawing)));
                            sourceDoc.Save();
                        }
                        if (xName == "None")
                        {
                            continue;
                        }
                    }
                }
            }
            return processing;
        }

        //public ProcessingModel AppendFooterFromSourseToDestinationFile(string sourceFilePath, string destinationFilePath = "")
        //{
        //    ProcessingModel processing = new ProcessingModel();
        //    ImageAdderInDocx img = new ImageAdderInDocx();
        //    // Open the source Word document
        //    using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(sourceFilePath, true))
        //    {
        //        // Access the first header part
        //        FooterPart firstFooterPart = sourceDoc.MainDocumentPart.FooterParts.FirstOrDefault();

        //        if (firstFooterPart != null && firstFooterPart.Footer != null)
        //        {
        //            // Extract the content from the header
        //            Footer footer = firstFooterPart.Footer;
        //            var footerContent = footer.Elements().ToList(); // Get all elements in the header

        //            //Update Model with number of header


        //            // Open the destination Word document
        //            //using (WordprocessingDocument destinationDoc = WordprocessingDocument.Open(destinationFilePath, true))
        //            //{
        //            //    // Access the body of the destination document
        //            //    Body destinationBody = destinationDoc.MainDocumentPart.Document.Body;

        //            // Iterate through header content and add paragraphs to destination document
        //            foreach (var element in footerContent)
        //            {
        //                int i = 0;
        //                // Clone the element to avoid modifying the original header content
        //                var clonedElement = element.CloneNode(true);

        //                var isElementEmpty = IsElementEmpty(clonedElement);

        //                if (!isElementEmpty)
        //                {
        //                    i++;
        //                    sourceDoc.MainDocumentPart.Document.Body.Append(clonedElement);

        //                }
        //                processing.NoOfHeaderElement = i;
        //                List<byte[]> images = new List<byte[]>();

        //                // Add a custom attribute to identify the cloned element
        //                //AddCustomAttribute(clonedElement, "class", "custom-header-element");

        //                // Also handle images embedded within drawing elements
        //                //foreach (var headerPart in sourceDoc.MainDocumentPart.HeaderParts)
        //                //{
        //                //    foreach (var part in headerPart.ImageParts)
        //                //    {
        //                //        using (var stream = part.GetStream())
        //                //        {
        //                //            byte[] imageData = ReadFully(stream); // Read image data into byte array
        //                //            images.Add(imageData);
        //                //        }
        //                //    }
        //                //}
        //                var imageModel = GetImagesFromFooterParts(sourceDoc);
        //                processing.Images = imageModel;
        //                processing.NoOfImages = imageModel.Count;


        //                //Handle images
        //                foreach (var draw in imageModel)
        //                {
        //                    var xName = CheckDrawingType(draw.Drawing);
        //                    if (xName == "Anchor")
        //                    {
        //                        var anchorDrawing = img.ConvertAnchorToInline(sourceDoc, draw.Drawing, draw.ImageData);
        //                        sourceDoc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(anchorDrawing)));
        //                        sourceDoc.Save();
        //                    }
        //                    if (xName == "Inline")
        //                    {
        //                        var inlineDrawing = img.GenerateInlineDrawing(sourceDoc, draw.Drawing, draw.ImageData);
        //                        sourceDoc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(inlineDrawing)));
        //                        sourceDoc.Save();
        //                    }
        //                    if (xName == "None")
        //                    {
        //                        continue;
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    return processing;
        //}

        public bool IsElementEmpty(OpenXmlElement element)
        {
            // Recursively check each child element
            foreach (var child in element.ChildElements)
            {
                if (child is Paragraph paragraph)
                {
                    // Check each run in the paragraph
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        foreach (var runChild in run.ChildElements)
                        {
                            if (runChild is Text text && !string.IsNullOrWhiteSpace(text.Text))
                            {
                                // Found non-whitespace text
                                return false;
                            }
                            else if (runChild is Drawing || runChild is Picture || runChild is Inline)
                            {
                                // Found an image or a logo
                                return false;
                            }
                        }
                    }
                }
                else if (child is Table)
                {
                    // Check table recursively
                    if (!IsElementEmpty(child))
                    {
                        return false;
                    }
                }
                else if (child is Run)
                {
                    // Check each child of Run
                    foreach (var runChild in child.ChildElements)
                    {
                        if (runChild is Text text && !string.IsNullOrWhiteSpace(text.Text))
                        {
                            // Found non-whitespace text
                            return false;
                        }
                        else if (runChild is Drawing || runChild is Picture || runChild is Inline)
                        {
                            // Found an image or a logo
                            return false;
                        }
                    }
                }
            }

            // If no non-whitespace text or images were found, it's empty
            return true;
        }


        private static List<ImageModel> GetImagesFromHeaderParts(WordprocessingDocument sourceDoc)
        {
            List<ImageModel> images = new List<ImageModel>();

            foreach (var headerPart in sourceDoc.MainDocumentPart.HeaderParts)
            {
                XDocument headerXml;
                Stream stream = headerPart.GetStream();
                headerXml = XDocument.Load(XmlReader.Create(stream));
                stream.Close();
                foreach (var part in headerPart.ImageParts)
                {
                    // Get the relationship ID of the ImagePart
                    string relId = headerPart.GetIdOfPart(part);


                    // Check if the image is in an anchor element or an inline element
                    var imageElement = headerXml.Descendants()
                        .FirstOrDefault(e =>
                            (e.Name.LocalName == "anchor" || e.Name.LocalName == "inline") &&
                            e.Descendants().Any(d => d.Name.LocalName == "blip" && d.Attribute("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")?.Value == relId));

                    if (imageElement != null)
                    {
                        string imageType = imageElement.Name.LocalName; // "anchor" or "inline"

                        Drawing drawingElement = headerPart.Header.Descendants<Drawing>()
                    .FirstOrDefault(d => d.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                    .Any(b => b.Embed == relId));

                        using (var streams = part.GetStream())
                        {
                            byte[] imageData = ReadFully(streams); // Read image data into byte array
                            images.Add(new ImageModel { ImageType = imageType, ImageData = imageData, Drawing = drawingElement });
                        }
                    }
                }
            }

            return images;
        }


        private static List<ImageModel> GetImagesFromFooterParts(WordprocessingDocument sourceDoc)
        {
            List<ImageModel> images = new List<ImageModel>();

            foreach (var footerPart in sourceDoc.MainDocumentPart.FooterParts)
            {
                XDocument footerXml;
                Stream stream = footerPart.GetStream();
                footerXml = XDocument.Load(XmlReader.Create(stream));
                stream.Close();
                foreach (var part in footerPart.ImageParts)
                {
                    // Get the relationship ID of the ImagePart
                    string relId = footerPart.GetIdOfPart(part);


                    // Check if the image is in an anchor element or an inline element
                    var imageElement = footerXml.Descendants()
                        .FirstOrDefault(e =>
                            (e.Name.LocalName == "anchor" || e.Name.LocalName == "inline") &&
                            e.Descendants().Any(d => d.Name.LocalName == "blip" && d.Attribute("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")?.Value == relId));

                    if (imageElement != null)
                    {
                        string imageType = imageElement.Name.LocalName; // "anchor" or "inline"

                        Drawing drawingElement = footerPart.Footer.Descendants<Drawing>()
                    .FirstOrDefault(d => d.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                    .Any(b => b.Embed == relId));

                        using (var streams = part.GetStream())
                        {
                            byte[] imageData = ReadFully(streams); // Read image data into byte array
                            images.Add(new ImageModel { ImageType = imageType, ImageData = imageData, Drawing = drawingElement });
                        }
                    }
                }
            }

            return images;
        }


        private void AddCustomAttribute(OpenXmlElement element, string attributeName, string attributeValue)
        {
            // Check if the element already has attributes and add the custom attribute
            if (element.HasAttributes)
            {
                element.SetAttribute(new OpenXmlAttribute(attributeName, null, attributeValue));
            }
            else
            {
                element.AddNamespaceDeclaration(attributeName, attributeValue);
            }
        }

        public string CheckDrawingType(Drawing drawing)
        {
            // Check if the drawing element contains an Anchor element
            if (drawing.Descendants<Anchor>().Any())
            {
                return "Anchor";
            }

            // Check if the drawing element contains an Inline element
            if (drawing.Descendants<Inline>().Any())
            {
                return "Inline";
            }

            // If neither Anchor nor Inline is present
            return "None";
        }

        private bool ContainsPageBreak(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
        {
            return paragraph.Descendants<Break>().Any(b => b.Type == BreakValues.Page);
        }

        // Helper method to read stream into byte array
        private static byte[] ReadFully(Stream input)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }

        public void AddRandomPageBreak(string sourceFilePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(sourceFilePath, true))
            {
                // Get the main document part
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                if (mainPart != null)
                {
                    // Get all paragraphs in the document
                    var paragraphs = mainPart.Document.Body.Elements<Paragraph>().ToList();

                    // Ensure there are enough paragraphs to add page breaks within two of them
                    if (paragraphs.Count >= 2)
                    {
                        Random rnd = new Random();
                        var selectedParagraphs = paragraphs.OrderBy(x => rnd.Next()).Take(2).ToList();

                        foreach (var paragraph in selectedParagraphs)
                        {
                            // Create a page break
                            Run pageBreakRun = new Run(new Break() { Type = BreakValues.Page });

                            // Add the page break within the selected paragraph
                            var runs = paragraph.Elements<Run>().ToList();
                            if (runs.Count > 0)
                            {
                                int insertPosition = rnd.Next(runs.Count);
                                runs[insertPosition].InsertAfterSelf(pageBreakRun);
                            }
                            else
                            {
                                // If the paragraph does not contain any run, add the page break as a new run
                                paragraph.AppendChild(pageBreakRun);
                            }
                        }
                    }
                    // Save the changes to the document
                    mainPart.Document.Save();
                }
            }
        }

        public bool DocumentContainsHeader(string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                // Access the main document part
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                mainPart.DeletePart(mainPart.ThemePart);

                // Check if the main part contains any headers
                if (mainPart.HeaderParts.Any())
                {
                    return true;
                }

                // Additionally, check the section properties for references to headers
                var sections = mainPart.Document.Body.Elements<SectionProperties>();
                foreach (var section in sections)
                {
                    if (section.GetFirstChild<HeaderReference>() != null)
                    {
                        return true;
                    }
                }

                return false;
            }
        }


        public bool DocumentContainsFooter(string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                // Access the main document part
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                // Check if the main part contains any headers
                if (mainPart.FooterParts.Any())
                {
                    return true;
                }

                // Additionally, check the section properties for references to headers
                var sections = mainPart.Document.Body.Elements<SectionProperties>();
                foreach (var section in sections)
                {
                    if (section.GetFirstChild<FooterReference>() != null)
                    {
                        return true;
                    }
                }

                return false;
            }
        }

    }
}
