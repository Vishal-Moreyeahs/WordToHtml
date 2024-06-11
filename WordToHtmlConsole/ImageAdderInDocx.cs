using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordToHtmlConsole
{
    public class ImageAdderInDocx
    {
        public Drawing ConvertAnchorToInline(WordprocessingDocument doc, Drawing anchorDrawing, byte[] imageData)
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;

            // Add the image part and get the relationship ID
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            imagePart.FeedData(new MemoryStream(imageData));
            string relationshipId = mainPart.GetIdOfPart(imagePart);

            // Extract existing anchor properties
            var extent = anchorDrawing.Anchor?.Extent;
            var distanceFromTop = anchorDrawing.Anchor.DistanceFromTop;
            var distanceFromBottom = anchorDrawing.Anchor.DistanceFromBottom;
            var distanceFromLeft = anchorDrawing.Anchor.DistanceFromLeft;
            var distanceFromRight = anchorDrawing.Anchor.DistanceFromRight;
            var editId = anchorDrawing.Anchor.EditId;

            // Create new Inline element with the image
            var inline = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = extent?.Cx ?? 0, Cy = extent?.Cy ?? 0 },
                    new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties(),
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(),
                                new PIC.BlipFill(new A.Blip() { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties()
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                {
                    DistanceFromTop = distanceFromTop,
                    DistanceFromBottom = distanceFromBottom,
                    DistanceFromLeft = distanceFromLeft,
                    DistanceFromRight = distanceFromRight,
                    EditId = editId
                }
            );



            return inline;
        }
        public Drawing GenerateInlineDrawing(WordprocessingDocument doc, Drawing anchorDrawing, byte[] imageData)
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;

            // Add the image part and get the relationship ID
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            imagePart.FeedData(new MemoryStream(imageData));
            string relationshipId = mainPart.GetIdOfPart(imagePart);

            var extent = anchorDrawing.Inline.Extent;
            var distanceFromTop = anchorDrawing.Inline.DistanceFromTop;
            var distanceFromLeft = anchorDrawing.Inline.DistanceFromLeft;
            var distanceFromBottom = anchorDrawing.Inline.DistanceFromBottom;
            var distanceFromRight = anchorDrawing.Inline.DistanceFromRight;
            var editId = anchorDrawing.Inline.EditId;


            var inline = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = extent?.Cx ?? 0, Cy = extent?.Cy ?? 0 },
                    new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties(),
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(),
                                new PIC.BlipFill(new A.Blip() { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties()
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                {
                    DistanceFromTop = distanceFromTop,
                    DistanceFromBottom = distanceFromBottom,
                    DistanceFromLeft = distanceFromLeft,
                    DistanceFromRight = distanceFromRight,
                    EditId = editId
                }
            );

            return inline;
        }


    }


}
