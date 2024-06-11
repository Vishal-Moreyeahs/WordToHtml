using DocumentFormat.OpenXml.Wordprocessing;

namespace WordToHtmlConsole
{
    public class ProcessingModel
    {
        public List<ImageModel> Images { get; set; }
        public int NoOfImages { get; set; } = 0;
        public int NoOfHeaderElement { get; set; } = 0;
    }

    public class ImageModel
    { 
        public string ImageType { get; set; } // inline or anchor
        public byte[] ImageData { get; set; }
        public Drawing Drawing { get; set; }
    }
}
