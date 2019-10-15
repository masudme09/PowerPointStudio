using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;

namespace PowerPointStudio
{
    public class ezShape
    {
        public string id { get; set; }
        public string @class { get; set; }
        public ezCss css { get; set; }
        public ezText text = null;
        public string audioUrl { get; set; }
        public ezImage image;
        internal static int shapeCount=0;
        internal static int mediaCount = 0;//to track number of medias detected

        [JsonConstructor]
        public ezShape()
        {

        }


        public ezShape(Shape shape)
        {
            id = "sh" + shapeCount;
            @class = "shape-type-" + shape.AutoShapeType + shapeCount;
            css = new ezCss(shape);
            text = new ezText(shape);

            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia && shape.MediaType == PpMediaType.ppMediaTypeSound)
            {
                mediaCount = mediaCount + 1;
                audioUrl = Utility.getExtractedAudioUrl(mediaCount);
                if (audioUrl != null)
                {
                    this.audioUrl = audioUrl;
                }

            }

            image = new ezImage(shape);
            

            shapeCount++;
        }
    }
}