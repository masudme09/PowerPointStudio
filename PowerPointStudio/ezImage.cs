using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System.IO;

namespace PowerPointStudio
{
   
    public class ezImage
    {
        [JsonProperty(Order = 1)]
        public ezCss css { get; set; }

        [JsonProperty(Order = 2)]
        public string imgurlLarge,imgurlMedium,imgurlSmall;

        internal string actualUrl;

        /// <summary>
        /// Default constructor
        /// </summary>
        [JsonConstructor]
        public ezImage()
        {

        }


        public ezImage(Shape shape)
        {
            //Export the image of the shape
            int slideWidth = (int)Utility.SlideWidthGet((shape.Parent).Parent);
            int slideHeight = (int)Utility.SlideHeightGet((shape.Parent).Parent);

            string shapeExportDirectory = PowerPointStudioRibbon.currentPPTPath + "\\images";
            if(!Directory.Exists(shapeExportDirectory))
            {
                Directory.CreateDirectory(shapeExportDirectory);
            }

            //Need to set rotation property 0 before export then set back to original
            float originalRotation = shape.Rotation;
            shape.Rotation = 0;
            string exportedUrl = shapeExportDirectory + "\\" + Utility.RandomNumber(1000, 10000, ezShape.shapeCount) + ".png";
            shape.Export(exportedUrl, PpShapeFormat.ppShapeFormatPNG, slideWidth * 4, slideHeight * 4, PpExportMode.ppClipRelativeToSlide);
            //Back rotation to original
            shape.Rotation = originalRotation;

            actualUrl = exportedUrl;
            exportedUrl = exportedUrl.Replace("\\", "/");

            imgurlLarge = exportedUrl.Replace(PowerPointStudioRibbon.currentPPTPath.Replace("\\","/"), "https://ezilmdev.org");
            imgurlMedium = exportedUrl.Replace(PowerPointStudioRibbon.currentPPTPath.Replace("\\", "/"), "https://ezilmdev.org");
            imgurlSmall = exportedUrl.Replace(PowerPointStudioRibbon.currentPPTPath.Replace("\\", "/"), "https://ezilmdev.org");

            css = new ezCss(shape);
        }

        /// <summary>
        /// This constructor is only for SLide Background
        /// </summary>
        /// <param name="slide"></param>
        public ezImage(Slide slide)
        {
            slide.Duplicate();
            Slide tempSlide = slide.Parent.Slides[slide.SlideIndex + 1];
            
            while(tempSlide.Shapes.Count>0)
            {
                tempSlide.Shapes[1].Delete();
            }
                        
           
            //Export the image of the shape
            int slideWidth = (int)Utility.SlideWidthGet(slide.Parent);
            int slideHeight = (int)Utility.SlideHeightGet(slide.Parent);

            string shapeExportDirectory = PowerPointStudioRibbon.currentPPTPath + "\\images";
            if (!Directory.Exists(shapeExportDirectory))
            {
                Directory.CreateDirectory(shapeExportDirectory);
            }

            string exportedUrl = shapeExportDirectory + "\\" + Utility.RandomNumber(1000, 10000, ezShape.shapeCount) + ".png";
            tempSlide.Export(exportedUrl, "PNG", slideWidth * 4, slideHeight * 4);

            actualUrl = exportedUrl;
            exportedUrl = exportedUrl.Replace("\\", "/");
            imgurlLarge = exportedUrl.Replace(PowerPointStudioRibbon.currentPPTPath.Replace("\\", "/"), "https://ezilmdev.org");
            imgurlMedium = exportedUrl.Replace(PowerPointStudioRibbon.currentPPTPath.Replace("\\", "/"), "https://ezilmdev.org");
            imgurlSmall = exportedUrl.Replace(PowerPointStudioRibbon.currentPPTPath.Replace("\\", "/"), "https://ezilmdev.org");

            tempSlide.Delete();
        }
    }
}