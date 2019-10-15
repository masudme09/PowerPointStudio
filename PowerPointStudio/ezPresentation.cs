using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace PowerPointStudio
{
    public class ezPresentation
    {
        public List<ezSlide> ezSlides = new List<ezSlide>();

        [JsonConstructor]
        public ezPresentation()
        {

        }

        public ezPresentation(Presentation presentation)
        {
            foreach(Slide sld in presentation.Slides)
            {
                ezSlide slide = new ezSlide(sld);
                ezSlides.Add(slide);
            }
        }
    }
}