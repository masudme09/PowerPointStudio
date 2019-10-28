using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;

namespace PowerPointStudio
{
    public class ezBackGround
    {
        public string id { get; set; }
        public ezImage image;
        private static int backgroundCount=0;

        [JsonConstructor]
        public ezBackGround()
        {

        }

        public ezBackGround(Slide sld )
        {
            id = "SlideBackGround";
            ezCss css = new ezCss(576, 420, 0, 0);
            image = new ezImage(sld);
            image.css = css;
            backgroundCount++;
        }
    }
}