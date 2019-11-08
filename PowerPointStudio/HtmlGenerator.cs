using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointStudio
{
    public class HtmlGenerator
    {

        /// <summary>
        /// Generate HTML from JSON of the ezPresentation
        /// </summary>
        /// <param name="JSON"></param>
        public HtmlGenerator(string JSON)
        {
            //delete directory if already exists
            if (Directory.Exists(PowerPointStudioRibbon.currentPPTPath + "\\HTML"))
            {
                Directory.Delete(PowerPointStudioRibbon.currentPPTPath + "\\HTML", true);
            }

            //Parse JSON to ezPresentation
            ezPresentation presentation = Newtonsoft.Json.JsonConvert.DeserializeObject<ezPresentation>(JSON);

            //Vatiables Global Level
            int PageWidth = 960;
            int PageHeight = 700;


            int htmlCount = 0;
            //Generate HTML for all slides
            foreach (ezSlide slide in presentation.ezSlides)
            {
                string styleClass = "a {" +
                "  text-decoration: none;" +
                "  display: inline-block;" +
                "  padding: 8px 16px;" +
                "}" +
                "a:hover {" +
                "  background-color: #ddd;" +
                "  color: black;" +
                "}" +
                ".previous {" +
                "  background-color: #d49764;" +
                "  color: black;" +
                "}" +
                ".next {" +
                "  background-color: #4CAF50;" +
                "  color: white;" +
                "}" +
                ".round {" +
                "  border-radius: 50%;" +
                "}";

                string headHtml = "<html lang=\"en\"><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=windows-1252\">" +
            "    <title>Test Application </title>" +
            "</head>" + "<style>" + styleClass + "</style>" +
            "<body style=\"width: " + PageWidth + "; height:" + PageHeight + ";  margin: 8px;\">";

                //headHtml = headHtml + SlideDiv(slide.backGround.image.actualUrl, (float)Convert.ToDouble(slide.backGround.css.width.Replace("px", "")),
                //    (float)Convert.ToDouble(slide.backGround.css.height.Replace("px", "")), (float)Convert.ToDouble(slide.backGround.css.left.Replace("px", "")),
                //    (float)Convert.ToDouble(slide.backGround.css.top.Replace("px", "")));

                foreach (ezShape shape in slide.shapes)
                {
                    if (shape.image.imgurlLarge != null)
                    {
                        string imageActualUrl = ((shape.image.imgurlLarge).Replace("https://ezilmdev.org",Globals.ThisAddIn.Application.ActivePresentation.Path +@"\temp")).Replace(@"/",@"\");

                        headHtml = headHtml + DivAdded(imageActualUrl, (float)Convert.ToDouble(shape.image.css.width.Replace("px", "")),
                        (float)Convert.ToDouble(shape.image.css.height.Replace("px", "")), (float)Convert.ToDouble(shape.image.css.left.Replace("px", "")),
                        (float)Convert.ToDouble(shape.image.css.top.Replace("px", "")), false, false, shape.image.css.rotation);
                    }
                }

                //Adding navigation button
                string navigation = "";

                if (htmlCount == 0 && presentation.ezSlides.Count>1)
                {
                    navigation = "<div style=\"margin: 0px; position: absolute; top: 720px; left: 400px;\">" +
                "            <a href=\"html0.html\" class=\"previous\">&laquo; Previous</a><a href=\"html" + (htmlCount + 1) + ".html\" class=\"next\">Next &raquo;</a>" +
                "        </div>";
                    headHtml = headHtml + navigation;
                }
                else if(htmlCount == 0 && presentation.ezSlides.Count == 1)
                {
                    navigation = "<div style=\"margin: 0px; position: absolute; top: 720px; left: 400px;\">" +
                "            <a href=\"html0.html\" class=\"previous\">&laquo; Previous</a><a href=\"html" + (htmlCount) + ".html\" class=\"next\">Next &raquo;</a>" +
                "        </div>";
                    headHtml = headHtml + navigation;
                }
                else if (htmlCount < presentation.ezSlides.Count - 1)
                {
                    navigation = "<div style=\"margin: 0px; position: absolute; top: 720px; left: 400px;\">" +
                "            <a href=\"html" + (htmlCount - 1) + ".html\" class=\"previous\">&laquo; Previous</a><a href=\"html" + (htmlCount + 1) + ".html\" class=\"next\">Next &raquo;</a>" +
                "        </div>";
                    headHtml = headHtml + navigation;
                }
                else if (htmlCount < presentation.ezSlides.Count)
                {
                    navigation = "<div style=\"margin: 0px; position: absolute; top: 720px; left: 400px;\">" +
                "            <a href=\"html" + (htmlCount - 1) + ".html\" class=\"previous\">&laquo; Previous</a><a href=\"html" + (htmlCount) + ".html\" class=\"next\">Next &raquo;</a>" +
                "        </div>";
                    headHtml = headHtml + navigation;
                }


                string closeString;
                if (headHtml != "")
                {
                    closeString = @"</div></body>" + Environment.NewLine + @"</html > ";
                    headHtml = headHtml + closeString;
                }

                //creates a directory to save html files
                if (!Directory.Exists(PowerPointStudioRibbon.currentPPTPath + "\\HTML"))
                {
                    Directory.CreateDirectory(PowerPointStudioRibbon.currentPPTPath + "\\HTML");
                }

                System.IO.File.WriteAllText(PowerPointStudioRibbon.currentPPTPath + "\\HTML" + "\\html" + htmlCount + ".html", headHtml);
                htmlCount++;

            }

            if (File.Exists(PowerPointStudioRibbon.currentPPTPath + "\\HTML" + "\\html" + 0 + ".html"))
            {
                Process.Start(PowerPointStudioRibbon.currentPPTPath + "\\HTML" + "\\html" + 0 + ".html");
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No HTML found");
            }
        }

        /// <summary>
        /// Generate HTML from ezPresentation
        /// </summary>
        /// <param name="presentation"></param>
        public HtmlGenerator(ezPresentation presentation)
        {

        }

        /// <summary>
        /// Generate HTML of a single ezSlide
        /// </summary>
        /// <param name="slide"></param>
        public HtmlGenerator(ezSlide slide)
        {

        }


        /// <summary>
        /// This method takes required parameter for image and return its div string
        /// </summary>
        /// <param name="imageSourcePath"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <returns></returns>
        private string DivAdded(string imageSourcePath, float width, float height, float left, float top, bool NextArrow = false, bool RightArrow = false, float rotation = 0)
        {
            string divString = "";
            string sourcePath = imageSourcePath;

            divString = "<div style=\"margin: 0px\">" + Environment.NewLine + "<img src = "
                    + "\"" + sourcePath + "\"" + " alt = \"img\" style = \" position: absolute; width:"
                    + width + "px; height:" + height + "px; left:" + left
                    + "px; top:" + top + "px" + ";transform: rotate(" + rotation + "deg)" + " ;display: block; margin: 0px;\" >" + Environment.NewLine +
                    @"</div>";

            return divString;
        }


        private string SlideDiv(string imageSourcePath, float width, float height, float left, float top, bool NextArrow = false, bool RightArrow = false, float rotation = 0)
        {
            string divString = "";
            string sourcePath = imageSourcePath;

            divString = "<div style=\"margin: 0px; position: absolute; width: 960.00px; height:700.00px; left: 100px;border:1px solid black;\">" + Environment.NewLine + "<img src = "
                    + "\"" + sourcePath + "\"" + " alt = \"img\" style = \" position: absolute; width:"
                    + (width - 2) + "px; height:" + (height - 2) + "px; left:" + (left + 1)
                    + "px; top:" + (top + 1) + "px" + ";transform: rotate(" + rotation + "deg)" + " ;display: block; margin: 0px;\" >";

            return divString;
        }
    }
}
