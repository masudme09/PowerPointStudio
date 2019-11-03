using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;

namespace PowerPointStudio
{
    public class ezAction
    {
        public string onClick { get; set; }
        public string onHover { get; set; }
        public string onLoad { get; set; }

        [JsonConstructor]
        public ezAction()
        {

        }

        public ezAction(Shape shape)
        {
            string altText = shape.AlternativeText;


            if (altText.Contains("$action$"))
            {
                string actionText = Utility.ezLangFinder(altText, "action");
                actionText = "{" + actionText + "}";
                ezAction parsedAction = Newtonsoft.Json.JsonConvert.DeserializeObject<ezAction>(actionText);

                if (parsedAction.onClick.ToLower().Contains("next"))
                {
                    this.onClick = "next." + Globals.Ribbons.PowerPointStudioRibbon.ediBxExerKey.Text + "_S" + (shape.Parent.SlideIndex + 1).ToString("000"); //sid = Globals.Ribbons.PowerPointStudioRibbon.ediBxExerKey.Text +"_S"+slide.SlideIndex.ToString("000");
                }
                else if (parsedAction.onClick.ToLower().Contains("back"))
                {
                    this.onClick = "next." + Globals.Ribbons.PowerPointStudioRibbon.ediBxExerKey.Text + "_S" + (shape.Parent.SlideIndex - 1).ToString("000");
                }
                else
                {
                    this.onClick = parsedAction.onClick;
                }

                this.onHover = parsedAction.onHover;
                this.onLoad = parsedAction.onLoad;
            }
            
        }

        
    }
}