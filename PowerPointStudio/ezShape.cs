using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;

namespace PowerPointStudio
{
    public class ezShape
    {
        [JsonProperty(Order = 1)]
        public string id { get; set; } //shape name

        [JsonProperty(Order = 2)]
        public string @class { get; set; }

        [JsonProperty(Order = 3)]
        public ezImage image;
       
        [JsonProperty(Order = 4)]
        public ezText text = null;

        [JsonProperty(Order = 5)]
        public ezAction actions { get; set; }

        [JsonProperty(Order = 6)]
        public string audioUrl { get; set; }           

        internal static int shapeCount=0;
        internal static int mediaCount = 0;//to track number of medias detected

        [JsonConstructor]
        public ezShape(string id,ezImage image, ezText text=null,string @class=null)
        {
            this.id = id;
            this.image = image;
            this.text = text;
            this.@class = @class;
        }


        public ezShape(Shape shape)
        {
                        
            string qulifiedShapeName = Utility.qulifiedNameGenerator(shape.Name);
            id = "sh" + qulifiedShapeName;
            @class = "temp"; //need to get it from alt text. if not found default is 'temp'
            if(shape.AlternativeText.Contains("$class$"))
            {
                @class = classFinder(shape.AlternativeText);
            }
            //text = new ezText(shape); //Need to its structure..When instructed

            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia && shape.MediaType == PpMediaType.ppMediaTypeSound)
            {
                
                audioUrl = Utility.getExtractedAudioUrl(shape);
                if (audioUrl != null)
                {
                    this.audioUrl = audioUrl;
                }

            }

            image = new ezImage(shape);
           
            actions = new ezAction(shape);
            
            shapeCount++;
        }

       

        private string classFinder(string altText)
        {
            string classContain = null;

            classContain = altText.Substring(altText.IndexOf("$class$")+7, (altText.IndexOf("$$class$$")- (altText.IndexOf("$class$")+7))); //It is returning first character index of the searched string

            return classContain;
        }
    }
}