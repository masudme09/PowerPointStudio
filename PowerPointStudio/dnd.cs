using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PowerPointStudio
{
    public class dnds
    {
        public string dragitem;
        public string droptarget;
        public string audioondrop;
        public string pictureondrop;
        public string shapeondrop;

        [JsonConstructor]
        public dnds()
        {

        }

        public dnds(Shape shape)
        {
            if (shape.AlternativeText.Contains("$dnd$"))
            {
                string dnd = "{"+Regex.Replace(dndFinder(shape.AlternativeText).Trim(), @"\t|\n|\r", "")+"}";
                try
                {
                    dnds parsedDnd = Newtonsoft.Json.JsonConvert.DeserializeObject<dnds>(dnd);
                    this.dragitem = parsedDnd.dragitem;
                    this.droptarget = parsedDnd.droptarget;
                    this.audioondrop = parsedDnd.audioondrop;
                    this.pictureondrop = parsedDnd.pictureondrop;
                    this.shapeondrop = parsedDnd.shapeondrop;
                }
                catch(Exception err)
                {
                    MessageBox.Show("dnd is not in JSON format. Couldn't deserialize/n"+err.ToString());
                }
               
                
            }
        }

        private string dndFinder(string altText)
        {
            string dndContain = null;

            dndContain = altText.Substring(altText.IndexOf("$dnd$") + 5, (altText.IndexOf("$$dnd$$") - (altText.IndexOf("$dnd$") + 5))); //It is returning first character index of the searched string

            return dndContain;
        }

    }
}