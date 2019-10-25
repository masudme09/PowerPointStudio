using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace PowerPointStudio
{
    public class ezSlide
    {
        [JsonProperty(Order = 1)]
        public string sid { get; set; }  //Get from textbox

        [JsonProperty(Order = 2)]
        public List<object> slide=new List<object>(); 

        [JsonProperty(Order = 3)]
        public ezShapes<ezShape> shapes = new ezShapes<ezShape>();

        [JsonProperty(Order = 4)]
        public ezSlideAnimations<ezSlideAnimation> slide_animations = new ezSlideAnimations<ezSlideAnimation>();//implement later

        [JsonProperty(Order = 5)]
        public ezDnd<dnds> dnd = new ezDnd<dnds>(); //implement later

        private static int slideCount=0;

        [JsonConstructor]
        public ezSlide()
        {

        }

        //With every ezSlide instance new sid will be created
        public ezSlide(Slide slide)
        {
            if(Globals.Ribbons.PowerPointStudioRibbon.ediBxExerKey.Text!="")
            {
                sid = Globals.Ribbons.PowerPointStudioRibbon.ediBxExerKey.Text +"_S"+slide.SlideIndex.ToString("000");//String.Format("{0:0.00}", width)
            }

            //Create background and convert that to ezshape to assign that to shape
            //As slide background belongs to shape 
            ezBackGround backGround = new ezBackGround(slide);
            ezShape backgroundShape = new ezShape(backGround.id, backGround.image,null,"temp");
            shapes.Add(backgroundShape);


            //Assigning ezShape to shapes
            foreach (Shape shape in slide.Shapes)
            {
                ezShape shp = new ezShape(shape);
                shapes.Add(shp);

                //To get dnds
                if (shape.AlternativeText.Contains("$dnd$"))
                {
                    dnds dn = new dnds(shape);
                    dnd.Add(dn);
                }
                    
            }

            
            slideCount++;
           
        }
    }
}