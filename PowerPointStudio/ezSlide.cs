using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PowerPointStudio
{
    public class ezSlide
    {
        [JsonProperty(Order = 1)]
        public string sid { get; set; }  //Get from textbox

        [JsonProperty(Order = 2)]
        public object slide=new object(); 

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
                //Qulify shape Name
                ezPlacement placement = new ezPlacement();
                string placementText;
                //Handle placement
                if (shape.AlternativeText.Contains("$Placement$"))
                {
                    //find placement text
                    placementText = getPacementText(shape);
                    placementText = "{" + placementText + "}";
                    placement = Newtonsoft.Json.JsonConvert.DeserializeObject<ezPlacement>(placementText);
                    handlePlacement(shape, placement);
                }

                if(!(placement.onSlide==ezOnSlide.exceptFirst && slide.SlideIndex==1))
                {
                    ezShape shp = new ezShape(shape);
                    shapes.Add(shp);
                }
                
                //To get dnds
                if (shape.AlternativeText.Contains("$dnd$"))
                {
                    dnds dn = new dnds(shape);
                    dnd.Add(dn);
                }
                    
            }

            
            slideCount++;
           
        }

        internal static void handlePlacement(Shape shape, ezPlacement placement)
        {

            Presentation presentation = shape.Parent.Parent; //Getting the presentation object
            shape.AlternativeText = (Regex.Replace(shape.AlternativeText, @"\t|\n|\r", "")).Trim();
            string placeText = shape.AlternativeText;
            placeText = placeText.Substring(placeText.IndexOf("$Placement$") + 11, (placeText.IndexOf("$$Placement$$") - (placeText.IndexOf("$Placement$") + 11)));

            string toReplace = ("$Placement$" + placeText + "$$Placement$$").Trim();
            shape.AlternativeText = shape.AlternativeText.Replace(toReplace, "");
            
            shape.Copy();
            //Shape addedDuplicate = shape.Parent.Shapes[shape.Parent.Shapes.Count];

            
            //addedDuplicate.Copy();
            int slideIndex = shape.Parent.SlideIndex;

            switch (placement.onSlide)
            {
                case ezOnSlide.every:
                    //Copy this shape to every other shape to the same location except the placement string on alt text

                    foreach (Slide sld in presentation.Slides)
                    {
                        if (sld.SlideIndex != slideIndex)
                        {
                            sld.Shapes.Paste();
                        }

                    }
                    //addedDuplicate.Delete();
                    break;
                case ezOnSlide.exceptFirst:
                    foreach (Slide sld in presentation.Slides)
                    {
                        if (sld.SlideIndex != slideIndex && sld.SlideIndex != 1)
                        {
                            sld.Shapes.Paste();
                        }

                    }
                    //addedDuplicate.Delete();
                    break;
                case ezOnSlide.exceptLast:
                    foreach (Slide sld in presentation.Slides)
                    {
                        if (sld.SlideIndex != slideIndex && sld.SlideIndex != presentation.Slides.Count)
                        {
                            sld.Shapes.Paste();
                        }

                    }
                    //addedDuplicate.Delete();
                    break;
                case ezOnSlide.evenPages:
                    foreach (Slide sld in presentation.Slides)
                    {
                        int evenCheck = sld.SlideIndex % 2;
                        if (sld.SlideIndex != slideIndex && evenCheck == 0)
                        {
                            sld.Shapes.Paste();
                        }

                    }
                    //addedDuplicate.Delete();
                    break;
                case ezOnSlide.oddPages:
                    foreach (Slide sld in presentation.Slides)
                    {
                        int oddCheck = sld.SlideIndex % 2;
                        if (sld.SlideIndex != slideIndex && oddCheck != 0)
                        {
                            sld.Shapes.Paste();
                        }

                    }
                    //addedDuplicate.Delete();
                    break;
                default:
                    break;
            }

        }

        /// <summary>
        /// Getting placement text from altText
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        internal static string getPacementText(Shape shape)
        {
            string placementText = shape.AlternativeText;
            if (placementText.Contains("$Placement$"))
            {
                return placementText.Substring(placementText.IndexOf("$Placement$") + 11, (placementText.IndexOf("$$Placement$$") - (placementText.IndexOf("$Placement$") + 11)));
            }
            else
            {
                return "";
            }

        }
    }
}