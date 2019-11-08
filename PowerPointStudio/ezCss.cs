using System;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;

namespace PowerPointStudio
{
    public class ezCss
    {
        public string width { get; set; }
        public string height { get; set; }
        public string left { get; set; }
        public string top { get; set; }
        public float rotation { get; set; }
        public int zIndex { get; set; }

        /// <summary>
        /// Just for deserialization
        /// Without default constructor, it is not possible to deserialize
        /// </summary>
        [JsonConstructor]
        public ezCss()
        {

        }


        public ezCss(float width, float height, float left, float top, int zIndex=0, float rotation = 0)
        {
            EzCssForWeb(width, height, left, top, zIndex, rotation);
            
        }

       /// <summary>
       /// Generate CSS from powerpoint shape
       /// </summary>
       /// <param name="shape"></param>

        public ezCss(Shape shape)
        {
            //Check shape type and do action accordingly
            ezShapeType shapeType = Utility.GetShapeType(shape);

            switch(shapeType)
            {
                case ezShapeType.Group:
                    Group grp = new Group(shape);
                    EzCssForWeb(grp.Width, grp.Height, grp.Left, grp.Top, grp.Zindex, grp.Rotation);
                    break;
                case ezShapeType.EllipseCallout:
                    WedgeEllipseCallout ellipseCallout = new WedgeEllipseCallout(shape);
                    EzCssForWeb(ellipseCallout.calculatedWidth, ellipseCallout.calculatedHeight, ellipseCallout.calculatedLeft, 
                        ellipseCallout.calculatedTop, ellipseCallout.Zindex, ellipseCallout.Rotation);
                    break;
                case ezShapeType.TextBox:
                    TexBox textox = new TexBox(shape);
                    EzCssForWeb(textox.width, textox.height, textox.left,
                        textox.top, textox.Zindex, textox.Rotation);
                    break;
                default:
                    float thickness = shape.Line.Weight;
                    //Thickness may come incorrect if there is no thickness
                    if (thickness < 0 || thickness > 20)
                    {
                        thickness = 0;
                    }

                    EzCssForWeb((float)(shape.Width + thickness), (float)(shape.Height + thickness), 
                        (float)(shape.Left - thickness / 2.0), (float)(shape.Top - thickness / 2.0), 
                        shape.ZOrderPosition, shape.Rotation);
                    break;
            }


        }

        /// <summary>
        /// This method is to generated original data from powerpoint
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="zIndex"></param>
        /// <param name="rotation"></param>
        private void EzCss(float width, float height, float left, float top, int zIndex = 0, float rotation = 0)
        {
            this.width = String.Format("{0:0.00}", width) + "px";
            this.height = String.Format("{0:0.00}", height) + "px";
            this.left = String.Format("{0:0.00}", left) + "px";
            this.top = String.Format("{0:0.00}", top) + "px";
            this.rotation = rotation;
            this.zIndex = zIndex;
        }

        /// <summary>
        /// To generate css for our custom web page size 
        /// Ratio is 1.6667
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="zIndex"></param>
        /// <param name="rotation"></param>
        private void EzCssForWeb(float width, float height, float left, float top, int zIndex = 0, float rotation = 0)
        {
            float ratio =(float) (960 / 576.0);
            width = (float) (width * ratio);
            height = (float)(height * ratio);
            left = (float)(left * ratio);
            top = (float)(top * ratio);

            this.width = String.Format("{0:0.00}", width) + "px";
            this.height = String.Format("{0:0.00}", height) + "px";
            this.left = String.Format("{0:0.00}", left) + "px";
            this.top = String.Format("{0:0.00}", top) + "px";
            this.rotation = rotation;
            this.zIndex = zIndex;
        }
    }
}