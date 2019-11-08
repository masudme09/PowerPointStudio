using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Drawing;
using System.IO;

namespace PowerPointStudio
{
    internal class TexBox
    {
        public float left { get; set; }
        public float top { get; set; }
        public float width { get; set; }
        public float height { get; set; }
        public int Zindex { get; set; }
        public float Rotation { get; set; }

        public TexBox(Shape shape)
        {            
            //Calculate the midpoint of the shape
            float midX = shape.Left + (float) (shape.Width / 2.0);
            float midY = shape.Top + (float)(shape.Height / 2.0);

            //Export the picture and get mid point of the picture
            Bitmap textBoxBitmap = exportSingleShapeAsImage(shape);

            //float imageToPoint = (float)(502.0 / 376.88);
            float imageToPointX = (float)(1.34214742);
            float imageToPointY = (float)(1.343623496);
            float exImageWidth = textBoxBitmap.Width / (float)(imageToPointX);//* 0.9954
            float exImageHeight = textBoxBitmap.Height / (float)(imageToPointY); //*0.9954


            left = midX - exImageWidth / 2;
            top = midY - exImageHeight / 2;
            width = exImageWidth;
            height = exImageHeight;
            Zindex = shape.ZOrderPosition;
            Rotation = shape.Rotation;

            //After finish work dispose the bitmap to clear resource
            textBoxBitmap.Dispose();

        }

        /// <summary>
        /// Export single shape to Bitmap image and return bitmap
        /// </summary>
        /// <param name="currentPresentation"></param>
        /// <param name="shp"></param>
        /// <returns></returns>
        public static Bitmap exportSingleShapeAsImage(Shape shp)
        {
            if (!Directory.Exists(@"C:\temp"))
            {
                Directory.CreateDirectory(@"C:\temp");
            }

            string exportImagePath = @"C:\temp\tempPic.png";
            if (File.Exists(exportImagePath))
            {
                File.Delete(exportImagePath);
            }
            Presentation currentPresentation = shp.Parent.Parent;
            Bitmap origialImage = null;
            
            try
            {                                
                //Save the shape dimensions
                int slideWidth = (int)Utility.SlideWidthGet(currentPresentation);
                int slideHeight = (int)Utility.SlideHeightGet(currentPresentation);          
               
                shp.Export(exportImagePath, PpShapeFormat.ppShapeFormatPNG);//, slideWidth, slideHeight);
                            
                origialImage = new Bitmap(exportImagePath);

            }
            catch (Exception err)
            {
                //MessageBox.Show(err.ToString());

            }
                       
            return origialImage;
        }
    }
}