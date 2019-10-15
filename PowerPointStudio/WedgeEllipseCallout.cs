using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows.Forms;

namespace PowerPointStudio
{
    internal class WedgeEllipseCallout
    {
        /// <summary>
        /// width of the rectagle bound
        /// </summary>
        private float w { get; set; }//width
        private float h { get; set; }//height of the rectangle bound
        private float adj1 { get; set; }
        private float adj2 { get; set; }
        private float dxPos;
        private float dyPos;
        private float xPos { get; set; }
        private float yPos { get; set; }
        private float sdx;
        private float sdy;
        private float pang { get; set; }
        private float stAng { get; set; }
        private float enAng { get; set; }
        private float dx1 { get; set; }
        private float dy1 { get; set; }
        private float x1 { get; set; }
        private float y1 { get; set; }
        private float dx2 { get; set; }
        private float dy2 { get; set; }
        private float x2 { get; set; }
        private float y2 { get; set; }
        private float stAng1 { get; set; }
        private float enAng1 { get; set; }
        private float swAng1 { get; set; }
        private float swAng2 { get; set; }
        private float swAng { get; set; }
        private float idx { get; set; }
        private float idy { get; set; }
        private float il { get; set; }
        private float ir { get; set; }
        private float it { get; set; }
        private float ib { get; set; }
        private float callOutMaxPointX { get; set; }
        private float callOutMaxPointY { get; set; }

        public float calculatedWidth { get; set; }
        public float calculatedHeight { get; set; }
        public float calculatedTop { get; set; }
        public float calculatedLeft { get; set; }
        public float maxRight { get; set; }
        public float maxBottom { get; set; }
        public float Rotation { get; set; }
        public int Zindex { get; set; }

        /// <summary>
        /// Generating from shape
        /// </summary>
        /// <param name="ovalShapeCallout"></param>
        public WedgeEllipseCallout(Shape ovalShapeCallout)
        {
            //Set rotation
            Rotation = ovalShapeCallout.Rotation;
            //Set Zindex
            Zindex = ovalShapeCallout.ZOrderPosition;

            //Calculate the variables in EMU unit
            //Getting width, height, adj1 and adj2
            w = ovalShapeCallout.Width * 12700;//Converting to EMU.To get formulas right
            h = ovalShapeCallout.Height * 12700;
            try
            {
                adj1 = ovalShapeCallout.Adjustments[1] * 100000;//Converting to openXML value
                adj2 = ovalShapeCallout.Adjustments[2] * 100000;
            }
            catch (Exception err)
            {
                MessageBox.Show("Adjustment Value error for oval callout shape");
            }


            float cnstVal1 = (float)(100000 * 72.0 / 914400.0);
            float angVal1 = (float)(11 * Math.PI / 180);
            w = w * (float)72.0 / 914400;
            h = h * (float)72.0 / 914400;
            adj1 = adj1 * (float)72.0 / 914400;
            adj2 = adj2 * (float)72.0 / 914400;

            float ss = Math.Min(w, h);
            float dxPos, dyPos, xPos, yPos, sdx, sdy, pang, stAng, enAng, dx1, dy1, x1, y1, dx2, dy2,
                x2, y2, stAng1, enAng1, swAng1, swAng2, swAng,
                vc = h / 2, hc = w / 2;
            dxPos = w * adj1 / cnstVal1;
            dyPos = h * adj2 / cnstVal1;
            xPos = hc + dxPos;
            yPos = vc + dyPos;
            sdx = dxPos * h;
            sdy = dyPos * w;
            pang = (float)Math.Atan(sdy / sdx);
            stAng = pang + angVal1;
            enAng = pang - angVal1;

            dx1 = (float)(hc * Math.Cos(stAng));
            dy1 = (float)(vc * Math.Sin(stAng));
            dx2 = (float)(hc * Math.Cos(enAng));
            dy2 = (float)(vc * Math.Sin(enAng));

            if (dxPos >= 0)
            {
                x1 = hc + dx1;
                y1 = vc + dy1;
                x2 = hc + dx2;
                y2 = vc + dy2;
            }
            else
            {
                x1 = hc - dx1;
                y1 = vc - dy1;
                x2 = hc - dx2;
                y2 = vc - dy2;
            }

            //test
            pang = (float)(Math.Atan(sdy / sdx) * (180 / Math.PI));
            //Calculating slope
            float m1 = (yPos - y1) / (xPos - x1);
            float m2 = (yPos - y2) / (xPos - x2);
            float angle1 = (float)(Math.Atan(m1) * (180 / Math.PI));
            float angle2 = (float)(Math.Atan(m2) * (180 / Math.PI));
            float minAngle = Math.Min(angle1, angle2);
            //Outer line travesing point
            float xOut1 = 0;
            float xOut2 = 0;
            float yOut1 = y1;
            float yOut2 = y2;

            float const1 = (float)(ovalShapeCallout.Line.Weight / 2.0 * Math.Sqrt(Math.Pow(m1, 2) + 1));
            float const2 = (float)(ovalShapeCallout.Line.Weight / 2.0 * Math.Sqrt(Math.Pow(m2, 2) + 1));
            float c1m1 = yPos - m1 * xPos;
            float c1m2 = yPos - m2 * xPos;
            float c2m2;
            float c2m1;
            if (xPos > hc)
            {
                c2m2 = c1m2 - const2;
                c2m1 = c1m1 + const1;
            }
            else
            {
                c2m2 = c1m2 + const2;
                c2m1 = c1m1 - const1;
            }



            if (x1 < x2)
            {
                xOut1 = x1 - (ovalShapeCallout.Line.Weight / 2);
                xOut2 = x2 + (ovalShapeCallout.Line.Weight / 2);
            }
            else if (x1 > x2)
            {
                xOut1 = x1 + (ovalShapeCallout.Line.Weight / 2);
                xOut2 = x2 - (ovalShapeCallout.Line.Weight / 2);
            }
            else if (y1 > y2)
            {

            }
            else if (y1 < y2)
            {

            }


            float xOut = (c2m1 - c2m2) / (m2 - m1);
            float yOut = (m2 * c2m1 - m1 * c2m2) / (m2 - m1);
            callOutMaxPointX = xOut;
            callOutMaxPointY = yOut;

            //Convert this point based on canvas origin
            callOutMaxPointX = ovalShapeCallout.Left + callOutMaxPointX;
            callOutMaxPointY = ovalShapeCallout.Top + callOutMaxPointY;
            //Calculating boundary
            float maxLeft = ovalShapeCallout.Left - (float)(ovalShapeCallout.Line.Weight / 2.0);
            float maxTop = ovalShapeCallout.Top - (float)(ovalShapeCallout.Line.Weight / 2.0);
            if (maxLeft > callOutMaxPointX)
            {
                maxLeft = callOutMaxPointX;
            }
            if (maxTop > callOutMaxPointY)
            {
                maxTop = callOutMaxPointY;
            }

            if ((ovalShapeCallout.Left + ovalShapeCallout.Width + (ovalShapeCallout.Line.Weight / 2.0)) > callOutMaxPointX)
            {
                maxRight = ovalShapeCallout.Left + ovalShapeCallout.Width + (float)(ovalShapeCallout.Line.Weight / 2.0);
            }
            else
            {
                maxRight = callOutMaxPointX;
            }

            if ((ovalShapeCallout.Top + ovalShapeCallout.Height + (ovalShapeCallout.Line.Weight / 2.0)) > callOutMaxPointY)
            {
                maxBottom = ovalShapeCallout.Top + ovalShapeCallout.Height + (float)(ovalShapeCallout.Line.Weight / 2.0);
            }
            else
            {
                maxBottom = callOutMaxPointY;
            }

            calculatedHeight = maxBottom - maxTop;
            calculatedWidth = maxRight - maxLeft;
            calculatedLeft = maxLeft;
            calculatedTop = maxTop;


            stAng1 = (float)Math.Atan2(dy1, dx1);
            enAng1 = (float)Math.Atan2(dy2, dx2);
            swAng1 = enAng1 - stAng1;
            swAng2 = swAng1 + 21600000;
            swAng = (swAng1 > 0) ? swAng1 : swAng2;
            idx = (float)(w / 2 * Math.Cos(2700000));
            idy = (float)(h / 2 * Math.Sin(2700000));
            il = w / 2 - idx;//1345084.25
            ir = w / 2 + idx;//369415.7
            it = h / 2 - idy;//65072.03
            ib = h / 2 + idy;//667267

        }
    }
}