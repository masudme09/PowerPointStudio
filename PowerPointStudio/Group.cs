using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointStudio
{
    internal class Group
    {
        public float Width { get; set; }
        public float Height { get; set; }
        public float Top { get; set; }
        public float Left { get; set; }
        public float Rotation { get; set; }
        public int Zindex { get; set; }
        public int itemCount { get; set; }//Holds number of items in a group

        /// <summary>
        /// Extract properties from Group Type shape
        /// </summary>
        /// <param name="groupTypeShape"></param>
        public Group(Shape groupTypeShape)
        {
            //Assigning primitive properties first
            //For group type shape primitive properties we need not to consider thickness as group does not have thickness
            itemCount = groupTypeShape.GroupItems.Count;
            Rotation = groupTypeShape.Rotation;
            Zindex = groupTypeShape.ZOrderPosition;
            Width = groupTypeShape.Width;
            Height = groupTypeShape.Height;
            Top = groupTypeShape.Top;
            Left = groupTypeShape.Left;

            //For Calculation 
            float grpMaxLeft = Left, grpMaxTop = Top, grpMaxRight = Left + Width, grpMaxBottom = Top + Height; //initialization

            //Now traversing though each element of the group and revising the boundaries
            foreach(Shape shp in groupTypeShape.GroupItems)
            {
                //Getting the type of shape 
                ezShapeType shapeType = Utility.GetShapeType(shp);

                switch(shapeType)
                {
                    case ezShapeType.EllipseCallout:
                        WedgeEllipseCallout ellipseCallout = new WedgeEllipseCallout(shp);

                        if(grpMaxLeft>ellipseCallout.calculatedLeft)
                        {
                            grpMaxLeft = ellipseCallout.calculatedLeft;
                        }

                        if(grpMaxTop>ellipseCallout.calculatedTop)
                        {
                            grpMaxTop = ellipseCallout.calculatedTop;
                        }

                        if(grpMaxRight<ellipseCallout.maxRight)
                        {
                            grpMaxRight = ellipseCallout.maxRight;
                        }

                        if(grpMaxBottom<ellipseCallout.maxBottom)
                        {
                            grpMaxBottom = ellipseCallout.maxBottom;
                        }
                        break;

                    case ezShapeType.Group:
                        Group group = new Group(shp);
                        if (grpMaxLeft > group.Left)
                        {
                            grpMaxLeft = group.Left;
                        }

                        if (grpMaxTop > group.Top)
                        {
                            grpMaxTop = group.Top;
                        }

                        if (grpMaxRight < group.Left+group.Width)
                        {
                            grpMaxRight = group.Left + group.Width;
                        }

                        if (grpMaxBottom < group.Top + group.Height)
                        {
                            grpMaxBottom = group.Top + group.Height;
                        }
                        break;
                    default:
                        break;
                }

            }

            //Assigning the updated values
            Width = grpMaxRight-grpMaxLeft;
            Height = grpMaxBottom-grpMaxTop;
            Top = grpMaxTop;
            Left = grpMaxLeft;
        }
    }
}