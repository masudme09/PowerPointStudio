using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;

namespace PowerPointStudio
{
    public class ezText
    {
        public string value { get; set; }
        public ezTextStyle style { get; set; }//Bold, italic or underlined
        public float fontSize { get; set; }
        public string fontName { get; set; }

        [JsonConstructor]
        public ezText()
        {

        }


        public ezText(Shape shape)
        {
            if(shape.HasTextFrame==Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.TextRange.Text.Trim() != "")//To ensure that shape contains text
            {
                value = shape.TextFrame.TextRange.Text;

                //Getting text style
                if (shape.TextFrame.TextRange.Font.Bold == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    style = ezTextStyle.Bold;
                }
                else if (shape.TextFrame.TextRange.Font.Italic == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    style = ezTextStyle.Italic;
                }
                else if (shape.TextFrame.TextRange.Font.Underline == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    style = ezTextStyle.Underlined;
                }

                //Getting font size
                fontSize = shape.TextFrame.TextRange.Font.Size;
                fontName = shape.TextFrame.TextRange.Font.Name; 
            }
        }
    }
}