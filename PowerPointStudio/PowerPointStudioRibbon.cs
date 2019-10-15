using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Threading;
using System.Diagnostics;


namespace PowerPointStudio
{
    public partial class PowerPointStudioRibbon
    {
        public static string currentPPTPath=""; //This store current active presentation path
        public static string mediaPath = "";

        private void PowerPointStudioRibbon_Load(object sender, RibbonUIEventArgs e)
        {
                        
        }

        private void BtnExtractSlides_Click(object sender, RibbonControlEventArgs e)
        {
            Application pptApp = new Application();
            Presentation presentation = pptApp.ActivePresentation;
            string pptPath = presentation.Path; //Provides directory

            //Copying this presentation to the same original directory with temp name
            //All the work will be done on the copied presentation
            //So not visible to user

            while(Directory.Exists(pptPath+@"\temp"))
            {
                try
                {
                    Directory.Delete(pptPath + @"\temp", true);
                }catch
                {

                }
            }
                        
            Directory.CreateDirectory(pptPath + @"\temp");
            
            File.Copy(pptPath + @"\"+presentation.Name, pptPath + @"\temp\temp.pptx");

            Presentation tempPresentation = pptApp.Presentations.Open(pptPath + @"/temp/temp.pptx",Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            currentPPTPath = tempPresentation.Path;

            //Extract medias from this presentation and copy them in a medias folder
            //Setting media path
            mediaPath = Utility.createZipAndExtract(tempPresentation);  //Getting copy of the current presentation with zip extension path


            //Extract info from this tempPresentation
            ezPresentation cusPresentation = new ezPresentation(tempPresentation);

            //Creating and writing JSON
            Utility.writeJsonToFile(cusPresentation, currentPPTPath + "\\Json.JSON");


            //Close the presentation
            tempPresentation.Close();
            System.Windows.Forms.MessageBox.Show("Extraction Complete");
            
        }

        private void BtnPreviewJSON_Click(object sender, RibbonControlEventArgs e)
        {
            if(File.Exists(currentPPTPath+ "\\Json.JSON") && currentPPTPath!="")
            {
                Process.Start(currentPPTPath + "\\Json.JSON");
            }else
            {
                System.Windows.Forms.MessageBox.Show("No JSON file exists. Please first click on extract to generate JSON");
            }
        }

        private void BtnPreviewWeb_Click(object sender, RibbonControlEventArgs e)
        {
            if(File.Exists(currentPPTPath + "\\Json.JSON"))
            {
                string JSON = File.ReadAllText(currentPPTPath + "\\Json.JSON");
                HtmlGenerator htmlGenerator = new HtmlGenerator(JSON);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No JSON file found on default directory. Please click on Extract to generate.");
            }
            
        }
    }
}
