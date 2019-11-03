using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Threading;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Data;
using MessageBox = System.Windows.Forms.MessageBox;

namespace PowerPointStudio
{
    public partial class PowerPointStudioRibbon
    {
        public static string currentPPTPath=""; //This store current active presentation path
        public static string mediaPath = "";
        private static int prevTotalShapeCount=0;
        private static int prevTotalSlideCount = 0;
        private static string prevPresentationName = "";

        private void PowerPointStudioRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Settings.ReadAndUpdateFromXML(); //Loading previous settings
        }

        private void BtnExtractSlides_Click(object sender, RibbonControlEventArgs e)
        {
            extractSlide();
        }

        /// <summary>
        /// Do all the extraction to generate JSON
        /// </summary>
        public void extractSlide()
        {
            if (Globals.Ribbons.PowerPointStudioRibbon.ediBxExerKey.Text == "")
            {
                System.Windows.Forms.MessageBox.Show("Please Enter Excercise Key First");
                return;
            }
            Application pptApp = new Application();
            Presentation presentation = pptApp.ActivePresentation;
            prevPresentationName = Globals.ThisAddIn.Application.ActivePresentation.Name; //it will be stored in a static variable for later use

            if (presentation.Name.Contains("pptx")) //pptx files are only extractable
            {


                string pptPath = presentation.Path; //Provides directory

                //Copying this presentation to the same original directory with temp name
                //All the work will be done on the copied presentation
                //So not visible to user
                if (File.Exists(pptPath + @"\temp\CSV.csv"))
                {
                    try
                    {
                        File.Delete(pptPath + @"\temp\CSV.csv");
                        if (File.Exists(pptPath + @"\temp\Json.JSON"))
                        {
                            File.Delete(pptPath + @"\temp\Json.JSON");
                        }

                    }
                    catch
                    {
                        System.Windows.Forms.MessageBox.Show("Please close the CSV or JSON file if already opened");
                    }
                }


                while (Directory.Exists(pptPath + @"\temp"))
                {
                    //Waiting until the temp directory is successfully deleted
                    try
                    {
                        Directory.Delete(pptPath + @"\temp", true);
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(err.ToString());
                    }
                }

                Directory.CreateDirectory(pptPath + @"\temp");

                File.Copy(pptPath + @"\" + presentation.Name, pptPath + @"\temp\temp.pptx");

                Presentation tempPresentation = pptApp.Presentations.Open(pptPath + @"/temp/temp.pptx", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
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

                Utility.staticResourceClear();
                prevTotalShapeCount = getTotalShapeCount(presentation);
                prevTotalSlideCount = presentation.Slides.Count;

            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Extraction is only possible with pptx files");
            }
        }

        private void BtnPreviewJSON_Click(object sender, RibbonControlEventArgs e)
        {
            if (isPresentationChanged() == true)
            {
                MessageBox.Show("Presentation updated! Extarcting first to get updated output");
                extractSlide();
            }

            if (File.Exists(currentPPTPath+ "\\Json.JSON") && currentPPTPath!="")
            {
                Process.Start(currentPPTPath + "\\Json.JSON");
            }else
            {
                System.Windows.Forms.MessageBox.Show("No JSON file exists. Please first click on extract to generate JSON");
            }
        }

        private void BtnPreviewCSV_Click(object sender, RibbonControlEventArgs e)
        {
            if (isPresentationChanged() == true)
            {
                MessageBox.Show("Presentation updated! Extarcting first to get updated output");
                extractSlide();
            }

            if (File.Exists(currentPPTPath + "\\Json.JSON"))
            {
                string JSON = File.ReadAllText(currentPPTPath + "\\Json.JSON");
                System.Data.DataTable dataTable = GetDataTableFromJSON(JSON);
                //Write this dataTable to file as CSV
                Utility.ToCSV(dataTable, currentPPTPath + "\\CSV.csv");

                while(!File.Exists(currentPPTPath + "\\CSV.csv")) { }//Until CSV file is created looping through
                Process.Start(currentPPTPath + "\\CSV.csv");

            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No JSON file found on default directory. Please click on Extract to generate.");
            }
        }

        private System.Data.DataTable GetDataTableFromJSON(string JSON)
        {

            System.Data.DataTable dataTable = new System.Data.DataTable();
            ezPresentation dt = JsonConvert.DeserializeObject<ezPresentation>(JSON);
            dataTable.Columns.Add("sid", typeof(string));
            dataTable.Columns.Add("shapeId", typeof(string));
            dataTable.Columns.Add("shClass", typeof(string));
            dataTable.Columns.Add("objectType", typeof(string));//can not get this from JSON
            dataTable.Columns.Add("width", typeof(string));
            dataTable.Columns.Add("height", typeof(string));
            dataTable.Columns.Add("left", typeof(string));
            dataTable.Columns.Add("top", typeof(string));
            dataTable.Columns.Add("rotation", typeof(float));
            dataTable.Columns.Add("zindex", typeof(int));
            dataTable.Columns.Add("imagePath", typeof(string));
            dataTable.Columns.Add("uploadImagePath", typeof(string));
            dataTable.Columns.Add("imageUrlLarge", typeof(string));
            dataTable.Columns.Add("imageUrlMedium", typeof(string));
            dataTable.Columns.Add("imageUrlSmall", typeof(string));
            dataTable.Columns.Add("onClick", typeof(string));
            dataTable.Columns.Add("onHover", typeof(string));
            dataTable.Columns.Add("onLoad", typeof(string));
            dataTable.Columns.Add("audioUrl", typeof(string));

            foreach (ezSlide sld in dt.ezSlides)
            {
                //DataRow dataRow=null;
               foreach (ezShape shp in sld.shapes)
               {
                    DataRow dataRow = dataTable.NewRow();
                    dataRow[0] = sld.sid;
                    dataRow[1] = shp.id;//shape id
                    dataRow[2] = shp.@class;//shape.class
                    dataRow[3] = "N//A";//objectType
                    if(shp.image.css!=null)
                    {
                        dataRow[4] = shp.image.css.width;//width
                        dataRow[5] = shp.image.css.height;//height
                        dataRow[6] = shp.image.css.left;//left
                        dataRow[7] = shp.image.css.top;//top
                        dataRow[8] = shp.image.css.rotation;//rotation
                        dataRow[9] = shp.image.css.zIndex;//zindex
                    }
                    
                    if(shp.image.imgurlLarge!=null)
                    {
                        dataRow[10] = currentPPTPath + (shp.image.imgurlLarge.Replace("https://ezilmdev.org", "")).Replace("/", @"\");//imagePath Actual
                        dataRow[11] = sld.sid.Replace("_", "/") + @"/" + sld.sid + "-" + (shp.image.imgurlLarge.Replace("https://ezilmdev.org/images/", ""));//uploadImagePath
                        dataRow[12] = shp.image.imgurlLarge;//imageUrlLarge
                        dataRow[13] = shp.image.imgurlMedium;//imageUrlMedium
                        dataRow[14] = shp.image.imgurlSmall;//imageUrlSmall
                    }
                    
                    if (shp.actions != null)
                    {
                        dataRow[15] = (shp.actions.onClick != null) ? shp.actions.onClick : "";//onClick
                        dataRow[16] = (shp.actions.onHover != null) ? shp.actions.onHover : "";//onHover
                        dataRow[17] = (shp.actions.onLoad != null) ? shp.actions.onLoad : "";//onLoad
                    }
                    dataRow[18] = shp.audioUrl!=null? shp.audioUrl:"";//audioUrl
                    dataTable.Rows.Add(dataRow);

                }
                

            }
           

            return dataTable;

        }

        /// <summary>
        /// Extract audios with the name of arabic text that appears on that slide
        /// if one audio with multiple text the concat all the text to name the audio
        /// if mutiple audio in a slide then use numbering to name them
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExtractAudio_Click(object sender, RibbonControlEventArgs e)
        {
            if(isPresentationChanged()==true)
            {
                MessageBox.Show("Presentation updated! Extarcting first to get updated output");
                extractSlide();
            }

           
            if (currentPPTPath!="")
            {
                Application pptApp = new Application();
                if(File.Exists(currentPPTPath + @"/temp.pptx"))
                {
                    Presentation presentation = pptApp.Presentations.Open(currentPPTPath + @"/temp.pptx", Microsoft.Office.Core.MsoTriState.msoFalse, 
                        Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
                    Slides slides = presentation.Slides;
                    int counta = 0;
                    foreach (Slide slide in slides)
                    {
                        
                        Shapes shapes = slide.Shapes;
                        foreach(Shape shape in shapes)
                        {
                            //If any shape is media type then search for arabic text and concat them 
                            //Then find the audio or media of that shape and copy to another directory with the arabic name
                            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia && shape.MediaType == PpMediaType.ppMediaTypeSound)
                            {
                               
                                string audioUrl = Utility.getExtractedAudioUrl(shape);
                                string audioName = "";
                                foreach(Shape shape1 in shapes)
                                {
                                    if(shape1.Name.ToLower().Contains("arabic"))
                                    {
                                        if(shape1.HasTextFrame==Microsoft.Office.Core.MsoTriState.msoTrue)
                                        {
                                            if(shape1.TextFrame.TextRange.Text!="")
                                            {
                                                audioName = audioName + shape1.TextFrame.TextRange.Text;
                                            }
                                        }
                                        
                                    }
                                    if(shape1.Type==Microsoft.Office.Core.MsoShapeType.msoGroup)
                                    {
                                        foreach(Shape shape2 in shape1.GroupItems)
                                        {
                                            if (shape2.Name.ToLower().Contains("arabic"))
                                            {
                                                if (shape2.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                                {
                                                    if (shape2.TextFrame.TextRange.Text != "")
                                                    {
                                                        audioName = audioName + shape2.TextFrame.TextRange.Text;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                //If no arabic text contain in a slide then search for all english text
                                if(audioName=="")
                                {
                                    foreach (Shape shape1 in shapes)
                                    {
                                        if (shape1.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                        {
                                            if (shape1.TextFrame.TextRange.Text != "")
                                            {
                                                audioName = audioName + shape1.TextFrame.TextRange.Text;
                                            }
                                        }
                                            
                                        if (shape1.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                        {
                                            foreach (Shape shape2 in shape1.GroupItems)
                                            {
                                                if(shape2.HasTextFrame==Microsoft.Office.Core.MsoTriState.msoTrue)
                                                {
                                                    if (shape2.TextFrame.TextRange.Text != "")
                                                    {
                                                        audioName = audioName + shape2.TextFrame.TextRange.Text;
                                                    }
                                                }                                                                                            
                                               
                                            }
                                        }
                                    }
                                    
                                }
                                //if still audioname is empty..means no text available on the slide
                                //asign noName_randomNumber as name
                                if(audioName == "")
                                {
                                    audioName = "noName_" + Utility.RandomNumber(1, 10000, counta++);
                                }

                                //Now create a directory and copy audio to that directory with arabic name
                                string audioDirectory = currentPPTPath+ @"\Audios";
                                if (!Directory.Exists(audioDirectory))
                                {
                                    Directory.CreateDirectory(audioDirectory);
                                }

                                //Copy the file with this arabic name
                                string newPath = audioDirectory+@"\"+audioName+"."+ audioUrl.Split('.')[audioUrl.Split('.').Length-1]; 
                                if(File.Exists(newPath))
                                {
                                    newPath = (newPath.Replace(newPath.Split('.')[newPath.Split('.').Length - 1], "")).Replace(".","") + Utility.RandomNumber(1, 100000, counta)
                                     +"."+ newPath.Split('.')[newPath.Split('.').Length - 1];
                                }
                                counta++;
                                File.Copy(audioUrl, newPath);
                            }
                        }
                    }

                    presentation.Close();
                    MessageBox.Show("Audios are copied to: "+ currentPPTPath + @"\Audios");

                }
                else
                {
                    MessageBox.Show("temp.pptx file not found on :"+ currentPPTPath + @"/temp.pptx");
                }
                

            }else
            {
                MessageBox.Show("Please Extract First");
            }

        }

        /// <summary>
        /// This method is to get total shape count of the presentation
        /// </summary>
        /// <param name="presentation"></param>
        private int getTotalShapeCount(Presentation presentation)
        {
            int TotalShapeCount=0;
            foreach (Slide slide in presentation.Slides)
            {
                TotalShapeCount = TotalShapeCount + slide.Shapes.Count;
            }
            return TotalShapeCount;
        }

        public bool isPresentationChanged()
        {
            //Some intelligence to detect presentation has been modified
            if (!(prevTotalShapeCount == getTotalShapeCount(Globals.ThisAddIn.Application.ActivePresentation) &&
                prevTotalSlideCount == Globals.ThisAddIn.Application.ActivePresentation.Slides.Count &&
                prevPresentationName == Globals.ThisAddIn.Application.ActivePresentation.Name))
            {
                return true;
            }
            else
            {
                return false;
            }
                
        }

        private void BtnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.updateSettings();
        }
        private void BtnPreviewWeb_Click(object sender, RibbonControlEventArgs e)
        {
            if (isPresentationChanged() == true)
            {
                MessageBox.Show("Presentation updated! Extarcting first to get updated output");
                extractSlide();
            }

            if (File.Exists(currentPPTPath + "\\Json.JSON"))
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
