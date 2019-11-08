using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using DataTable = System.Data.DataTable;

namespace PowerPointStudio
{
    public static class Utility
    {       
        /// <summary>
        /// Takes shape and return its type
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public static ezShapeType GetShapeType(Shape shape)
        {
            //Checking Group
            #region Group Object Checking
            int itemsCount;
            try
            {
                itemsCount = shape.GroupItems.Count;
            }
            catch
            {
                itemsCount = 0;
            }
            if (itemsCount > 0)
            {
                return ezShapeType.Group;
            }
            #endregion Group
            //Checking ellipse Callout
            if (shape.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeOvalCallout)
            {
                return ezShapeType.EllipseCallout;
            }

            if(shape.Type==Microsoft.Office.Core.MsoShapeType.msoTextBox || (shape.Type==Microsoft.Office.Core.MsoShapeType.msoPlaceholder && shape.TextFrame.TextRange.Text!=""))
            {
                return ezShapeType.TextBox;
            }

            //For others. Later we will go one by one. Now keep those in a common name 'other'
            return ezShapeType.Other;
        }

        /// <summary>
        /// This method extract the current presentation to get medias
        /// </summary>
        /// <param name="presentation"></param>
        /// <returns></returns>
        public static string createZipAndExtract(Presentation presentation)
        {
            //Copying the current presentation to 'Media' folder
            //Make Zip and extract then take the audio to another directory

            Directory.CreateDirectory(presentation.Path + "\\Media");
            string mediaPath = presentation.Path +
                "\\Media" + "\\" + presentation.Name.Replace("pptm", "zip");
            mediaPath = mediaPath.Replace("pptx", "zip");

            if (File.Exists(mediaPath))
            {
                File.Delete(mediaPath);
            }
            File.Copy(presentation.Path + "\\" + presentation.Name, mediaPath);


            if (Directory.Exists(Path.GetDirectoryName(mediaPath) + "\\Extract"))
            {
                Directory.Delete(Path.GetDirectoryName(mediaPath) + "\\Extract", true);
            }

            //Extracting to Extract Directory
            ZipFile.ExtractToDirectory(mediaPath, Path.GetDirectoryName(mediaPath) + "\\Extract");

            //Now copy the midea files to the 'Medias' folder from Extract\ppt\media and delete extract directory
            if (Directory.Exists(presentation.Path + "\\Medias"))
            {
                Directory.Delete(presentation.Path + "\\Medias", true);
            }
            while(!Directory.Exists(presentation.Path + "\\Medias"))
            {
                Directory.CreateDirectory(presentation.Path + "\\Medias");//Creating Medias directory
            }

            //Copy the slides directory to different location
            if (Directory.Exists(Path.GetDirectoryName(mediaPath) + "\\Extract" + @"\ppt\slides\"))
            {

                Copy(Path.GetDirectoryName(mediaPath) + "\\Extract" + @"\ppt\slides\", presentation.Path + "\\Medias\\Slides");
                 
            }


            //COpying the audio or media file to different location
            if (Directory.Exists(Path.GetDirectoryName(mediaPath) + "\\Extract" + @"\ppt\media\"))
            {
                foreach (string file in Directory.GetFiles(Path.GetDirectoryName(mediaPath) + "\\Extract" + @"\ppt\media\"))
                {
                    string fileName = file.Replace(Path.GetDirectoryName(mediaPath) + "\\Extract" + @"\ppt\media\", "");
                    if (fileName.Contains("media") || fileName.Contains("audio"))
                    {
                        File.Copy(file, presentation.Path + "\\Medias\\" + fileName);
                    }

                }
            }

            //Deleting the media directory with all contents or extraction directory
            if (Directory.Exists(presentation.Path + "\\Media"))
            {
                Directory.Delete(presentation.Path + "\\Media", true);
            }

            return presentation.Path + "\\Medias";
        }

        /// <summary>
        /// Copy all files and folders from a directory
        /// </summary>
        /// <param name="sourceDirectory"></param>
        /// <param name="targetDirectory"></param>
        public static void Copy(string sourceDirectory, string targetDirectory)
        {
            DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
            DirectoryInfo diTarget = new DirectoryInfo(targetDirectory);

            CopyAll(diSource, diTarget);
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                    target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }



        /// <summary>
        /// Return url of the media that have same audio id
        /// </summary>
        /// <param name="audioId"></param>
        /// <returns></returns>
        public static string getExtractedAudioUrl(Shape mediaShape) //Need to use from open xml
        {
            string mediaDirectory = PowerPointStudioRibbon.mediaPath;
            string mediaShapeName = mediaShape.Name;

            //Read from xml to know which media file its is targeting
            int slideIndex = mediaShape.Parent.SlideIndex;
            string r_link="";
            XmlDocument doc = new XmlDocument();
            doc.Load(mediaDirectory+ "\\Slides\\slide"+slideIndex+".xml");

            XmlNodeList elemList = doc.GetElementsByTagName("p:pic");
            foreach (XmlNode elem in elemList)
            {
                if(elem.ChildNodes[0].ChildNodes[0].Attributes["name"].Value==mediaShapeName)
                {
                    if (elem.InnerXml.Contains("audioFile"))
                    {
                        r_link = elem.ChildNodes[0].ChildNodes[2].ChildNodes[0].Attributes["r:link"].Value;
                        break;
                    }
                    else
                    {
                        r_link = "";
                    }
                }
                
            }

            string mediaName="";
            doc.Load(mediaDirectory + "\\Slides\\_rels\\slide" + slideIndex + ".xml.rels");
            XmlNodeList elems = doc.GetElementsByTagName("Relationship");
            foreach (XmlNode elem in elems)
            {
                if(elem.Attributes["Id"].Value==r_link)
                {
                    mediaName = (elem.Attributes["Target"].Value).Split('/')[(elem.Attributes["Target"].Value).Split('/').Length-1];
                }
            }
            return mediaDirectory+@"\"+mediaName;

        }


        public static float SlideWidthGet(Presentation presentation)
        {
            PageSetup dimensions = presentation.PageSetup;
            return dimensions.SlideWidth;
        }

        public static void SlideWidthSet(Presentation presentation, float value)
        {
            presentation.PageSetup.SlideWidth = value;
        }

        public static float SlideHeightGet(Presentation presentation)
        {

            PageSetup dimensions = presentation.PageSetup;
            return dimensions.SlideHeight;

        }
        public static void SlideHeightSet(Presentation presentation, float value)
        {
            presentation.PageSetup.SlideHeight = value;
        }


        //COnverts image close to 300dpi
        public static void CustomDpi(Bitmap original, int new_wid, int new_hgt,int dpi, string savingPathWithExtension)
        {
            Bitmap returnBmp;
            using (Graphics gr = Graphics.FromImage(original))
            {
                float dpiX = gr.DpiX;
                float dpiY = gr.DpiY;
                //gr.Dispose();
            }
            float originalWidth = original.Width;
            float originalHeight = original.Height;


            using (Bitmap bm = new Bitmap(new_wid, new_hgt))
            {
                System.Drawing.Point[] points =
                {
                new System.Drawing.Point(0, 0),
                new System.Drawing.Point(new_wid, 0),
                new System.Drawing.Point(0, new_hgt),
            };
                using (Graphics gr = Graphics.FromImage(bm))
                {
                    gr.DrawImage(original, points);
                    //gr.Dispose();
                }
                float dpix = dpi;
                float dpiy = dpi;
                bm.SetResolution(dpix, dpiy);
                returnBmp = bm;
                if(File.Exists(savingPathWithExtension))
                {
                    original.Dispose();
                    File.Delete(savingPathWithExtension);
                }
                bm.Save(savingPathWithExtension);
                bm.Dispose();
            }

        }

        // Generate a random number between two numbers in a string format
        public static string RandomNumber(int min, int max, int shapeCount)
        {
            Random random = new Random();
            string rand = random.Next(min, max) +"_"+ String.Format("{0:0.00}",
            shapeCount.ToString());

            return rand;
        }

        /// <summary>
        /// Create serialize JSON indented and null removed
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string createJson(object obj)
        {
            var settings = new JsonSerializerSettings()
            {
                //ContractResolver = new OrderedContractResolver(),
                NullValueHandling = NullValueHandling.Ignore
            };

            var json = JsonConvert.SerializeObject(obj, Newtonsoft.Json.Formatting.Indented, settings);

            return json;
        }

        /// <summary>
        /// Create JSON and write that to a specified path
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="filePathWithExtension"></param>
        public static void writeJsonToFile(object obj, string filePathWithExtension)
        {
            string json = createJson(obj);
            File.WriteAllText(filePathWithExtension, json);
        }

        public static string qulifiedNameGenerator(string unqualifiedName)
        {
            string qualifiedName = Regex.Replace(unqualifiedName, @"\t|\n|\r|-|/|\\|:|\?|\*|<|>|""|", "");
            qualifiedName = qualifiedName.Replace(" ", "");
            return qualifiedName;
        }
        /// <summary>
        /// Send shape ezLang here and it will find out text on on corresponding tag
        /// </summary>
        /// <param name="altText"></param>
        /// <param name="searchedTag"></param>
        /// <returns></returns>
        public static string ezLangFinder(string altText, string searchedTag) 
        {
            string text = null;

            text = altText.Substring(altText.IndexOf("$"+searchedTag+"$") + (searchedTag.Length+2), (altText.IndexOf("$$"+searchedTag+"$$") - (altText.IndexOf("$" + searchedTag + "$") + (searchedTag.Length + 2)))); //It is returning first character index of the searched string

            return text;
        }
        /// <summary>
        /// Export Datatable to CSV 
        /// </summary>
        /// <param name="dtDataTable"></param>
        /// <param name="strFilePath"></param>
        public static void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (System.Data.DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        /// <summary>
        /// This method will clear all static resource that is not required after extraction is complete
        /// </summary>
        public static void staticResourceClear()
        {
           
        }

    }

    public class OrderedContractResolver : DefaultContractResolver
    {
        protected override System.Collections.Generic.IList<JsonProperty> CreateProperties(System.Type type, MemberSerialization memberSerialization)
        {
            return base.CreateProperties(type, memberSerialization).OrderBy(p => p.PropertyName).ToList();
        }
    }
}
