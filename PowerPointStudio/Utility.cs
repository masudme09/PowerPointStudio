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
using System.Threading.Tasks;

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
            Directory.CreateDirectory(presentation.Path + "\\Medias");//Creating Medias directory
            foreach(string file in Directory.GetFiles(Path.GetDirectoryName(mediaPath) + "\\Extract" + @"\ppt\media\"))
            {
                string fileName = file.Replace(Path.GetDirectoryName(mediaPath) + "\\Extract" + @"\ppt\media\", "");
                if (fileName.Contains("media")|| fileName.Contains("audio"))
                {
                    File.Copy(file, presentation.Path + "\\Medias\\" + fileName);
                }
               
            }

            //Deleting the media directory with all contents
            if (Directory.Exists(presentation.Path + "\\Media"))
            {
                Directory.Delete(presentation.Path + "\\Media", true);
            }

            return presentation.Path + "\\Medias";
        }


        /// <summary>
        /// Return url of the media that have same audio id
        /// </summary>
        /// <param name="audioId"></param>
        /// <returns></returns>
        public static string getExtractedAudioUrl(int audioId)
        {
            string mediaDirectory = PowerPointStudioRibbon.mediaPath;
           
            foreach(string file in Directory.GetFiles(mediaDirectory))
            {
                string fileName = file.Replace(mediaDirectory, "");
                if(fileName.Contains("audio"+audioId)||fileName.Contains("media"+audioId))
                {
                    return file;
                }
            }
            return null;
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
        public static void CustomDpi(Bitmap original, int new_wid, int new_hgt, string savingPathWithExtension)
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
                float dpix = 300;
                float dpiy = 300;
                bm.SetResolution(dpix, dpiy);
                returnBmp = bm;
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

            var json = JsonConvert.SerializeObject(obj, Formatting.Indented, settings);

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


    }

    public class OrderedContractResolver : DefaultContractResolver
    {
        protected override System.Collections.Generic.IList<JsonProperty> CreateProperties(System.Type type, MemberSerialization memberSerialization)
        {
            return base.CreateProperties(type, memberSerialization).OrderBy(p => p.PropertyName).ToList();
        }
    }
}
