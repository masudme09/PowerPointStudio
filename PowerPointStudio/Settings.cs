using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MessageBox=System.Windows.Forms.MessageBox;
using System.Xml;

namespace PowerPointStudio
{
    public static class Settings
    {
        public static int exportImageDpi { get; set; }
        public static ShapeExportOptions shapeExportOptions { get; set; }
        
        public static void updateSettings()
        {
            FrmSettings frmSettings = new FrmSettings();
            frmSettings.Show();
        }

        /// <summary>
        /// Write the current settings to XML in public douments folder
        /// </summary>
        internal static bool SaveSettings(int exportImageDpi, ShapeExportOptions shapeExportOptions)
        {
            try
            {
                string saveDirectory = @"C:\Users\Public\Documents\EzilmStudioSettings";
                Directory.CreateDirectory(saveDirectory);

                XmlDocument xmlDoc = new XmlDocument();
                XmlNode rootNode = xmlDoc.CreateElement("Settings");
                xmlDoc.AppendChild(rootNode);

                XmlNode exportDpi = xmlDoc.CreateElement("exportDpi");
                exportDpi.InnerText = exportImageDpi.ToString();
                rootNode.AppendChild(exportDpi);

                XmlNode shapeExportOption = xmlDoc.CreateElement("shapeExportOptions");
                shapeExportOption.InnerText = shapeExportOptions.ToString();
                rootNode.AppendChild(shapeExportOption);

                xmlDoc.Save(saveDirectory + @"\settings.xml");
                return true;
                
            }
            catch
            {
                return false;
            }

        }

        internal static bool ReadAndUpdateFromXML()
        {
            string xmlPath = @"C:\Users\Public\Documents\EzilmStudioSettings" + @"\settings.xml";
            if (File.Exists(xmlPath))
            {                
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(xmlPath);
                exportImageDpi = Convert.ToInt32(xmlDocument.GetElementsByTagName("exportDpi")[0].InnerText);
                shapeExportOptions = (ShapeExportOptions)Enum.Parse(typeof(ShapeExportOptions), xmlDocument.GetElementsByTagName("shapeExportOptions")[0].InnerText, true);
            }
            else
            {
                SaveSettings(72, ShapeExportOptions.OneShapeExportOnce);
            }
            return true;

        }


    }
}
