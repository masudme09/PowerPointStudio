using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointStudio
{
    public partial class FrmSettings : Form
    {
        public FrmSettings()
        {
            InitializeComponent();
            LoadPreviousSettingsToUI();
            dropDownExportDpi.SelectedIndexChanged += DropDownExportDpi_SelectedIndexChanged;
            dropDownExportOption.SelectedIndexChanged += DropDownExportOption_SelectedIndexChanged;
            
        }

        private void DropDownExportOption_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
        }

        private void DropDownExportDpi_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
        }

        private void BtnModify_Click(object sender, EventArgs e)
        {
            dropDownExportDpi.Enabled = true;
            dropDownExportOption.Enabled = true;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            ShapeExportOptions shapeExportOptions=ShapeExportOptions.OneShapeExportOnce;
            int exportDpi = Convert.ToInt32(dropDownExportDpi.SelectedItem.ToString());

            switch (dropDownExportOption.SelectedItem)
            {
                case "Same Export Once":
                    shapeExportOptions = ShapeExportOptions.OneShapeExportOnce;
                    break;
                case "Export Irrespective":
                    shapeExportOptions = ShapeExportOptions.ShapeExportIrrespective;
                    break;
                default:
                    break;
            }
            if (Settings.SaveSettings(exportDpi, shapeExportOptions))
            {
                MessageBox.Show("Settings Updated Successfully");
            }else
            {
                MessageBox.Show("Error!!! Settings Update Unsuccessful!");
            }

            Settings.ReadAndUpdateFromXML();
            this.Close();
        }

        private void LoadPreviousSettingsToUI()
        {
            int dpiSettings = Settings.exportImageDpi;
            ShapeExportOptions shapeExportOption = Settings.shapeExportOptions;
            dropDownExportDpi.SelectedItem = dpiSettings.ToString();

            switch (shapeExportOption)
            {
                case ShapeExportOptions.OneShapeExportOnce:
                    dropDownExportOption.SelectedItem = "Same Export Once";
                    break;
                case ShapeExportOptions.ShapeExportIrrespective:
                    dropDownExportOption.SelectedItem = "Export Irrespective";
                    break;
                default:
                    break;
            }
        }
    }
}
