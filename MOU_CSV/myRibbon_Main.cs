using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using System.Reflection;
using System.IO;

namespace MOU_CSV
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class myRibbon_Main : IExternalApplication
    {
        public static string image_path = @"C:\Program Files (x86)\IIC\WSRevitAddin\Resources";
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {

                cls_Bien.GetRevitLanguage(application);
                // Create a custom ribbon tab
                string tabName = "WS EXPORT";
                application.CreateRibbonTab(tabName);

                // Create a ribbon panel
                RibbonPanel panel = application.CreateRibbonPanel(tabName, SimpleTranslator.T("Export"));

                //// Add a button to the panel
                PushButtonData buttonData = createButton("btnMOU_CSV", "CSV Export", image_path + "\\import_xml.png", " ", "", "cls_RnW");
                PushButton pushButton = panel.AddItem(buttonData) as PushButton;

                // Optionally set button properties
                pushButton.ToolTip = "Execute MOU_CSV Command";

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show(SimpleTranslator.T("Error"), ex.Message);
                return Result.Failed;
            }
        }

        public PushButtonData createButton(string btnName, string btnText, string path_image, string text_Tooltip, string url_Help, string cls_name)
        {
            //Tạo ribbon button
            string path = Assembly.GetExecutingAssembly().Location;
            PushButtonData button = new PushButtonData(btnName, btnText, path, "MOU_CSV." + cls_name);
            //addTooltip
            button.ToolTip = text_Tooltip;

            //add Help
            //Set icon cho button
            button.LargeImage = LoadImage(path_image);
            return button;
        }
        public System.Windows.Media.Imaging.BitmapImage LoadImage(string uri)
        {
            // Create a new BitmapImage from the provided URI
            // 指定された URI から新しい BitmapImage を作成します
            System.Windows.Media.Imaging.BitmapImage _bi = new System.Windows.Media.Imaging.BitmapImage(new Uri(uri));

            // Return the BitmapImage
            // BitmapImage を返します
            return _bi;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private System.Windows.Media.ImageSource LoadPng(string iconFileName)
        {
            var uri = new Uri($"pack://application:,,,/MOU_CSV;component/Resource/{iconFileName}");
            return new System.Windows.Media.Imaging.PngBitmapDecoder(uri,
                System.Windows.Media.Imaging.BitmapCreateOptions.PreservePixelFormat,
                System.Windows.Media.Imaging.BitmapCacheOption.OnLoad).Frames[0];
        }



    }
}