using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;

namespace MOU_CSV
{
    public static class cls_Bien
    {
        public static List<string> lstId = new List<string>();
        public static Autodesk.Revit.UI.UIDocument uidoc;
        public static Autodesk.Revit.DB.Document doc;
        public static Autodesk.Revit.UI.ExternalCommandData commandData;
        public static Autodesk.Revit.ApplicationServices.Application app;
        public static Autodesk.Revit.UI.UIApplication uiApp;
        public static List<Dictionary<string, string>> data;
        public static List<Dictionary<string, string>> dataImport;

        public static List<Dictionary<string, string>> _basicInfo;
        public static List<Dictionary<string, string>> _identityInfo;
        public static List<Dictionary<string, string>> _constraintsInfo;
        public static List<Dictionary<string, string>> _dimensionsInfo;

        // List các element ID đã thay đổi
        public static List<string> changedElementIds;


        public static string NGONNGU = "ENG";

        public static void GetRevitLanguage(UIControlledApplication application)
        {
            try
            {
                LanguageType language = application.ControlledApplication.Language;
                NGONNGU = (language == LanguageType.Japanese) ? "JAP" : "ENG";
            }
            catch
            {
                NGONNGU = "ENG";
            }
        }
    }
}
