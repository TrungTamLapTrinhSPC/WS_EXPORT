using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace MOU_CSV
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cls_RnW : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            //Load dictionary trước khi dùng
            SimpleTranslator.dicTranslations = SimpleTranslator.LoadEngToJapDictionary();
            SimpleTranslator.NgonNgu = cls_Bien.NGONNGU; // Đảm bảo ngôn ngữ được set

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            cls_Bien.uidoc = uidoc;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            cls_Bien.doc = doc;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            // Use HashSet to store IDs to automatically remove duplicates
            var elementIds = new HashSet<string>();

            var collector = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType()
                .Where(e =>
                {
                    // 1. Check category
                    if (e.Category == null)
                        return false;

                    // 2. Only take Model categories
                    if (e.Category.CategoryType != CategoryType.Model)
                        return false;

                    // 3. Exclude unwanted categories (HashSet O(1) lookup)
                    var builtInCat = (BuiltInCategory)e.Category.Id.IntegerValue;
                    if (ExcludedCategoriesForStatistics.Contains(builtInCat))
                        return false;

                    // 4. Exclude nested components if it's a FamilyInstance
                    if (e is FamilyInstance fi && fi.SuperComponent != null)
                        return false;

                    return true;
                })
                .ToList();

            // Add element IDs to HashSet (removes duplicates)
            foreach (var el in collector)
            {
                elementIds.Add(el.Id.IntegerValue.ToString());
            }

            // Or use a shorter way:
            // elementIds.UnionWith(collector.Select(el => el.Id.IntegerValue.ToString()));

            // Convert HashSet to List if cls_Bien.lstId requires a List
            cls_Bien.lstId = elementIds.ToList();

            // Lấy thông tin cơ bản
            List<Dictionary<string, string>> basicInfo;
            GetInformation(out basicInfo);

            // Lấy Identity Data
            List<Dictionary<string, string>> identityInfo;
            GetIdentityData(out identityInfo);

            //Lấy Constraints Data
            List<Dictionary<string, string>> constraintsInfo;
            GetConstraints(out constraintsInfo);

            //Lấy Dimensions Data
            List<Dictionary<string, string>> dimensionsInfo;
            GetDimensions(out dimensionsInfo);

            cls_Bien._basicInfo = basicInfo;
            cls_Bien._identityInfo = identityInfo;
            cls_Bien._constraintsInfo = constraintsInfo;
            cls_Bien._dimensionsInfo = constraintsInfo;

            // Mở form quản lý
            Form_ManagerObject form_ManagerObject = new Form_ManagerObject(doc, uidoc);
            form_ManagerObject.ShowDialog(); // Dùng ShowDialog thay vì Show để đợi form đóng

            // Sau khi form đóng, kiểm tra xem có dữ liệu cần update không
            if (cls_Bien.dataImport != null && cls_Bien.dataImport.Count > 0)
            {
                // Thực hiện update
                int successCount;
                List<string> failedElementIds;

                Result updateResult = ApplyChangesFromImport(doc, out successCount, out failedElementIds);

                if (updateResult == Result.Succeeded)
                {
                    string resultMessage = $"Successfully updated {successCount} objects!";

                    if (failedElementIds.Count > 0)
                    {
                        resultMessage += $"\n\n{failedElementIds.Count} objects failed to update.";
                    }

                    TaskDialog.Show("Update Complete", resultMessage);

                    // Highlight các objects đã thay đổi thành công
                    HighlightChangedElements(uidoc, cls_Bien.changedElementIds, failedElementIds);
                }
                else
                {
                    TaskDialog.Show("Error", "Failed to apply changes.");
                }
            }

            return Result.Succeeded;
        }

        private void GetInformation(out List<Dictionary<string, string>> dataList)
        {
            UIDocument uidoc = cls_Bien.uidoc;
            Autodesk.Revit.DB.Document doc = cls_Bien.doc;

            dataList = new List<Dictionary<string, string>>();

            foreach (string idStr in cls_Bien.lstId)
            {
                ElementId elId = new ElementId(int.Parse(idStr));
                Element el = doc.GetElement(elId);
                if (el == null) continue;

                Dictionary<string, string> dataRow = new Dictionary<string, string>();

                // Level
                string levelName = "N/A";
                string[] allowedLevelParams = new string[]
                {
            T("Reference Level"),
            T("Base Level"),
            T("Schedule Level"),
            T("Level"),
            T("Host Level"),
            T("Work Plane Level")
                };

                try
                {
                    foreach (string paramName in allowedLevelParams)
                    {
                        Parameter p = el.LookupParameter(paramName);
                        if (p != null && p.HasValue)
                        {
                            levelName = p.AsValueString();
                            break;
                        }
                    }
                    Level lvl = doc.GetElement(el.LevelId) as Level;
                    if (lvl != null)
                        levelName = lvl.Name;
                }
                catch
                {
                    levelName = T("N/A");
                }

                dataRow[T("Level")] = levelName;

                // Object category
                string categoryName = el.Category != null ? el.Category.Name : T("NoCategory");
                dataRow[T("Category")] = categoryName;

                // Original Revit ID
                dataRow[T("Revit ID")] = elId.IntegerValue.ToString();

                // Type Name
                string typeName = el.Name;
                if (el is FamilyInstance fi && fi.Symbol != null)
                    typeName = fi.Symbol.FamilyName + " - " + fi.Symbol.Name;
                dataRow[T("Type Name")] = typeName;

                dataList.Add(dataRow);
            }
        }

        private void GetIdentityData(out List<Dictionary<string, string>> dataList)
        {
            Autodesk.Revit.DB.Document doc = cls_Bien.doc;

            dataList = new List<Dictionary<string, string>>();

            foreach (string idStr in cls_Bien.lstId)
            {
                ElementId elId = new ElementId(int.Parse(idStr));
                Element el = doc.GetElement(elId);
                if (el == null) continue;

                Dictionary<string, string> dataRow = new Dictionary<string, string>();
                List<Parameter> allParams = new List<Parameter>();

                // Get Identity Data from instance
                try
                {
                    foreach (Parameter p in el.Parameters)
                    {
                        if (p.Definition != null &&
                            p.Definition.ParameterGroup == BuiltInParameterGroup.PG_IDENTITY_DATA)
                        {
                            allParams.Add(p);
                        }
                    }
                }
                catch { }

                // Get Identity Data from type element
                try
                {
                    ElementId typeId = el.GetTypeId();
                    Element typeEl = doc.GetElement(typeId);

                    if (typeEl != null)
                    {
                        foreach (Parameter p in typeEl.Parameters)
                        {
                            if (p.Definition != null &&
                                p.Definition.ParameterGroup == BuiltInParameterGroup.PG_IDENTITY_DATA)
                            {
                                // Only add if not exists (instance overrides type)
                                if (!allParams.Any(existing => existing.Definition.Name == p.Definition.Name))
                                {
                                    allParams.Add(p);
                                }
                            }
                        }
                    }
                }
                catch { }

                // Sắp xếp theo Definition.Id để giữ thứ tự như Revit
                allParams = allParams.OrderBy(p => p.Id.IntegerValue).ToList();

                // Thêm vào dataRow theo thứ tự đã sắp xếp
                foreach (var p in allParams)
                {
                    string paramName = p.Definition.Name;
                    string paramValue = p.HasValue ? p.AsValueString() : " ";

                    if (!dataRow.ContainsKey(paramName))
                        dataRow[paramName] = paramValue;
                }

                dataList.Add(dataRow);
            }
        }

        private void GetConstraints(out List<Dictionary<string, string>> dataList)
        {
            try
            {
                Autodesk.Revit.DB.Document doc = cls_Bien.doc;

                dataList = new List<Dictionary<string, string>>();

                foreach (string idStr in cls_Bien.lstId)
                {
                    ElementId elId = new ElementId(int.Parse(idStr));
                    Element el = doc.GetElement(elId);
                    if (el == null) continue;

                    Dictionary<string, string> dataRow = new Dictionary<string, string>();
                    List<Parameter> allParams = new List<Parameter>();

                    // Get Constraints from instance
                    try
                    {
                        foreach (Parameter p in el.Parameters)
                        {
                            if (p.Definition != null &&
                                p.Definition.ParameterGroup == BuiltInParameterGroup.PG_CONSTRAINTS)
                            {
                                allParams.Add(p);
                            }
                        }
                    }
                    catch { }

                    // Get Constraints from type element
                    try
                    {
                        ElementId typeId = el.GetTypeId();
                        Element typeEl = doc.GetElement(typeId);

                        if (typeEl != null)
                        {
                            foreach (Parameter p in typeEl.Parameters)
                            {
                                if (p.Definition != null &&
                                    p.Definition.ParameterGroup == BuiltInParameterGroup.PG_CONSTRAINTS)
                                {
                                    // Only add if not exists (instance overrides type)
                                    if (!allParams.Any(existing => existing.Definition.Name == p.Definition.Name))
                                    {
                                        allParams.Add(p);
                                    }
                                }
                            }
                        }
                    }
                    catch { }

                    // Sắp xếp theo Definition.Id để giữ thứ tự như Revit
                    allParams = allParams.OrderBy(p => p.Id.IntegerValue).ToList();

                    // Thêm vào dataRow theo thứ tự đã sắp xếp
                    foreach (var p in allParams)
                    {
                        string paramName = p.Definition.Name;
                        string paramValue = p.HasValue ? p.AsValueString() : " ";

                        if (!dataRow.ContainsKey(paramName))
                            dataRow[paramName] = paramValue;
                    }

                    dataList.Add(dataRow);
                }
            }
            catch
            {
                dataList = new List<Dictionary<string, string>>();
            }
        }

        private void GetDimensions(out List<Dictionary<string, string>> dataList)
        {
            try
            {
                Autodesk.Revit.DB.Document doc = cls_Bien.doc;

                dataList = new List<Dictionary<string, string>>();

                foreach (string idStr in cls_Bien.lstId)
                {
                    ElementId elId = new ElementId(int.Parse(idStr));
                    Element el = doc.GetElement(elId);
                    if (el == null) continue;

                    Dictionary<string, string> dataRow = new Dictionary<string, string>();
                    List<Parameter> allParams = new List<Parameter>();

                    // Get Dimensions from instance
                    try
                    {
                        foreach (Parameter p in el.Parameters)
                        {
                            if (p.Definition != null &&
                                p.Definition.ParameterGroup == BuiltInParameterGroup.PG_GEOMETRY)
                            {
                                allParams.Add(p);
                            }
                        }
                    }
                    catch { }

                    // Get Dimensions from type element
                    try
                    {
                        ElementId typeId = el.GetTypeId();
                        Element typeEl = doc.GetElement(typeId);

                        if (typeEl != null)
                        {
                            foreach (Parameter p in typeEl.Parameters)
                            {
                                if (p.Definition != null &&
                                    p.Definition.ParameterGroup == BuiltInParameterGroup.PG_GEOMETRY)
                                {
                                    // Only add if not exists (instance overrides type)
                                    if (!allParams.Any(existing => existing.Definition.Name == p.Definition.Name))
                                    {
                                        allParams.Add(p);
                                    }
                                }
                            }
                        }
                    }
                    catch { }

                    // Sắp xếp theo Definition.Id để giữ thứ tự như Revit
                    allParams = allParams.OrderBy(p => p.Id.IntegerValue).ToList();

                    // Thêm vào dataRow theo thứ tự đã sắp xếp
                    foreach (var p in allParams)
                    {
                        string paramName = p.Definition.Name;
                        string paramValue = p.HasValue ? p.AsValueString() : " ";

                        if (!dataRow.ContainsKey(paramName))
                            dataRow[paramName] = paramValue;
                    }

                    dataList.Add(dataRow);
                }
            }
            catch
            {
                dataList = new List<Dictionary<string, string>>();
            }
        }

        private class GroupInfo
        {
            public Dictionary<string, string> DataRow { get; set; }
            public int Count { get; set; }
            public List<string> RevitIds { get; set; }
        }

        // List of categories to exclude from statistics
        private static readonly HashSet<BuiltInCategory> ExcludedCategoriesForStatistics = new HashSet<BuiltInCategory>
        {
            // Annotation
            BuiltInCategory.OST_Dimensions,
            BuiltInCategory.OST_TextNotes,
            BuiltInCategory.OST_GenericAnnotation,
            BuiltInCategory.OST_Tags,
            BuiltInCategory.OST_RoomTags,
            BuiltInCategory.OST_DoorTags,
            BuiltInCategory.OST_WindowTags,
            BuiltInCategory.OST_WallTags,
            BuiltInCategory.OST_MultiCategoryTags,
            BuiltInCategory.OST_StairsTags,
            BuiltInCategory.OST_AreaTags,
            BuiltInCategory.OST_SpotElevations,
            BuiltInCategory.OST_SpotSlopes,
            BuiltInCategory.OST_SpotCoordinates,
            BuiltInCategory.OST_RevisionClouds,
            BuiltInCategory.OST_Callouts,
            BuiltInCategory.OST_ElevationMarks,
            BuiltInCategory.OST_Sections,

            // Reference/Datum
            BuiltInCategory.OST_Grids,
            BuiltInCategory.OST_Levels,
            BuiltInCategory.OST_CLines,
            BuiltInCategory.OST_Lines,
            BuiltInCategory.OST_SketchLines,
            BuiltInCategory.OST_SectionBox,
            BuiltInCategory.OST_Matchline,
            BuiltInCategory.OST_Constraints,

            // Analytical
            BuiltInCategory.OST_AnalyticalNodes,
            BuiltInCategory.OST_LinksAnalytical,
            BuiltInCategory.OST_ColumnAnalytical,
            BuiltInCategory.OST_FramingAnalyticalGeometry,
            BuiltInCategory.OST_FoundationSlabAnalytical,
            BuiltInCategory.OST_LoadCases,
            BuiltInCategory.OST_Loads,

            // Spatial
            BuiltInCategory.OST_Rooms,
            BuiltInCategory.OST_MEPSpaces,
            BuiltInCategory.OST_Areas,
            BuiltInCategory.OST_AreaSchemes,

            // Systems
            BuiltInCategory.OST_PipingSystem,
            BuiltInCategory.OST_DuctSystem,
            BuiltInCategory.OST_ElectricalCircuit,

            // Views
            BuiltInCategory.OST_Views,
            BuiltInCategory.OST_Sheets,
            BuiltInCategory.OST_Viewports,
            BuiltInCategory.OST_Cameras,
            BuiltInCategory.OST_RenderRegions,
            BuiltInCategory.OST_Viewers,

            // Detail/2D
            BuiltInCategory.OST_DetailComponents,
            BuiltInCategory.OST_DetailComponentTags,
            BuiltInCategory.OST_FilledRegion,
            BuiltInCategory.OST_MaskingRegion,
            BuiltInCategory.OST_InsulationLines,

            // Import 
            BuiltInCategory.OST_ImportObjectStyles,
            BuiltInCategory.OST_RvtLinks,

            // Others
            BuiltInCategory.OST_Topography,
            BuiltInCategory.OST_CurtainGrids,
            BuiltInCategory.OST_CurtainGridsWall,
            BuiltInCategory.OST_CurtainGridsSystem,
        };

        private string T(string key)
        {
            return SimpleTranslator.T(key);
        }
        /// <summary>
        /// Hàm thực hiện update từ dữ liệu import
        /// </summary>
        private Result ApplyChangesFromImport(Autodesk.Revit.DB.Document doc, out int successCount, out List<string> failedElementIds)
        {
            successCount = 0;
            failedElementIds = new List<string>();

            using (Transaction trans = new Transaction(doc, "Apply Parameter Changes"))
            {
                trans.Start();

                try
                {
                    foreach (var row in cls_Bien.dataImport)
                    {
                        if (!row.ContainsKey("Revit ID"))
                            continue;

                        string revitId = row["Revit ID"];
                        bool elementSuccess = false;

                        try
                        {
                            if (!int.TryParse(revitId, out int elementIdInt))
                            {
                                failedElementIds.Add(revitId);
                                continue;
                            }

                            ElementId elementId = new ElementId(elementIdInt);
                            Element element = doc.GetElement(elementId);

                            if (element == null)
                            {
                                failedElementIds.Add(revitId);
                                continue;
                            }

                            foreach (var kvp in row)
                            {
                                string paramName = kvp.Key;
                                string newValue = kvp.Value;

                                if (paramName == "Revit ID" || paramName == "Category")
                                    continue;

                                Parameter param = element.LookupParameter(paramName);

                                if (param == null || param.IsReadOnly)
                                    continue;

                                // Skip type parameters
                                ElementType elementType = doc.GetElement(element.GetTypeId()) as ElementType;
                                if (elementType != null)
                                {
                                    Parameter typeParam = elementType.LookupParameter(paramName);
                                    if (typeParam != null)
                                        continue;
                                }

                                try
                                {
                                    switch (param.StorageType)
                                    {
                                        case StorageType.String:
                                            param.Set(newValue);
                                            elementSuccess = true;
                                            break;

                                        case StorageType.Integer:
                                            if (int.TryParse(newValue, out int intValue))
                                            {
                                                param.Set(intValue);
                                                elementSuccess = true;
                                            }
                                            break;

                                        case StorageType.Double:
                                            if (double.TryParse(newValue, out double doubleValue))
                                            {
                                                // Kiểm tra loại đơn vị và convert nếu cần
                                                ForgeTypeId specType = param.Definition.GetDataType();

                                                if (specType == SpecTypeId.Length ||
                                                    specType == SpecTypeId.Distance ||
                                                    specType == SpecTypeId.SectionDimension ||
                                                    specType == SpecTypeId.CableTraySize ||
                                                    specType == SpecTypeId.ConduitSize ||
                                                    specType == SpecTypeId.DuctSize ||
                                                    specType == SpecTypeId.PipeDimension ||
                                                    specType == SpecTypeId.PipeSize ||
                                                    specType == SpecTypeId.ReinforcementLength ||
                                                    specType == SpecTypeId.Displacement)
                                                {
                                                    // Convert từ mm sang feet
                                                    doubleValue = doubleValue / 304.8;
                                                }
                                                else if (specType == SpecTypeId.Area)
                                                {
                                                    // Convert từ m² sang feet²
                                                    doubleValue = doubleValue / 0.09290304;
                                                }
                                                else if (specType == SpecTypeId.Volume)
                                                {
                                                    // Convert từ m³ sang feet³
                                                    doubleValue = doubleValue / 0.028316847;
                                                }
                                                else if (specType == SpecTypeId.Angle)
                                                {
                                                    // Convert từ độ sang radian
                                                    doubleValue = doubleValue * Math.PI / 180.0;
                                                }

                                                param.Set(doubleValue);
                                                elementSuccess = true;
                                            }
                                            break;

                                        case StorageType.ElementId:
                                            if (int.TryParse(newValue, out int elemIdInt))
                                            {
                                                param.Set(new ElementId(elemIdInt));
                                                elementSuccess = true;
                                            }
                                            break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"Error updating {paramName} for {revitId}: {ex.Message}");
                                }
                            }

                            if (elementSuccess)
                            {
                                successCount++;
                            }
                            else
                            {
                                failedElementIds.Add(revitId);
                            }
                        }
                        catch (Exception ex)
                        {
                            failedElementIds.Add(revitId);
                            Debug.WriteLine($"Error updating element {revitId}: {ex.Message}");
                        }
                    }

                    trans.Commit();
                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    Debug.WriteLine($"Transaction error: {ex.Message}");
                    return Result.Failed;
                }
            }
        }


        /// <summary>
        /// Highlight các elements đã thay đổi thành công trong Revit
        /// </summary>
        private void HighlightChangedElements(UIDocument uidoc, List<string> changedIds, List<string> failedIds)
        {
            if (changedIds == null || changedIds.Count == 0)
                return;

            // Lọc ra những ID thành công (không có trong failed list)
            var successIds = changedIds.Except(failedIds).ToList();

            if (successIds.Count == 0)
                return;

            // Convert sang ElementId
            List<ElementId> elementIds = new List<ElementId>();
            foreach (string idStr in successIds)
            {
                if (int.TryParse(idStr, out int idInt))
                {
                    elementIds.Add(new ElementId(idInt));
                }
            }

            if (elementIds.Count > 0)
            {
                uidoc.Selection.SetElementIds(elementIds);
                uidoc.ShowElements(elementIds);
            }
        }
    }

}


