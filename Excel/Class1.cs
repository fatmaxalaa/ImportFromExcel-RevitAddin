using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace Excel
{
    [Transaction (TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            string fileName;

            OpenFileDialog openDlg = new OpenFileDialog(); //allow user to browse a new file and open it in revit

            openDlg.Title = "Select a file";
         
            DialogResult result = openDlg.ShowDialog();


            if (result == DialogResult.OK)  //if user click ok 
            {
                fileName = openDlg.FileName;   //selected file name will store in filename variable
                StreamReader sr = new StreamReader(fileName);  //read the file



                #region sheets
                ElementId titleBlocktytypeId = new FilteredElementCollector(doc)
                                              .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                              .WhereElementIsElementType()
                                              .First().Id;

                string csvLine;

                Transaction t = new Transaction(doc, "Create Sheets");
                t.Start();
                while ((csvLine = sr.ReadLine()) != null)
                {
                    char[] separator = new char[] { ',' }; 
                    string[] values = csvLine.Split(separator, StringSplitOptions.None);


                    if (values[0] != null && values[1] != null && values[2] != null)
                    {
                        ViewSheet sheet = ViewSheet.Create(doc, titleBlocktytypeId);
                        sheet.SheetNumber = values[0];
                        sheet.Name = values[1];
                        


                    }
                }
                t.Commit();
                #endregion



                #region floor plan
                //Level level = new FilteredElementCollector(doc)
                //            .OfCategory(BuiltInCategory.OST_Levels)
                //            .WhereElementIsNotElementType()
                //            .Cast<Level>()
                //            .First(e => e.Name == "Level 1");

                //ViewFamilyType viewFamilyType = new FilteredElementCollector(doc)
                //            .OfClass(typeof(ViewFamilyType))
                //            .WhereElementIsElementType()
                //            .Cast<ViewFamilyType>()
                //            .First(e => e.ViewFamily == ViewFamily.FloorPlan);


                //ViewPlan viewPlan = null;
                //string csvLine;

                //using (Transaction t = new Transaction(doc, "Create Floor plan"))
                //{

                //    t.Start();
                //    while ((csvLine = sr.ReadLine()) != null)
                //    {
                //        char[] separator = new char[] { ',' }; //to separete between (number and name and view columns)
                //        string[] values = csvLine.Split(separator, StringSplitOptions.None);


                //        if (values[0] != null && values[1] != null)
                //        {
                //            /*values = values.Skip(1).ToArray();*/

                //            viewPlan = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                //            viewPlan.Name = values[0];

                //        }
                //    }
                //    t.Commit();

                //}
                #endregion



                #region level



                //using (Transaction t = new Transaction(doc, "Create Floor plan"))
                //{
                //    Level level = null;
                //    string csvLine;
                //    t.Start();
                //    while ((csvLine = sr.ReadLine()) != null)
                //    {
                //        char[] separator = new char[] { ',' }; //to separete between columns
                //        string[] values = csvLine.Split(separator, StringSplitOptions.None);


                //        if (values[0] != null && values[1] != null)
                //        {
                //            //values = values.Skip(1).ToArray();
                //            double elevation = double.Parse(values[1]);

                //            level = Level.Create(doc, elevation);
                //            level.Name = values[0];

                //        }
                //    }
                //    t.Commit();

                //}

                #endregion



                #region Reflected Ceiling Plan
                //Level level = new FilteredElementCollector(doc)
                //            .OfCategory(BuiltInCategory.OST_Levels)
                //            .WhereElementIsNotElementType()
                //            .Cast<Level>()
                //            .First(e => e.Name == "Level 1");

                // ViewFamilyType viewFamilyType = new FilteredElementCollector(doc)
                //             .OfClass(typeof(ViewFamilyType))
                //             .WhereElementIsElementType()
                //             .Cast<ViewFamilyType>()
                //             .First(e => e.ViewFamily == ViewFamily.CeilingPlan);


                // ViewPlan viewPlan = null;
                // string csvLine;

                // using (Transaction t = new Transaction(doc, "Create Ceiling plan"))
                // {

                //     t.Start();
                //     while ((csvLine = sr.ReadLine()) != null)
                //     {
                //         char[] separator = new char[] { ',' }; //to separete between
                //         string[] values = csvLine.Split(separator, StringSplitOptions.None);


                //             if (values[0] != null && values[1] != null)
                //             {
                //                 //values = values.Skip(1).ToArray();

                //                 viewPlan = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                //                 viewPlan.Name = values[0];

                //             }
                //         }
                //         t.Commit();

                //     }
                #endregion



            }

            return Result.Succeeded;
        }
    }
}
