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
    [Transaction(TransactionMode.Manual)]
    public class FloorPlan : IExternalCommand
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



                Level level = new FilteredElementCollector(doc)
                              .OfCategory(BuiltInCategory.OST_Levels)
                              .WhereElementIsNotElementType()
                              .Cast<Level>()
                              .First(e => e.Name == "Level 1");

                ViewFamilyType viewFamilyType = new FilteredElementCollector(doc)
                            .OfClass(typeof(ViewFamilyType))
                            .WhereElementIsElementType()
                            .Cast<ViewFamilyType>()
                            .First(e => e.ViewFamily == ViewFamily.FloorPlan);


                ViewPlan viewPlan = null;
                string csvLine;

                using (Transaction t = new Transaction(doc, "Create Floor plan"))
                {

                    t.Start();
                    while ((csvLine = sr.ReadLine()) != null)
                    {
                        char[] separator = new char[] { ',' }; //to separete between Columns
                        string[] values = csvLine.Split(separator, StringSplitOptions.None);


                        if (values[0] != null && values[1] != null)
                        {
                          
                            viewPlan = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                            viewPlan.Name = values[0];

                        }
                    }
                    t.Commit();

                }

            }



            return Result.Succeeded;
        }
    }
}
