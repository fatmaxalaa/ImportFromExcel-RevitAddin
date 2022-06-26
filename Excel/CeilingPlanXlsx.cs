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
    public class CeilingPlanXlsx : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            string fileName;

            OpenFileDialog openDlg = new OpenFileDialog(); //allow user to browse a new file and open it in revit

            openDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openDlg.ShowDialog();
            fileName = openDlg.FileName;


            List<string> sheetName = new List<string>();
            List<string> levelCol = new List<string>();



            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
            {



                ExcelWorksheet excelSheet = package.Workbook.Worksheets[3]; 



                //column No. 1

                for (int c1 = 1; c1 < 1222; c1++)
                {
                    var cellval = excelSheet.Cells[c1, 1].Value;  //[row,col]
                    if (cellval == null)
                    {
                        break;
                    }

                    sheetName.Add(cellval.ToString());


                }


                //TaskDialog.Show("sheet", info);


                ////column No. 2

                for (int c2 = 1; c2 < 12222; c2++)
                {
                    var cellval2 = excelSheet.Cells[c2, 2].Value;
                    if (cellval2 == null)
                    {
                        break;
                    }
                    //col2list.Add(cellval2 + Environment.NewLine); //--------col2
                    levelCol.Add(cellval2.ToString());

                }





                Level level = new FilteredElementCollector(doc)
                              .OfCategory(BuiltInCategory.OST_Levels)
                              .WhereElementIsNotElementType()
                              .Cast<Level>()
                              .First(e => e.Name == "Level 1");

                ViewFamilyType viewFamilyType = new FilteredElementCollector(doc)
                            .OfClass(typeof(ViewFamilyType))
                            .WhereElementIsElementType()
                            .Cast<ViewFamilyType>()
                            .First(e => e.ViewFamily == ViewFamily.CeilingPlan);


                ViewPlan viewPlan = null;


                for (int i = 1; i < sheetName.Count; i++)
                {
                    Transaction t = new Transaction(doc, "Create Sheets");
                    t.Start();


                    if (sheetName != null && levelCol != null)
                    {
                        viewPlan = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                        viewPlan.Name = sheetName[i];

                    }

                    t.Commit();

                }
            }



            return Result.Succeeded;
        }

    }
}
