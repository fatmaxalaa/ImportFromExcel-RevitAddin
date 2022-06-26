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
    public class SheetXlsx : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            string fileName;

            OpenFileDialog openDlg = new OpenFileDialog(); //allow user to browse a new file and open it in revit

            openDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            //DialogResult result = openDlg.ShowDialog();

            openDlg.ShowDialog();
            fileName = openDlg.FileName;

            #region sheets

            List<string> sheetNumber = new List<string>();
            List<string> sheetName = new List<string>();
            List<string> view = new List<string>();


            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
            {



                ExcelWorksheet excelSheet = package.Workbook.Worksheets[4];  //last sheet in excel  "sheet"



                //column No. 1

                for (int c1 = 1; c1 < 1222; c1++)
                {
                    var cellval = excelSheet.Cells[c1, 1].Value;  //[row,col]
                    if (cellval == null)
                    {
                        break;
                    }

                    sheetNumber.Add(cellval.ToString());


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
                    sheetName.Add(cellval2.ToString());

                }

                ////column No. 3

                for (int c3 = 1; c3 < 12222; c3++)
                {
                    var cellval3 = excelSheet.Cells[c3, 3].Value;
                    if (cellval3 == null)
                    {
                        break;
                    }
                    view.Add(cellval3 + Environment.NewLine); //--------col3

                }



                ElementId titleBlocktytypeId = new FilteredElementCollector(doc)
                                           .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                           .WhereElementIsElementType()
                                           .First().Id;


                for (int i = 1; i < sheetNumber.Count; i++)
                {
                    Transaction t = new Transaction(doc, "Create Sheets");
                    t.Start();


                    if (sheetNumber != null && sheetName != null && view != null)
                    {
                        ViewSheet sheet = ViewSheet.Create(doc, titleBlocktytypeId);
                        sheet.SheetNumber = sheetNumber[i];
                        sheet.Name = sheetName[i];

                    }

                    t.Commit();

                }






            }
            #endregion








            return Result.Succeeded;
        }
    }
}

