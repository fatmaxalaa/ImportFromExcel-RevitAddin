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
    public class LevelXlsx : IExternalCommand
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
            List<string> elev = new List<string>();


            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
            {



                ExcelWorksheet excelSheet = package.Workbook.Worksheets[1];



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
                    elev.Add(cellval2.ToString());

                }



                for (int i = 1; i < sheetName.Count; i++)
                {
                    Transaction t = new Transaction(doc, "Create levels");
                    t.Start();
                    Level level = null;

                    if (sheetName != null && elev != null)
                    {
                        double elevation = double.Parse(elev[i]);

                        level = Level.Create(doc, elevation);
                        level.Name = sheetName[i];

                    }

                    t.Commit();

                }
            }




            return Result.Succeeded;

        }
    }
}
