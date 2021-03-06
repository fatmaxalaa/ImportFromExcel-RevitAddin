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
    public class CreateLevel : IExternalCommand
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

                using (Transaction t = new Transaction(doc, "Create Floor plan"))
                {
                  Level level = null;
                  string csvLine;
                  t.Start();
                  while ((csvLine = sr.ReadLine()) != null)
                  {
                      char[] separator = new char[] { ',' }; //to separete between columns
                      string[] values = csvLine.Split(separator, StringSplitOptions.None);


                      if (values[0] != null && values[1] != null)
                      {
                          //values = values.Skip(1).ToArray();
                          double elevation = double.Parse(values[1]);

                          level = Level.Create(doc, elevation);
                          level.Name = values[0];

                      }
                  }
                  t.Commit();

            }

        }



            return Result.Succeeded;
           
        }
    }
}
