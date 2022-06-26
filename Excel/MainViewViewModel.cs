using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;
using Autodesk.Revit.Attributes;


namespace Excel
{
   
    public class MainViewViewModel
    {
      

        private ExternalCommandData _commandData;
        private UIDocument _uidoc;
        private Document _doc;

        public DelegateCommand sheet { get; }

        public DelegateCommand ceilingPlan { get; }

        public DelegateCommand floorPlan { get; }

        public DelegateCommand level { get; }

        //public DelegateCommand SheetsXlsx { get; }


        public MainViewViewModel(ExternalCommandData commandData)
        {
            _commandData = commandData;
            _uidoc = _commandData.Application.ActiveUIDocument;
            _doc = _uidoc.Document;

            sheet = new DelegateCommand(sheets);
            ceilingPlan = new DelegateCommand(ceilingPlans);
            floorPlan = new DelegateCommand(floorPlans);
            level = new DelegateCommand(levels);
            //SheetsXlsx = new DelegateCommand(sheetXlsx);


        }

        private void sheets()
        {
            System.Windows.Forms.OpenFileDialog openDlg = new System.Windows.Forms.OpenFileDialog();

            string fileName;


            openDlg.Title = "Select a file";
            openDlg.DefaultExt = ".csv";
            openDlg.Filter = "CsvFile (.csv)|*.csv";
            DialogResult result = openDlg.ShowDialog();

            fileName = openDlg.FileName;
            StreamReader sr = new StreamReader(fileName);



            #region sheets
            ElementId titleBlocktytypeId = new FilteredElementCollector(_doc)
                                          .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                          .WhereElementIsElementType()
                                          .First().Id;

            string csvLine;

            Transaction t = new Transaction(_doc, "Create Sheets");
            t.Start();
            while ((csvLine = sr.ReadLine()) != null)
            {
                char[] separator = new char[] { ',' }; //to separete between (number and name and view columns)
                string[] values = csvLine.Split(separator, StringSplitOptions.None);


                if (values[0] != null && values[1] != null && values[2] != null)
                {
                    ViewSheet sheet = ViewSheet.Create(_doc, titleBlocktytypeId);
                    sheet.SheetNumber = values[0];
                    sheet.Name = values[1];
                    //sheet.ViewType = values[2];


                }
            }
            t.Commit();
            #endregion

            RaiseCloseRequest();



        }


        //private void sheetXlsx()
        //{
        //    string fileName;
        //    System.Windows.Forms.OpenFileDialog openDlg = new System.Windows.Forms.OpenFileDialog();

        //    openDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        //    //DialogResult result = openDlg.ShowDialog();

        //    openDlg.ShowDialog();
        //    fileName = openDlg.FileName;



        //    List<string> sheetNumber = new List<string>();
        //    List<string> sheetName = new List<string>();
        //    List<string> view = new List<string>();


        //    using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
        //    {

        //        ExcelWorksheet excelSheet = package.Workbook.Worksheets[4];  //last sheet in excel  "sheet"



        //        //column No. 1

        //        for (int c1 = 1; c1 < 1222; c1++)
        //        {
        //            var cellval = excelSheet.Cells[c1, 1].Value;  //[row,col]
        //            if (cellval == null)
        //            {
        //                break;
        //            }

        //            sheetNumber.Add(cellval.ToString());


        //        }


        //        //TaskDialog.Show("sheet", info);


        //        ////column No. 2

        //        for (int c2 = 1; c2 < 12222; c2++)
        //        {
        //            var cellval2 = excelSheet.Cells[c2, 2].Value;
        //            if (cellval2 == null)
        //            {
        //                break;
        //            }
        //            //col2list.Add(cellval2 + Environment.NewLine); //--------col2
        //            sheetName.Add(cellval2.ToString());

        //        }

        //        ////column No. 3

        //        for (int c3 = 1; c3 < 12222; c3++)
        //        {
        //            var cellval3 = excelSheet.Cells[c3, 3].Value;
        //            if (cellval3 == null)
        //            {
        //                break;
        //            }
        //            view.Add(cellval3 + Environment.NewLine); //--------col3

        //        }



        //        ElementId titleBlocktytypeId = new FilteredElementCollector(_doc)
        //                                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
        //                                   .WhereElementIsElementType()
        //                                   .First().Id;


        //        for (int i = 1; i < sheetNumber.Count; i++)
        //        {
        //            Transaction t = new Transaction(_doc, "Create Sheets");
        //            t.Start();


        //            if (sheetNumber != null && sheetName != null && view != null)
        //            {
        //                ViewSheet sheet = ViewSheet.Create(_doc, titleBlocktytypeId);
        //                sheet.SheetNumber = sheetNumber[i];
        //                sheet.Name = sheetName[i];

        //            }

        //            t.Commit();

        //        }






        //    }

        //}



        //private void sheets()
        //{
        //    System.Windows.Forms.OpenFileDialog openDlg = new System.Windows.Forms.OpenFileDialog();

        //    string fileName;


        //    openDlg.Title = "Select a file";
        //    DialogResult result = openDlg.ShowDialog();

        //    fileName = openDlg.FileName;
        //    StreamReader sr = new StreamReader(fileName);



        //    #region sheets
        //    ElementId titleBlocktytypeId = new FilteredElementCollector(_doc)
        //                                  .OfCategory(BuiltInCategory.OST_TitleBlocks)
        //                                  .WhereElementIsElementType()
        //                                  .First().Id;

        //    string csvLine;

        //    Transaction t = new Transaction(_doc, "Create Sheets");
        //    t.Start();
        //    while ((csvLine = sr.ReadLine()) != null)
        //    {
        //        char[] separator = new char[] { ',' }; //to separete between (number and name and view columns)
        //        string[] values = csvLine.Split(separator, StringSplitOptions.None);


        //        if (values[0] != null && values[1] != null && values[2] != null)
        //        {
        //            ViewSheet sheet = ViewSheet.Create(_doc, titleBlocktytypeId);
        //            sheet.SheetNumber = values[0];
        //            sheet.Name = values[1];
        //            //sheet.ViewType = values[2];


        //        }
        //    }
        //    t.Commit();
        //    #endregion

        //    RaiseCloseRequest();



        //}

        private void ceilingPlans()
        {


            System.Windows.Forms.OpenFileDialog openDlg = new System.Windows.Forms.OpenFileDialog(); //allow user to browse a new file and open it in revit

            string fileName;

            openDlg.Title = "Select a file";

            DialogResult result = openDlg.ShowDialog();

            if (result == DialogResult.OK)  //if user click ok 
            {
                fileName = openDlg.FileName;   //selected file name will store in filename variable
                StreamReader sr = new StreamReader(fileName);  //read the file


                Level level = new FilteredElementCollector(_doc)
                           .OfCategory(BuiltInCategory.OST_Levels)
                           .WhereElementIsNotElementType()
                           .Cast<Level>()
                           .First(e => e.Name == "Level 1");

                ViewFamilyType viewFamilyType = new FilteredElementCollector(_doc)
                            .OfClass(typeof(ViewFamilyType))
                            .WhereElementIsElementType()
                            .Cast<ViewFamilyType>()
                            .First(e => e.ViewFamily == ViewFamily.CeilingPlan);


                ViewPlan viewPlan = null;
                string csvLine;

                using (Transaction t = new Transaction(_doc, "Create Ceiling plan"))
                {

                    t.Start();
                    while ((csvLine = sr.ReadLine()) != null)
                    {
                        char[] separator = new char[] { ',' }; //to separete between
                        string[] values = csvLine.Split(separator, StringSplitOptions.None);


                        if (values[0] != null && values[1] != null)
                        {
                            //values = values.Skip(1).ToArray();

                            viewPlan = ViewPlan.Create(_doc, viewFamilyType.Id, level.Id);
                            viewPlan.Name = values[0];

                        }
                    }
                    t.Commit();

                }




            }
        }

        private void floorPlans()
        {
            string fileName;

            System.Windows.Forms.OpenFileDialog openDlg = new System.Windows.Forms.OpenFileDialog(); //allow user to browse a new file and open it in revit

            openDlg.Title = "Select a file";

            DialogResult result = openDlg.ShowDialog();


            if (result == DialogResult.OK)  //if user click ok 
            {
                fileName = openDlg.FileName;   //selected file name will store in filename variable
                StreamReader sr = new StreamReader(fileName);  //read the file



                Level level = new FilteredElementCollector(_doc)
                              .OfCategory(BuiltInCategory.OST_Levels)
                              .WhereElementIsNotElementType()
                              .Cast<Level>()
                              .First(e => e.Name == "Level 1");

                ViewFamilyType viewFamilyType = new FilteredElementCollector(_doc)
                            .OfClass(typeof(ViewFamilyType))
                            .WhereElementIsElementType()
                            .Cast<ViewFamilyType>()
                            .First(e => e.ViewFamily == ViewFamily.FloorPlan);


                ViewPlan viewPlan = null;
                string csvLine;

                using (Transaction t = new Transaction(_doc, "Create Floor plan"))
                {

                    t.Start();
                    while ((csvLine = sr.ReadLine()) != null)
                    {
                        char[] separator = new char[] { ',' }; //to separete between Columns
                        string[] values = csvLine.Split(separator, StringSplitOptions.None);


                        if (values[0] != null && values[1] != null)
                        {

                            viewPlan = ViewPlan.Create(_doc, viewFamilyType.Id, level.Id);
                            viewPlan.Name = values[0];

                        }
                    }
                    t.Commit();

                }

            }
        }

        private void levels()
        {
            string fileName;

            System.Windows.Forms.OpenFileDialog openDlg = new System.Windows.Forms.OpenFileDialog();
            //allow user to browse a new file and open it in revit

            openDlg.Title = "Select a file";

            DialogResult result = openDlg.ShowDialog();

            if (result == DialogResult.OK)  //if user click ok 
            {
                fileName = openDlg.FileName;   //selected file name will store in filename variable
                StreamReader sr = new StreamReader(fileName);  //read the file

                using (Transaction t = new Transaction(_doc, "Create Floor plan"))
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

                            level = Level.Create(_doc, elevation);
                            level.Name = values[0];


                        }
                    }
                    t.Commit();

                }

            }

        }


        public event EventHandler CloseRequest;

        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }
    }
}
