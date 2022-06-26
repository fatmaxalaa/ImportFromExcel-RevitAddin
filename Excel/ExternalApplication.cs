using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace Excel
{
     public class ExternalApplication:IExternalApplication
    {
        Assembly assembly = Assembly.GetExecutingAssembly();
        public Result OnShutdown(UIControlledApplication application)
        {

            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                //Create Tab

                application.CreateRibbonTab("BIM III Plugins");


                //Create Excel Panel
                RibbonPanel ribbonPanel = application.CreateRibbonPanel("BIM III Plugins", "CSV-Addins");
                string path = Assembly.GetExecutingAssembly().Location;  //assemply path

                //Sheet Button
                #region Sheets btn

                PushButtonData btnData = new PushButtonData("sheetBtn", "Create Sheets", path, typeof(Class1).FullName);
                PushButton pushButton = ribbonPanel.AddItem(btnData) as PushButton;
                pushButton.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\csvsheet.png"));
                pushButton.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\sheet3d.png"));
                pushButton.ToolTip = "Import Sheets From .CSV Format and Excel SpreadSheets \n this Add-in Loads Sheet Number and Sheet Name of the Sheets";
               
                #endregion

                //Floor Plan Button
                #region floor plan btn

                //RibbonPanel ribbonPanel2 = application.CreateRibbonPanel("BIM III Plugins", "ExcelAddins2");

                PushButtonData btnplan = new PushButtonData("Floor Plan Button", "Create Floor Plans", path, typeof(FloorPlan).FullName);
                PushButton pushButtonplan = ribbonPanel.AddItem(btnplan) as PushButton;
                pushButtonplan.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\FloorplanCsv.png"));
                pushButtonplan.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\floorplan3d.png"));
                pushButtonplan.ToolTip = "Import Sheets From .CSV Format and Excel SpreadSheets \n this Add-in Loads  Name and Level of the Floor Plans";
                #endregion


                //Levels Button
                #region Levels btn

                PushButtonData levelBtnplan = new PushButtonData("Level Button", "Create Levels", path, typeof(CreateLevel).FullName);
                PushButton pushButtonLevl = ribbonPanel.AddItem(levelBtnplan) as PushButton;
                pushButtonLevl.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\levelCsv.png"));
                pushButtonLevl.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\levels3d.png"));
                pushButtonLevl.ToolTip = "Import Sheets From .CSV Format and Excel SpreadSheets \n this Add-in Loads Name and Elevation of the levels";
                #endregion


                //Ceiling Plan Button
                #region Ceiling btn

                PushButtonData ceilingBtnplan = new PushButtonData("Ceiling Plan Button", "Create Ceiling Plans", path, typeof(CreateCeilingPlan).FullName);
                PushButton pushButtonCeil = ribbonPanel.AddItem(ceilingBtnplan) as PushButton;
                pushButtonCeil.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\CeilingPlanCsv.png"));
                pushButtonCeil.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\ceiling 3d.png"));
                pushButtonCeil.ToolTip = "Import Sheets From .CSV Format and Excel SpreadSheets \n this Add-in Loads Name and Level of the Ceiling Plans";
                #endregion

                ///////////////////////////////////////////////////////////////////////////////////////////////////////
                RibbonPanel ribbonPanel2 = application.CreateRibbonPanel("BIM III Plugins", "Excel-Addins");
                string path2 = Assembly.GetExecutingAssembly().Location;  //assemply path

                //Sheet Button
                #region Sheets btn

                PushButtonData btnData2 = new PushButtonData("sheet excel Btn", "Create Excel Sheets", path2, typeof(SheetXlsx).FullName);
                PushButton pushButton2 = ribbonPanel2.AddItem(btnData2) as PushButton;
                pushButton2.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\Excel.png"));
                //pushButton2.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\FloorplanCsv.png"));

                #endregion

                //Floor Plan Button
                #region floor plan btn


                PushButtonData btnplan2 = new PushButtonData("Floor Plan Button", "Create Excel Floor Plans", path2, typeof(FloorPlanXlsx).FullName);
                PushButton pushButtonplan2 = ribbonPanel2.AddItem(btnplan2) as PushButton;
                pushButtonplan2.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\floorPlan.png"));
                //pushButtonplan2.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\FloorplanCsv.png"));

                #endregion


                //Levels Button
                #region Levels btn

                PushButtonData levelBtnplan2 = new PushButtonData("Level Button", "Create Excel Levels", path2, typeof(LevelXlsx).FullName);
                PushButton pushButtonLevl2 = ribbonPanel2.AddItem(levelBtnplan2) as PushButton;
                pushButtonLevl2.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\Levels2.png"));
                //pushButtonLevl2.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\FloorplanCsv.png"));

                #endregion


                //Ceiling Plan Button
                #region Ceiling btn

               PushButtonData ceilingBtnplan2 = new PushButtonData("Ceiling Plan Button", "Create Ceiling Plans", path2, typeof(CeilingPlanXlsx).FullName);
               PushButton pushButtonCeil2 = ribbonPanel2.AddItem(ceilingBtnplan2) as PushButton;
               pushButtonCeil2.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\ceilingExcel.png"));
               //pushButtonCeil2.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\Code\\ExcelV3 - xlsx\\Excel\\Excel\\images\\FloorplanCsv.png"));

                #endregion


                ///////////////////////////////////////////////////////////////////////////////////////////////////////
                //Create Tab

                application.CreateRibbonTab("WPF Plugins");


                //Create Wpf Panel
                RibbonPanel ribbonPanel3 = application.CreateRibbonPanel("WPF Plugins", " WPF Addin");
                string path3 = Assembly.GetExecutingAssembly().Location;  //assemply path

                PushButtonData btnData3 = new PushButtonData("Btn", "Excel WPF", path3, typeof(Class2).FullName);
                PushButton pushButton3 = ribbonPanel3.AddItem(btnData3) as PushButton;
                pushButton3.LargeImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\new code\\ExcelV3 - xlsx-all\\Excel\\Excel\\images\\WPF.png"));
                //pushButton3.ToolTipImage = new BitmapImage(new Uri("E:\\Courses\\ITI CEI 2021-2022\\BIM III\\sheet.png"));


            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }

            return Result.Succeeded;
        }


    }
}
