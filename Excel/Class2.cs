using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Excel
{
    [Transaction(TransactionMode.Manual)]
    public class Class2 : IExternalCommand  
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            MainView window = new MainView(commandData);
            window.ShowDialog();
            return Result.Succeeded;
        }
    }
}
