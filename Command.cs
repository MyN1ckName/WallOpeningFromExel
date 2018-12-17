using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows.Forms;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

[TransactionAttribute(TransactionMode.Manual)]
[RegenerationAttribute(RegenerationOption.Manual)]

public class Command : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData,
		ref string messege,
		ElementSet elements)
	{
		UIApplication uiApp = commandData.Application;
		Document doc = uiApp.ActiveUIDocument.Document;

		WorkFile file = new WorkFile();
		file.ShowFileDialog();
		//string st = file.Path;
		TaskDialog.Show("show", (file.Path).ToString());	

		return Result.Succeeded;
	}
}