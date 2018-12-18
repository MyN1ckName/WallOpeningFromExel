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

		OpenFile openFile = new OpenFile();
		openFile.ShowFileDialog();
		//string st = file.Path;
		//TaskDialog.Show("show", (openFile.Path).ToString());

		ExeleFile xlFale = new ExeleFile();
		var xlRange = xlFale.getExeleRange(openFile.Path);


		xlFale.ColumnNamber("WIDTH", xlRange);

		
		return Result.Succeeded;
	}
	public class OpenFile
	{
		OpenFileDialog ofd = new OpenFileDialog();

		public OpenFileDialog ShowFileDialog()
		{
			ofd.ShowDialog();
			return ofd;
		}

		public string Path
		{
			get { return ShowFileDialog().FileName; }
		}
	}
}