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

		ExeleFile xlFale = new ExeleFile();
		OpenFile openFile = new OpenFile();

		
		var xlWorkbooks = xlFale.xlWorkbooks(xlFale.App);
		var xlWorkbook = xlFale.xlWorkbook(xlWorkbooks,
			openFile.Path);
		var xlWorksheet = xlFale.xlWorksheet(xlWorkbook);
		var xlRange = xlFale.xlRange(xlWorksheet);

		/*
				string st = xlFale.CellsContent(2,1, xlRange);
				TaskDialog.Show("show", st.ToString());
				xlFale.CloseAndQuit(xlFale.App, xlWorkbooks, xlWorkbook,
					xlWorksheet, xlRange);
		*/


		// NewFamilyInstance Method (XYZ, FamilySymbol, Element, Level, StructuralType)

		XYZ coords = new XYZ();
		Autodesk.Revit.DB.View view = doc.ActiveView; // Нужно добавить фильтр, вид должен быть планом!!!
		Level level = doc.GetElement(view.LevelId) as Level;

		FamilySymbol familySymbol = 
		
		
		//xlFale.ColumnNamber("WIDTH", xlRange);


		return Result.Succeeded;
	}
	public class OpenFile
	{
		OpenFileDialog ofd = new OpenFileDialog();

		public OpenFileDialog ShowFileDialog()
		{
			ofd.Filter = "Excel|*.xls";
			ofd.ShowDialog();
			return ofd;
		}  
		public string Path
		{
			get { return ShowFileDialog().FileName; }
			//get { return ofd.FileName; }
		}
	}
}