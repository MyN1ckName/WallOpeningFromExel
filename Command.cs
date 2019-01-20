using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows.Forms;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
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
		string st = (xlFale.CellsContent(2,4, xlRange)).ToString();
		TaskDialog.Show("show", st.ToString());
		xlFale.CloseAndQuit(xlFale.App, xlWorkbooks, xlWorkbook,
		xlWorksheet, xlRange);
		*/	

		//NewFamilyInstance Method (XYZ, FamilySymbol, Element, Level, StructuralType) 
				
		Autodesk.Revit.DB.View view = doc.ActiveView;						 // Нужно добавить фильтр, вид должен быть планом!!!
		FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
		ICollection<Element> lvlCollection = lvlCollector.OfClass(typeof(Level)).ToElements();

		ElementId levelId = null;
		Level level = null; 

		foreach(Element l in lvlCollection)
		{
			Level lvl = l as Level;
			if(lvl.Name == view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsString())
				{
				levelId = lvl.Id;
				level = lvl;
				} 
		}
		
		double dx = Convert.ToDouble(xlFale.CellsContent(2, 4, xlRange)) / 304.8;
		double dy = Convert.ToDouble(xlFale.CellsContent(2, 5, xlRange)) / 304.8;

		XYZ coords = new XYZ(dx, dy, level.Elevation); // В XYZ должен зайти doudle!!!

		//XYZ coords = new XYZ(xlFale.CellsContent(2, 4, xlRange) / 304.8, xlFale.CellsContent(2, 5, xlRange) / 304.8, level.Elevation);

		FilteredElementCollector collector = new FilteredElementCollector(doc);
		/*
		FamilySymbol familySymbol = collector.OfClass(typeof(FamilySymbol)).  // НЕ РАБОТАЕТ СО СТРОКИ 74
			OfCategory(BuiltInCategory.OST_Windows).
			OfClass(typeof(Element)).FirstOrDefault(e => e.Name.
			Equals("231_Проем прямоуг (Окно_Стена)")) as FamilySymbol;
		if (!familySymbol.IsActive)
		{ familySymbol.Activate(); }
		*/

		/////////////////

		collector.OfClass(typeof(FamilySymbol)).
			OfCategory(BuiltInCategory.OST_Windows);

		FamilySymbol familySymbol = collector.
			FirstOrDefault(e => e.Name.
			Equals("231_Проем прямоуг (Окно_Стена)")) as FamilySymbol; 

		if (!familySymbol.IsActive)
		{ familySymbol.Activate(); }

		///////////////

		Element el = new FilteredElementCollector(doc).
			OfCategory(BuiltInCategory.OST_Walls).ToElements() as Element;

		using (Transaction t = new Transaction(doc, "Create"))
		{
			t.Start("Start");
			try
			{
				Element e = doc.Create.NewFamilyInstance(coords, familySymbol, el,
					level, StructuralType.NonStructural);
				Parameter p1 = e.LookupParameter("Рзм.Ширина");
				Parameter p2 = e.LookupParameter("Рзм.Высота");
				Parameter p3 = e.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
				Parameter p4 = e.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);

				p1.Set(1550.0 / 304.8);
				p2.Set(650.0 / 304.8);
				p3.Set(levelId);
				p4.Set(0.0);

			}
			catch { }
			t.Commit();
		}


		/*
		int c = xlFale.ColumnNamber("WIDTH", xlRange);
		if(c != 0)
		{
			// Выполнять только если ColumnNamber != 0 
		}
		*/

		//xlFale.CloseAndQuit(xlFale.App, xlWorkbooks, xlWorkbook,
		//xlWorksheet, xlRange);
		return Result.Succeeded;
	}
	public class OpenFile
	{
		static OpenFileDialog ofd = new OpenFileDialog();

		static OpenFileDialog ShowFileDialog()
		{
			ofd.Filter = "Excel|*.xls";
			ofd.ShowDialog();
			return ofd;
		}  
		public string Path
		{
			get { return ShowFileDialog().FileName; }
		}
	}
}