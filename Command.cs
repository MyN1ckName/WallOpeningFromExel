using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using System.Collections;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.Utility;

using ExeleCommand;

namespace RevitCommand
{
	[TransactionAttribute(TransactionMode.Manual)]
	[RegenerationAttribute(RegenerationOption.Manual)]

	public class Command : IExternalCommand
	{
		List<ElementId> _added_element_ids = new List<ElementId>();

		public Result Execute(ExternalCommandData commandData,
			ref string messege,
			ElementSet elements)
		{
			UIApplication uiApp = commandData.Application;
			Document doc = uiApp.ActiveUIDocument.Document;
			Application app = uiApp.Application;

			ExeleFile xsl = new ExeleFile();

			Reference pickedRef = null;
			Selection sel = uiApp.ActiveUIDocument.Selection;
			WallPickFilter selFiter = new WallPickFilter();
			pickedRef = sel.PickObject(ObjectType.Element, selFiter, "Выберите стену");

			Element elem = doc.GetElement(pickedRef.ElementId);

			OpeningWatcherUpdater _updater = new OpeningWatcherUpdater(app.ActiveAddInId);
			UpdaterRegistry.RegisterUpdater(_updater, true);
			UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(),
				new ElementCategoryFilter(BuiltInCategory.OST_Windows),
				Element.GetChangeTypeElementAddition());

			using (Transaction t = new Transaction(doc, "CreateWindows"))
			{
				t.Start("Create");
				CreateWindows(elem as Wall, app);
				t.Commit();

				if (_updater.GetElmtId != null)
				{
					SetParameters(doc, xsl, _updater.GetElmtId);
					xsl.CloseAndQuit();
				}
				else { xsl.CloseAndQuit(); }

			}

			UpdaterRegistry.UnregisterUpdater(_updater.GetUpdaterId());

			if (_updater.GetElmtId != null)
			{
				TaskDialog.Show("This message is command", _updater.GetElmtId.IntegerValue.ToString());
			}

			return Result.Succeeded;
		}

		private void SetParameters(Document doc, ExeleFile xsl, ElementId elmtid)
		{
			using (Transaction t = new Transaction(doc, "SetParameters"))
			{
				t.Start("SetParameters");
				Element elmt = doc.GetElement(elmtid);
				Parameter width = elmt.LookupParameter("Рзм.Ширина");
				Parameter height = elmt.LookupParameter("Рзм.Высота");
				//Parameter offsetLvl = elmt.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM); // Высота нижнего бруса

				width.SetValueString(xsl.CellsContent(2, xsl.ColumnNamber("WIDTH")));
				height.SetValueString(xsl.CellsContent(2, xsl.ColumnNamber("HEIGHT")));
				//height.Set(xsl.CellsContent(2, xsl.ColumnNamber("HEIGHT")));
				t.Commit();
			}
		}
		private void CreateWindows(Wall wall, Application app)
		{
			var locationCurve = (LocationCurve)wall.Location;

			var position = locationCurve.Curve.Evaluate(
			  0.5, true);

			Document document = wall.Document;

			Level level = (Level)document.GetElement(
			  wall.LevelId);

			FilteredElementCollector collector = new FilteredElementCollector(document);
			collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_Windows);
			FamilySymbol fs = collector.FirstOrDefault<Element>(
				s => s.Name.Equals("231_Отверстие прямоуг (Окно_Стена)")) as FamilySymbol;
			if (!fs.IsActive)
				fs.Activate();

			document.Create.NewFamilyInstance(position, fs,
				wall, level, StructuralType.NonStructural);
		}


		//Использовать эту перегрузку
		//NewFamilyInstance Method (XYZ, FamilySymbol, Element, Level, StructuralType)

		// Нужно добавить фильтр, вид должен быть планом!!!

		/*
		int colWidht = xsl.ColumnNamber("WIDTH");
		string content = xsl.CellsContent(2, colWidht);
		TaskDialog.Show("Show", content);								
		xsl.CloseAndQuit();
		*/
		private static Level GetActiveLevel(Document doc)
		{
			Autodesk.Revit.DB.View view = doc.ActiveView;
			FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
			ICollection<Element> lvlCollection = lvlCollector.OfClass(typeof(Level)).ToElements();

			ElementId levelId = null;
			Level level = null;
			foreach (Element l in lvlCollection)
			{
				Level lvl = l as Level;
				if (lvl.Name == view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsString())
				{
					levelId = lvl.Id;
					level = lvl;
				}
			}
			return level;
		}
	}

	class WallPickFilter : ISelectionFilter
	{
		public bool AllowElement(Element element)
		{
			return (element.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Walls));
		}
		public bool AllowReference(Reference reference, XYZ position)
		{
			return false;
		}
	}

	public class OpeningWatcherUpdater : IUpdater
	{
		static AddInId _appId;
		static UpdaterId _updaterId;
		ElementId eid;

		public OpeningWatcherUpdater(AddInId id)
		{
			_appId = id;
			_updaterId = new UpdaterId(_appId,
				new Guid("07175566-9bdf-4eed-b456-e3bbce5f9e22"));
		}

		public void Execute(UpdaterData data)
		{
			Document doc = data.GetDocument();
			Application app = doc.Application;

			if (data.GetAddedElementIds().Count == 1)
			{
				eid = data.GetAddedElementIds().First<ElementId>();
				//TaskDialog.Show("OpeningWatcherUpdater", eid.IntegerValue.ToString());
			}
		}

		public ElementId GetElmtId
		{
			get { return eid; }
		}

		public string GetAdditionalInformation()
		{
			return "This uhdater return Id added element";
		}

		public ChangePriority GetChangePriority()
		{
			return ChangePriority.DoorsOpeningsWindows;
		}

		public UpdaterId GetUpdaterId()
		{
			return _updaterId;
		}

		public string GetUpdaterName()
		{
			return "OpeningWatcherUpdater";
		}
	}
}