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

public class WorkFile
{
	OpenFileDialog ofd = new OpenFileDialog();

	public OpenFileDialog ShowFileDialog()
	{
		ofd.ShowDialog();
		return ofd;
	}
	
	public string Path
	{
		get	{return ShowFileDialog().FileName;}
	}		
}