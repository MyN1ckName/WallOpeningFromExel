using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Exele = Microsoft.Office.Interop.Excel;

using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows.Forms;

public class WorkExele
{
	public static void getExeleFile(string path)
	{
		Exele.Application xlApp = new Exele.Application();
		Exele.Workbook xlWorkbook = xlApp.Workbooks.Open(@path);
		Exele._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
		Exele.Range xlRange = xlWorksheet.UsedRange;
	}
}