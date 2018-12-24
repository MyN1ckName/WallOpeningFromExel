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

public class ExeleFile
{
	
	Exele.Application xlApp = new Exele.Application();

	public Exele.Workbooks xlWorkbooks (Exele.Application xlApp)
	{
		Exele.Workbooks xlWorkbooks = xlApp.Workbooks;
		return xlWorkbooks;
	}
	public Exele.Workbook xlWorkbook (Exele.Workbooks xlWorkbooks, 
		string path)
	{
		Exele.Workbook xlWorkbook = xlWorkbooks.Open(path);
		return xlWorkbook;
	}
	public Exele._Worksheet xlWorksheet (Exele.Workbook xlWorkbook)
		{
			Exele._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
			return xlWorksheet;
		}
	public Exele.Range xlRange (Exele._Worksheet xlWorksheet)
		{
			Exele.Range xlRange = xlWorksheet.UsedRange;
			return xlRange;
		}
	public Exele.Application App
			{
				get { return xlApp; }
			}

	public int ColumnNamber(string name, Exele.Range xlRange)
	{
		int rowCount = xlRange.Rows.Count;
		rowCount = 1;
		int colCount = xlRange.Columns.Count;

		int colNamber = 0;
	   
		for (int i = 1; i <= colCount; i++)
		{
			if (xlRange.Cells[rowCount,i] == name &&
				xlRange.Cells[rowCount,i] != null &&
				xlRange.Cells[rowCount,i].Value2 != null)
			{
				colNamber = i;
			}				
		} 
		return colNamber;
	}

	public string CellsContent(int rowNamber, 
		int colNamber, Exele.Range xlRange)
	{
		string content =
			xlRange.Cells[rowNamber, colNamber].Value2.ToString();
		return content;
	}

	public void CloseAndQuit(
		Exele.Application xlApp,
		Exele.Workbooks xlWorbooks,
		Exele.Workbook xlWorkbook,
		Exele._Worksheet xlWorksheet,
		Exele.Range xlRange)
	{
		GC.Collect();
		GC.WaitForPendingFinalizers();

		Marshal.ReleaseComObject(xlRange);
		Marshal.ReleaseComObject(xlWorksheet);

		xlWorkbook.Close();
		Marshal.ReleaseComObject(xlWorkbook);

		Marshal.ReleaseComObject(xlWorbooks);

		xlApp.Quit();
		Marshal.ReleaseComObject(xlApp);
	}
}