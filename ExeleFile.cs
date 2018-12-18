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
	public Exele.Range getExeleRange(string path)
	{
		Exele.Application xlApp = new Exele.Application();
		Exele.Workbooks xlWorkbooks = xlApp.Workbooks;
		Exele.Workbook xlWorkbook = xlWorkbooks.Open(@path);
		Exele._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
		Exele.Range xlRange = xlWorksheet.UsedRange;

		return xlRange;

		//int rowCount = xlRange.Rows.Count;
		//int colCount = xlRange.Columns.Count;
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
}