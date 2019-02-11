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

namespace ExeleCommand
{
	public class ExeleFile
	{
		static OpenFile ofd = new OpenFile();

		static Exele.Application xlApp = new Exele.Application();
		static Exele.Workbooks xlWorkBooks = xlApp.Workbooks;
		static Exele.Workbook xlWorkBook = xlWorkBooks.Open(ofd.Path);
		static Exele._Worksheet xlWorkSheet = xlWorkBook.Sheets[1];
		static Exele.Range xlRange = xlWorkSheet.UsedRange;

		// метод возвращает номер столбца
		//с задонным содержанием первой строки name
		public int ColumnNamber(string name) 
		{ 
			int rowCount = xlRange.Rows.Count;
			rowCount = 1;
			int colCount = xlRange.Columns.Count;

			int colNamber = 0;

			for (int i = 1; i <= colCount; i++)
			{
				if (xlRange.Cells[rowCount, i].Value2.ToString() == name &&
					xlRange.Cells[rowCount, i] != null &&
					xlRange.Cells[rowCount, i].Value2 != null)
				{
					colNamber = i;
				}
			}
			return colNamber;
		}

		// метод возвращает содержимое ячейки
		public string CellsContent(int rowNamber,
			int colNamber)
		{
			string content =
				xlRange.Cells[rowNamber, colNamber].Value2.ToString();
			return content;
		}

		// метод закрывает файл и процессы
		public void CloseAndQuit()
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();

			Marshal.ReleaseComObject(xlRange);
			Marshal.ReleaseComObject(xlWorkSheet);

			xlWorkBook.Close();
			Marshal.ReleaseComObject(xlWorkBook);

			Marshal.ReleaseComObject(xlWorkBooks);

			xlApp.Quit();
			Marshal.ReleaseComObject(xlApp);   
		}
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