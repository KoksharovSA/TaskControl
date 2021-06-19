using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskСontrol
{
	internal static class ExcelData
	{
		internal static Dictionary<string, Collection<Detail>> details = new Dictionary<string, Collection<Detail>>();
		public static async void ExcelDataLoadAsync(string dir, int col) 
		{
			await Task.Run(() => ExcelDataLoad(dir, 4, new int[] { 1, 2, 12, 13, col }));
		}
		public static void ExcelDataLoad(string dir, int startRow, int[] column)
		{
			Collection<Detail> det = new Collection<Detail>();
			Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
			Workbook wb = excel.Workbooks.Open(dir);
			Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.ActiveSheet;

			int currentRow = startRow;
			bool isNotEmptyField = true;
			while (isNotEmptyField)
			{
				var b = excelSheet.Cells[currentRow, column[1]].Value;
				if (b != null
					&& Convert.ToString(b) != ""
					&& Convert.ToString(b) != " ")
				{
					Detail detail = new Detail((string)(excelSheet.Cells[currentRow, column[0]] as Microsoft.Office.Interop.Excel.Range)?.Value?.Trim(),
						Convert.ToString((excelSheet.Cells[currentRow, column[2]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
						Convert.ToString((excelSheet.Cells[currentRow, column[3]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
						Convert.ToString((excelSheet.Cells[currentRow, column[1]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
						Convert.ToString((excelSheet.Cells[currentRow, column[4]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
						new FileInfo(dir).Name);
					det.Add(detail);

					currentRow += 1;
				}
				else 
				{ 
					isNotEmptyField = false;					
				}				
			}
			details.Add(dir, det);
			wb.Close(0);
			excel.Quit();
		}      
	}
}
